import os
import glob
import pandas as pd
import re
import unicodedata

# === SET BASE DIRECTORY ===
project_dir = os.path.dirname(os.path.abspath(__file__))
base_dir = os.path.join(project_dir, "Master_database")
sf_archive_dir = os.path.join(project_dir, "SF_Archive")

# === FILE PATHS ===
names_file = os.path.join(base_dir, "Names.xlsx")
sflist_pattern = os.path.join(sf_archive_dir, "SFlist_*.xlsx")
history_pattern = os.path.join(base_dir, "Historical_data_*.xlsx")
helper_file = os.path.join(base_dir, "Helper.xlsx")

# === LOAD FILES ===
if not os.path.isdir(base_dir):
    raise FileNotFoundError(f"❌ Missing Master_database folder: {base_dir}")
if not os.path.isdir(sf_archive_dir):
    raise FileNotFoundError(f"❌ Missing SF_Archive folder: {sf_archive_dir}")

sflist_candidates = glob.glob(sflist_pattern)
history_candidates = glob.glob(history_pattern)

if not sflist_candidates:
    raise FileNotFoundError(f"❌ No SFlist files found with pattern: {sflist_pattern}")
if not history_candidates:
    raise FileNotFoundError(f"❌ No historical files found with pattern: {history_pattern}")

sflist_file = max(sflist_candidates, key=os.path.getmtime)
history_file = max(history_candidates, key=os.path.getmtime)

df_sflist = pd.read_excel(sflist_file, dtype=str)
df_history = pd.read_excel(history_file, dtype=str)
# Create a standardized working copy of historical data
df_hist_std = df_history.copy()


df_names_raw = pd.read_excel(names_file, sheet_name="Names")
df_names_map = df_names_raw.copy()

# === Read Final Field List ===
field_mapping_file = os.path.join(base_dir, "combined_field_mapping.xlsx")
final_fields_df = pd.read_excel(field_mapping_file, sheet_name="Field list")
print("🧾 Columns in Field list tab:", final_fields_df.columns.tolist())
final_fields = final_fields_df.iloc[:, 0].dropna().tolist()

# === Helper Function to Align Fields ===
def align_fields(df, label=""):
    current_cols = set(df.columns)
    missing_cols = [col for col in final_fields if col not in current_cols]
    for col in missing_cols:
        df[col] = pd.NA
    df = df[final_fields]
    print(f"✅ {label} aligned to {len(df.columns)} fields. Added {len(missing_cols)} columns.")
    return df

# === Create Standardized Copies ===
df_hist_std = df_history.copy()
df_sflist_std = df_sflist.copy()
df_names_std = df_names_raw.copy()
df_names_map = df_names_raw.copy()

# === Align All DataFrames ===
df_hist_std = align_fields(df_hist_std, label="Historical")
df_sflist_std = align_fields(df_sflist_std, label="SFlist")
df_names_std = align_fields(df_names_std, label="Names")




# === Step 1: Create Title--Project in SFlist ===
df_sflist_std["Project--Title"] = (
    df_sflist_std["Project job title"].fillna("") + "--" + df_sflist_std["Project"].fillna("")
).str.strip()

# === Step 2: Create CM--Project in SFlist ===
df_sflist_std["CM--Project"] = (
    df_sflist_std["Crew member id"].fillna("") + "--" + df_sflist_std["Project"].fillna("")
).str.strip()

# === Step 3: Load GCMID helper table ===
df_gcmid_helper = pd.read_excel(helper_file, sheet_name="GCMID", dtype=str)
df_gcmid_helper["CM-Job"] = df_gcmid_helper["CM-Job"].astype(str).str.strip()

# === Step 4: Map GCMID to SFlist ===
gcmid_map = df_gcmid_helper.set_index("CM-Job")["CM ID"].to_dict()
df_sflist_std["GCMID"] = df_sflist_std["CM--Project"].map(gcmid_map)

# === Step 5: Map General Title from Helper ===
# Load "Title conv" sheet from Helper
df_title_conv = pd.read_excel(helper_file, sheet_name="Title conv", dtype=str)

# Clean up and prepare the mapping
df_title_conv["Title-Project"] = df_title_conv["Title-Project"].astype(str).str.strip()
df_title_conv["General Title"] = df_title_conv["General Title"].astype(str).str.strip()
title_map = df_title_conv.set_index("Title-Project")["General Title"].to_dict()

# Apply mapping to SFlist
df_sflist_std["General Title"] = df_sflist_std["Project--Title"].map(title_map)
print(f"🧩 Mapped General Title in SFlist: {df_sflist_std['General Title'].notna().sum()} rows filled")


# === Prepare names lookup dictionary (GCMID → {Actual fields}) ===
fields_to_fill = ["Actual Name", "Actual Title", "Actual Phone", "Actual Email", "Note"]
df_names_map["CM ID"] = df_names_map["CM ID"].astype(str).str.strip()
name_lookup = df_names_map.set_index("CM ID")[fields_to_fill].to_dict(orient="index")

# === Ensure GCMID is string and trimmed ===
for df in [df_sflist_std, df_hist_std]:
    df["GCMID"] = df["GCMID"].astype(str).str.strip()

# === Fill from GCMID ===
def fill_from_lookup(df, label="Data"):
    for field in fields_to_fill:
        df[field] = df["GCMID"].map(lambda gcmid: name_lookup.get(gcmid, {}).get(field, ""))
    print(f"✅ Filled fields in {label}: {[field for field in fields_to_fill]}")

fill_from_lookup(df_sflist_std, "SFlist")
fill_from_lookup(df_hist_std, "Historical")

# Fill from original df_names_raw into df_names_std
df_names_std["Project"] = "Phone Book"
df_names_std["Surname"] = df_names_raw["Sure Name"]
df_names_std["Firstname"] = df_names_raw["First Name"]
df_names_std["Nickname"] = df_names_raw["Nick Name"]
df_names_std["Mobile number"] = df_names_raw["Actual Phone"]
df_names_std["Crew list name"] = df_names_raw["Actual Name"]
df_names_std["Crew email"] = df_names_raw["Actual Email"]
df_names_std["GCMID"] = df_names_raw["CM ID"]
df_names_std["General Title"] = df_names_raw["Actual Title"]

# === Load the 'General Title' helper table ===
df_general_title = pd.read_excel(helper_file, sheet_name="General Title", dtype=str)

# Rename for consistency
df_general_title = df_general_title.rename(columns={
    "Title": "General Title",
    "Department": "General Department"
})

print("📊 Columns in df_general_title:", df_general_title.columns.tolist())
print(df_general_title.head())


# === Create mapping dictionaries ===
dept_map = df_general_title.set_index("General Title")["General Department"].to_dict()
deptid_map = df_general_title.set_index("General Title")["Department ID"].to_dict()
titleid_map = df_general_title.set_index("General Title")["Title ID"].to_dict()

# === Function to apply mappings ===
def fill_general_fields(df, label=""):
    df["General Department"] = df["General Title"].map(dept_map)
    df["Department ID"] = df["General Title"].map(deptid_map)
    df["Title ID"] = df["General Title"].map(titleid_map)
    print(f"✅ Mapped General fields to {label}")
    return df

# === Apply to all 3 standardized DataFrames ===
df_hist_std = fill_general_fields(df_hist_std, "Historical")
df_sflist_std = fill_general_fields(df_sflist_std, "SFlist")
df_names_std = fill_general_fields(df_names_std, "Names")

# === Load FProjects tab from Helper.xlsx ===
df_fprojects = pd.read_excel(helper_file, sheet_name="FProjects", dtype=str)

# Clean up Project column and ensure no NaNs
df_fprojects["Project"] = df_fprojects["Project"].astype(str).str.strip()
df_fprojects = df_fprojects.dropna(subset=["Project"])

# Create mapping dictionaries
project_start_map = df_fprojects.set_index("Project")["Project start date"].to_dict()
project_end_map = df_fprojects.set_index("Project")["Project end date"].to_dict()

# Function to map project dates
def map_project_dates(df, label=""):
    df["Project start date"] = df["Project"].map(project_start_map)
    df["Project end date"] = df["Project"].map(project_end_map)
    print(f"📅 Project dates mapped for {label}")
    return df

# Apply to all three standardized DataFrames
df_hist_std = map_project_dates(df_hist_std, "Historical")
df_sflist_std = map_project_dates(df_sflist_std, "SFlist")
df_names_std = map_project_dates(df_names_std, "Names")

# Add Origin field to each dataframe
df_hist_std["Origin"] = "Historical"
df_sflist_std["Origin"] = "SFlist"
df_names_std["Origin"] = "Names"

print("🧷 Origin column set for all three DataFrames.")

def normalize_gcmid_column(df, column="GCMID"):
    if column in df.columns:
        def clean_gcmid_value(x):
            try:
                x = str(x).strip()
                if x == "" or x.lower() in ["nan", "none"]:
                    return ""
                return str(int(float(x)))
            except:
                return ""

        df[column] = df[column].apply(clean_gcmid_value)

    return df

df_hist_std = normalize_gcmid_column(df_hist_std)
df_sflist_std = normalize_gcmid_column(df_sflist_std)
df_names_std = normalize_gcmid_column(df_names_std)


# Combine all three DataFrames
df_combined = pd.concat([df_hist_std, df_sflist_std, df_names_std], ignore_index=True)


# === Export Final Combined File ===
output_path = os.path.join(base_dir, "Combined_All_CrewData.xlsx")
with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
    df_combined.to_excel(writer, sheet_name="Combined", index=False)
print(f"✅ Final file saved to: {output_path}")
