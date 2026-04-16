import os
import re
import pandas as pd
import unicodedata

# === BASE DIRECTORY SETUP ===
base_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Master_database")

if not os.path.isdir(base_dir):
    raise FileNotFoundError(f"❌ Missing Master_database folder: {base_dir}")

combined_path = os.path.join(base_dir, "Combined_All_CrewData.xlsx")

# === Nickname dictionary ===
nickname_map = {
    "gabi": "gabriella", "zsuzsa": "zsuzsanna", "zsuzsi": "zsuzsanna", "gergo": "gergely",
    "kati": "katalin", "erzsi": "erzsebet", "bobe": "erzsebet", "bori": "borbala",
    "dani": "daniel", "moni": "monika", "zoli": "zoltan", "niki": "nikoletta",
    "pisti": "istvan", "magdi": "magdolna", "jr": "junior", "jrxx": "junior",
    "orsi": "orsolya", "ricsi": "richard", "gyuri": "gyorgy"
}

# === Helpers ===
def strip_accents(text):
    return ''.join(c for c in unicodedata.normalize('NFD', str(text)) if unicodedata.category(c) != 'Mn')

def clean_token(text):
    if not isinstance(text, str):
        return ""
    text = strip_accents(text.lower())
    text = re.sub(r"[\"'’‘“”`,]", "", text)
    text = text.replace("-", " ")
    text = re.sub(r"[().]", "", text)
    return text.strip()

def tokenize_name(name):
    tokens = clean_token(name).split()
    result = set()
    for token in tokens:
        if len(token) >= 3 and token != "né":
            result.add(token)
            if token in nickname_map:
                result.add(nickname_map[token])
    return sorted(result)

def format_phone(phone):
    if pd.isna(phone):
        return ""
    phone = re.sub(r"\D", "", str(phone))
    if len(phone) < 8:
        return ""  # too short
    if phone.startswith("36") or phone.startswith("00"):
        return phone
    if phone.startswith("06"):
        phone = "36" + phone[2:]
    elif phone.startswith("6"):
        phone = "36" + phone[1:]
    elif len(phone) == 9:
        phone = "36" + phone
    return phone

# === Load Combined and Names sheets ===
df_combined = pd.read_excel(combined_path, sheet_name="Combined", dtype=str)
names_file = os.path.join(base_dir, "Names.xlsx")
df_names = pd.read_excel(names_file, sheet_name="Names", dtype=str)


# === Tokenized Names ===
df_tokens_raw = df_combined[["GCMID", "Crew list name", "Origin"]].dropna(subset=["GCMID", "Crew list name"])
df_tokens_raw["GCMID"] = df_tokens_raw["GCMID"].str.strip()
df_tokens_raw["Crew list name"] = df_tokens_raw["Crew list name"].str.strip()
df_tokens_raw["Origin"] = df_tokens_raw["Origin"].fillna("")
df_tokens_raw["origin_rank"] = df_tokens_raw["Origin"].apply(lambda x: 0 if x == "Names" else 1)
df_tokens_raw = df_tokens_raw.sort_values(by=["GCMID", "origin_rank"])

tokens_set = set()
for _, row in df_tokens_raw.iterrows():
    gcmid = row["GCMID"]
    name = row["Crew list name"]
    for token in tokenize_name(name):
        tokens_set.add((gcmid, token))

df_tokenized_names = pd.DataFrame(sorted(tokens_set), columns=["GCMID", "Token"])

# === Phones ===
df_phones_raw = df_combined[["GCMID", "Mobile number"]].dropna(subset=["GCMID", "Mobile number"]).copy()
df_phones_raw["Original Phone"] = df_phones_raw["Mobile number"]
df_phones_raw["Phone"] = df_phones_raw["Mobile number"].apply(format_phone)
df_phones = df_phones_raw[["GCMID", "Phone", "Original Phone"]].drop_duplicates().sort_values(by="GCMID")

# === Emails ===
df_emails = df_combined[["GCMID", "Crew email"]].dropna(subset=["GCMID", "Crew email"]).copy()
df_emails["Crew email"] = df_emails["Crew email"].astype(str).str.strip()
df_emails = df_emails[df_emails["Crew email"].str.contains("@", na=False)]
df_emails = df_emails.rename(columns={"Crew email": "Email"})
df_emails = df_emails[["GCMID", "Email"]].drop_duplicates().sort_values(by="GCMID")


# === Actual Details ===
df_names = df_names[["CM ID", "Sure Name", "First Name", "Actual Title", "Actual Phone", "Actual Email"]].dropna(subset=["CM ID"])
df_names["CM ID"] = df_names["CM ID"].astype(str).str.strip()
df_names["Actual Name"] = df_names["Sure Name"].fillna("") + " " + df_names["First Name"].fillna("")
df_names["Actual Name"] = df_names["Actual Name"].str.strip()
df_actual = df_names[["CM ID", "Actual Name", "Actual Title", "Actual Phone", "Actual Email"]].drop_duplicates().sort_values(by="CM ID")

# === Names tab (GCMID + Crew list name) ===
df_names_tab = df_combined[["GCMID", "Crew list name"]].dropna().drop_duplicates().sort_values(by=["GCMID", "Crew list name"])

# === General Departments ===
df_general_dept = df_combined[["GCMID", "General Department"]].dropna(subset=["GCMID", "General Department"])
df_general_dept = df_general_dept.drop_duplicates().sort_values(by="GCMID")


# === Save All ===
with pd.ExcelWriter(combined_path, mode="a", engine="openpyxl", if_sheet_exists="replace") as writer:
    df_tokenized_names.to_excel(writer, sheet_name="Tokenized Names", index=False)
    df_phones.to_excel(writer, sheet_name="Phones", index=False)
    df_emails.to_excel(writer, sheet_name="Emails", index=False)
    df_actual.to_excel(writer, sheet_name="Actual Details", index=False)
    df_names_tab.to_excel(writer, sheet_name="Names", index=False)
    df_general_dept.to_excel(writer, sheet_name="General Departments", index=False)



print("✅ All tabs saved: Tokenized Names, Phones, Emails, Actual Details, Names, General Departments.")
