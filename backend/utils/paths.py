from pathlib import Path

BASE_DIR = Path(__file__).resolve().parents[2]
MASTER_DIR = BASE_DIR / "Master_database"
SF_ARCHIVE_DIR = BASE_DIR / "SF_Archive"
CONFIG_DATA_DIR = BASE_DIR / "backend" / "config_data"
NEW_MASTER_DIR = BASE_DIR / "New_Master_Database"
GCMID_MAP_FILE = NEW_MASTER_DIR / "GCMID_Map.xlsx"
CREW_REGISTRY_FILE = NEW_MASTER_DIR / "CrewRegistry.xlsx"
TITLEMAP_FILE = NEW_MASTER_DIR / "TitleMap.xlsx"
PROJECTS_JSON = NEW_MASTER_DIR / "projects.json"
