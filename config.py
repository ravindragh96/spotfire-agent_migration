
from pathlib import Path

# ==========================================
# MODULE 1: CONFIGURATION MANAGEMENT
# ==========================================

# 1. Main project directory
BASE_DIR = Path(__file__).resolve().parent

# 2. Excel migration tracker
EXCEL_FILE = BASE_DIR / "migration_tracker.xlsx"

# 3. Local project directories
CONFIG_DIR = BASE_DIR / "config"
LOG_DIR = BASE_DIR / "logs"
STAGING_DIR = BASE_DIR / "staging"

# 4. Migration destination on D drive
DESTINATION_ROOT = Path(r"D:\Spotfire_Migration")

# 5. Backup directories
BACKUP_ROOT = DESTINATION_ROOT / "Backups"

SOURCE_BACKUP_DIR = BACKUP_ROOT / "Source_Backups"
DESTINATION_BACKUP_DIR = BACKUP_ROOT / "Destination_Backups"

# 6. Monitoring configuration
CHECK_INTERVAL_SECONDS = 10

# 7. Supported migration operations
SUPPORTED_MIGRATION_TYPES = [
    "Copy",
    "Move"
]

# 8. Allowed migration statuses
MIGRATION_STATUSES = [
    "Pending",
    "In Progress",
    "Completed",
    "Failed",
    "Needs Review"
]


# ==========================================
# CREATE REQUIRED DIRECTORIES
# ==========================================

def create_required_directories():

    local_directories = [
        CONFIG_DIR,
        LOG_DIR,
        STAGING_DIR
    ]

    for folder in local_directories:
        folder.mkdir(parents=True, exist_ok=True)

    # Check whether D drive is available
    if not Path("D:/").exists():
        raise FileNotFoundError(
            "D drive is not available on this system."
        )

    # Create migration and backup directories
    required_directories = [
        DESTINATION_ROOT,
        BACKUP_ROOT,
        SOURCE_BACKUP_DIR,
        DESTINATION_BACKUP_DIR
    ]

    for folder in required_directories:
        folder.mkdir(parents=True, exist_ok=True)

    print("All required directories are ready.")


# ==========================================
# CONFIGURATION SUMMARY
# ==========================================

def display_configuration():

    print("\n========== MIGRATION CONFIGURATION ==========")

    print("Project Directory :", BASE_DIR)
    print("Excel Tracker     :", EXCEL_FILE)
    print("Staging Directory :", STAGING_DIR)
    print("Log Directory     :", LOG_DIR)

    print("Destination       :", DESTINATION_ROOT)
    print("Backup Root       :", BACKUP_ROOT)

    print("Source Backups    :", SOURCE_BACKUP_DIR)
    print("Destination Backup:", DESTINATION_BACKUP_DIR)

    print("Monitoring Interval:", CHECK_INTERVAL_SECONDS)
    print("Supported Types    :", SUPPORTED_MIGRATION_TYPES)

    print("=============================================\n")


# ==========================================
# MAIN TEST
# ==========================================

if __name__ == "__main__":

    create_required_directories()
    display_configuration()
