"""
ΣΑΠ-ΦΘ
System Configuration
"""

from pathlib import Path

# Root directory
ROOT_DIR = Path(__file__).resolve().parent.parent

# Project folders
DATA_DIR = ROOT_DIR / "data"
EXCEL_DIR = ROOT_DIR / "excel"
REPORTS_DIR = ROOT_DIR / "reports"
DOCS_DIR = ROOT_DIR / "docs"

# Source files
GOOGLE_EXPORT = DATA_DIR / "google_forms_export.xlsx"

# Output files
MASTER_FILE = EXCEL_DIR / "ΣΑΠ-ΦΘ_MASTER.xlsx"

# Project information
PROJECT_NAME = "ΣΑΠ-ΦΘ"
PROJECT_FULL_NAME = "Σύστημα Αξιολόγησης και Ποιότητας Τμήματος Φυσικοθεραπείας"

ORGANIZATION = "Αυτοτελές Γραφείο Φυσικοθεραπείας"

HOSPITAL = "Εθνικό Κέντρο Αποκατάστασης (Ε.Κ.Α.)"

VERSION = "MASTER"

YEAR = 2026
