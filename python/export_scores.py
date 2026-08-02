"""
Export Scores
"""

from config import REPORTS_DIR
from section_scores import section_means

REPORTS_DIR.mkdir(exist_ok=True)

section_means().to_excel(

    REPORTS_DIR / "Section_Scores.xlsx",

    index=False

)

print("Export completed.")
