"""
Export Metadata
"""

from metadata_engine import build_metadata
from config import DOCS_DIR

DOCS_DIR.mkdir(exist_ok=True)

meta = build_metadata()

meta.to_excel(
    DOCS_DIR / "Question_Metadata.xlsx",
    index=False,
)

print("Metadata exported.")
