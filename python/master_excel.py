"""
ΣΑΠ-ΦΘ
MASTER Excel Builder
"""

from openpyxl import Workbook
from openpyxl.styles import Font

from overall_iqi import calculate_iqi
from section_scores import section_means
from config import MASTER_FILE


def create_master():

    wb = Workbook()

    ws = wb.active

    ws.title = "Dashboard"

    ws["A1"] = "ΣΑΠ-ΦΘ"

    ws["A2"] = "Σύστημα Αξιολόγησης και Ποιότητας"

    ws["A1"].font = Font(size=18, bold=True)

    ws["A4"] = "IQI"

    ws["B4"] = calculate_iqi()

    ws["A6"] = "Μέσοι Όροι Ενοτήτων"

    scores = section_means()

    row = 8

    for _, r in scores.iterrows():

        ws.cell(row=row, column=1).value = r["Section"]

        ws.cell(row=row, column=2).value = r["Mean"]

        row += 1

    MASTER_FILE.parent.mkdir(exist_ok=True)

    wb.save(MASTER_FILE)

    print("MASTER.xlsx δημιουργήθηκε.")
