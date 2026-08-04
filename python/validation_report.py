"""
ΣΑΠ-ΦΘ
Data Validation Report
"""

from pathlib import Path

import pandas as pd

from import_data import load_data
from metadata_engine import build_metadata
from likert_engine import convert_value, is_non_applicable


PROJECT_ROOT = Path(__file__).resolve().parent.parent
OUTPUT_FILE = PROJECT_ROOT / "excel" / "VALIDATION_REPORT.xlsx"


def create_validation_report():
    df = load_data()
    meta = build_metadata()

    OUTPUT_FILE.parent.mkdir(parents=True, exist_ok=True)

    summary = pd.DataFrame(
        [
            {
                "Έλεγχος": "Εγγραφές",
                "Τιμή": len(df),
            },
            {
                "Έλεγχος": "Συνολικά πεδία",
                "Τιμή": len(df.columns),
            },
            {
                "Έλεγχος": "Ερωτήσεις LIKERT",
                "Τιμή": int((meta["Type"] == "LIKERT").sum()),
            },
            {
                "Έλεγχος": "Ερωτήσεις NPS",
                "Τιμή": int((meta["Type"] == "NPS").sum()),
            },
            {
                "Έλεγχος": "Ερωτήσεις TEXT",
                "Τιμή": int((meta["Type"] == "TEXT").sum()),
            },
            {
                "Έλεγχος": "Ερωτήσεις στα Λοιπά",
                "Τιμή": int((meta["Section"] == "Λοιπά").sum()),
            },
        ]
    )

    unmapped = meta.loc[
        meta["Section"] == "Λοιπά",
        [
            "Question",
            "CleanQuestion",
            "Type",
            "Section",
        ],
    ].copy()

    invalid_records = []

    likert_questions = meta.loc[
        meta["Type"] == "LIKERT",
        "Question",
    ].tolist()

    for question in likert_questions:
        if question not in df.columns:
            continue

        for row_number, original_value in df[question].items():
            if pd.isna(original_value):
                continue

            if is_non_applicable(original_value):
                continue

            converted_value = convert_value(original_value)

            if pd.isna(converted_value):
                invalid_records.append(
                    {
                        "Γραμμή": row_number + 2,
                        "Ερώτηση": question,
                        "Αρχική απάντηση": original_value,
                    }
                )

    invalid_likert = pd.DataFrame(
        invalid_records,
        columns=[
            "Γραμμή",
            "Ερώτηση",
            "Αρχική απάντηση",
        ],
    )

    section_rows = []

    sections = sorted(
        meta.loc[
            meta["Type"] == "LIKERT",
            "Section",
        ]
        .dropna()
        .unique()
    )

    for section in sections:
        questions = meta.loc[
            (meta["Section"] == section)
            & (meta["Type"] == "LIKERT"),
            "Question",
        ].tolist()

        existing_questions = [
            question
            for question in questions
            if question in df.columns
        ]

        total_cells = len(df) * len(existing_questions)

        if existing_questions:
            missing_cells = int(
                df[existing_questions].isna().sum().sum()
            )
        else:
            missing_cells = 0

        if total_cells:
            missing_percentage = round(
                (missing_cells / total_cells) * 100,
                2,
            )
        else:
            missing_percentage = 0

        section_rows.append(
            {
                "Ενότητα": section,
                "Ερωτήσεις LIKERT": len(existing_questions),
                "Συνολικές απαντήσεις": total_cells,
                "Κενές απαντήσεις": missing_cells,
                "Ποσοστό κενών": missing_percentage,
            }
        )

    section_coverage = pd.DataFrame(section_rows)

    with pd.ExcelWriter(
        OUTPUT_FILE,
        engine="openpyxl",
    ) as writer:
        summary.to_excel(
            writer,
            sheet_name="Summary",
            index=False,
        )

        meta.to_excel(
            writer,
            sheet_name="Metadata",
            index=False,
        )

        unmapped.to_excel(
            writer,
            sheet_name="Unmapped",
            index=False,
        )

        invalid_likert.to_excel(
            writer,
            sheet_name="Invalid_Likert",
            index=False,
        )

        section_coverage.to_excel(
            writer,
            sheet_name="Section_Coverage",
            index=False,
        )

    print(f"Validation report δημιουργήθηκε: {OUTPUT_FILE}")
    print(f"Μη αντιστοιχισμένες ερωτήσεις: {len(unmapped)}")
    print(
        "Μη αναγνωρίσιμες απαντήσεις Likert: "
        f"{len(invalid_likert)}"
    )


if __name__ == "__main__":
    create_validation_report()