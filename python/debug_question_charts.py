"""
Διαγνωστικός έλεγχος για το question_charts.py.
"""

from import_data import load_data
from metadata_engine import build_metadata


def main():
    raw_data = load_data()
    metadata = build_metadata()

    print()
    print("=" * 70)
    print("ΔΙΑΓΝΩΣΤΙΚΟΣ ΕΛΕΓΧΟΣ")
    print("=" * 70)

    print(
        "Στήλες metadata:",
        list(metadata.columns),
    )

    print(
        "Γραμμές metadata:",
        len(metadata),
    )

    if "Type" in metadata.columns:
        print(
            "Τύποι ερωτήσεων:",
            metadata["Type"]
            .value_counts(dropna=False)
            .to_dict(),
        )
    else:
        print(
            "Δεν υπάρχει στήλη Type στο metadata."
        )

    if "Question" not in metadata.columns:
        print(
            "Δεν υπάρχει στήλη Question στο metadata."
        )
        return

    exact_matches = sum(
        question in raw_data.columns
        for question in metadata["Question"]
    )

    print(
        "Ακριβείς αντιστοιχίσεις:",
        exact_matches,
    )

    print()
    print("ΠΡΩΤΕΣ 5 ΕΡΩΤΗΣΕΙΣ METADATA")

    for question in metadata["Question"].head(5):
        print(repr(question))

    print()
    print("ΠΡΩΤΕΣ 5 ΣΤΗΛΕΣ ΔΕΔΟΜΕΝΩΝ")

    for column in raw_data.columns[:5]:
        print(repr(column))

    print()
    print("ΠΡΩΤΕΣ 10 ΓΡΑΜΜΕΣ METADATA")

    available_columns = [
        column
        for column in [
            "Question",
            "Section",
            "Type",
        ]
        if column in metadata.columns
    ]

    print(
        metadata[
            available_columns
        ].head(10).to_string(index=False)
    )


if __name__ == "__main__":
    main()