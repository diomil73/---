"""
ΣΑΠ-ΦΘ
Question Analytics Report
"""

from pathlib import Path

import pandas as pd

from import_data import load_data
from likert_engine import convert_value, is_non_applicable
from metadata_engine import build_metadata
from scoring_engine import build_scored_dataset


PROJECT_ROOT = Path(__file__).resolve().parent.parent

OUTPUT_FILE = (
    PROJECT_ROOT
    / "excel"
    / "QUESTION_ANALYTICS.xlsx"
)


SECTION_ORDER = [
    "Α",
    "Β",
    "Γ",
    "Δ",
    "Ε",
    "ΣΤ",
    "Ζ",
    "Η",
    "Θ",
]


SECTION_TITLES = {
    "Α": "Υποδοχή και ενημέρωση",
    "Β": "Οργάνωση πρακτικής άσκησης",
    "Γ": "Κλινική εκπαίδευση και εμπειρία",
    "Δ": "Υποδομές και εξοπλισμός",
    "Ε": "Συνεργασία με επαγγελματίες υγείας",
    "ΣΤ": "Καθοδήγηση και υποστήριξη",
    "Ζ": "Ηγεσία και επαγγελματισμός",
    "Η": "Κλίμα συνεργασίας και ασφάλειας",
    "Θ": "Χώροι και συνθήκες πρακτικής",
}


def clean_question(question):
    """
    Αφαιρεί περιττά κενά και αγκύλες
    από το όνομα της ερώτησης.
    """

    return str(question).strip().strip("[]").strip()


def section_sort_value(section):
    """
    Επιστρέφει τη σωστή σειρά εμφάνισης
    της ενότητας.
    """

    if section in SECTION_ORDER:
        return SECTION_ORDER.index(section)

    return len(SECTION_ORDER)


def calculate_question_statistics(
    raw_data,
    scored_data,
    metadata,
):
    """
    Υπολογίζει αναλυτικά στατιστικά
    για κάθε ερώτηση Likert.
    """

    records = []

    likert_metadata = metadata.loc[
        metadata["Type"] == "LIKERT"
    ].copy()

    for _, metadata_row in likert_metadata.iterrows():
        question = metadata_row["Question"]
        section = metadata_row["Section"]

        if question not in raw_data.columns:
            continue

        raw_series = raw_data[question]

        if question in scored_data.columns:
            scored_series = pd.to_numeric(
                scored_data[question],
                errors="coerce",
            )
        else:
            scored_series = raw_series.apply(
                convert_value
            )

            scored_series = pd.to_numeric(
                scored_series,
                errors="coerce",
            )

        total_responses = len(raw_series)

        answered_responses = int(
            raw_series.notna().sum()
        )

        non_applicable_count = int(
            raw_series.apply(
                lambda value: (
                    False
                    if pd.isna(value)
                    else is_non_applicable(value)
                )
            ).sum()
        )

        valid_scores = scored_series.dropna()

        valid_count = int(valid_scores.count())

        missing_count = int(
            total_responses
            - answered_responses
        )

        invalid_count = int(
            answered_responses
            - non_applicable_count
            - valid_count
        )

        mean_value = (
            round(float(valid_scores.mean()), 2)
            if valid_count
            else None
        )

        median_value = (
            round(float(valid_scores.median()), 2)
            if valid_count
            else None
        )

        standard_deviation = (
            round(float(valid_scores.std(ddof=1)), 2)
            if valid_count > 1
            else 0
        )

        minimum_value = (
            float(valid_scores.min())
            if valid_count
            else None
        )

        maximum_value = (
            float(valid_scores.max())
            if valid_count
            else None
        )

        positive_count = int(
            valid_scores.isin([4, 5]).sum()
        )

        neutral_count = int(
            (valid_scores == 3).sum()
        )

        negative_count = int(
            valid_scores.isin([1, 2]).sum()
        )

        positive_percentage = (
            round(
                positive_count
                / valid_count
                * 100,
                2,
            )
            if valid_count
            else 0
        )

        neutral_percentage = (
            round(
                neutral_count
                / valid_count
                * 100,
                2,
            )
            if valid_count
            else 0
        )

        negative_percentage = (
            round(
                negative_count
                / valid_count
                * 100,
                2,
            )
            if valid_count
            else 0
        )

        participation_percentage = (
            round(
                valid_count
                / total_responses
                * 100,
                2,
            )
            if total_responses
            else 0
        )

        records.append(
            {
                "Ενότητα": section,
                "Περιγραφή ενότητας": (
                    SECTION_TITLES.get(
                        section,
                        "",
                    )
                ),
                "Ερώτηση": clean_question(question),
                "Σύνολο φοιτητών": total_responses,
                "Έγκυρες βαθμολογίες": valid_count,
                "Δεν συμμετείχαν": non_applicable_count,
                "Κενές απαντήσεις": missing_count,
                "Μη αναγνωρίσιμες απαντήσεις": invalid_count,
                "Ποσοστό συμμετοχής": participation_percentage,
                "Μέσος όρος": mean_value,
                "Διάμεσος": median_value,
                "Τυπική απόκλιση": standard_deviation,
                "Ελάχιστο": minimum_value,
                "Μέγιστο": maximum_value,
                "Θετικές απαντήσεις": positive_count,
                "Ουδέτερες απαντήσεις": neutral_count,
                "Αρνητικές απαντήσεις": negative_count,
                "Θετικές απαντήσεις %": positive_percentage,
                "Ουδέτερες απαντήσεις %": neutral_percentage,
                "Αρνητικές απαντήσεις %": negative_percentage,
                "_section_order": section_sort_value(section),
            }
        )

    results = pd.DataFrame(records)

    if results.empty:
        return results

    results = results.sort_values(
        by=[
            "_section_order",
            "Ερώτηση",
        ]
    ).reset_index(drop=True)

    return results.drop(
        columns=["_section_order"]
    )


def create_section_summary(question_statistics):
    """
    Δημιουργεί συγκεντρωτικά αποτελέσματα
    ανά ενότητα.
    """

    if question_statistics.empty:
        return pd.DataFrame()

    summary = (
        question_statistics
        .groupby(
            [
                "Ενότητα",
                "Περιγραφή ενότητας",
            ],
            as_index=False,
        )
        .agg(
            {
                "Ερώτηση": "count",
                "Έγκυρες βαθμολογίες": "sum",
                "Δεν συμμετείχαν": "sum",
                "Κενές απαντήσεις": "sum",
                "Μέσος όρος": "mean",
                "Θετικές απαντήσεις": "sum",
                "Ουδέτερες απαντήσεις": "sum",
                "Αρνητικές απαντήσεις": "sum",
            }
        )
    )

    summary = summary.rename(
        columns={
            "Ερώτηση": "Αριθμός ερωτήσεων",
        }
    )

    summary["Μέσος όρος"] = (
        summary["Μέσος όρος"]
        .round(2)
    )

    total_valid = (
        summary["Θετικές απαντήσεις"]
        + summary["Ουδέτερες απαντήσεις"]
        + summary["Αρνητικές απαντήσεις"]
    )

    summary["Θετικές απαντήσεις %"] = (
        summary["Θετικές απαντήσεις"]
        .div(total_valid.where(total_valid != 0))
        .mul(100)
        .fillna(0)
        .round(2)
    )

    summary["Ουδέτερες απαντήσεις %"] = (
        summary["Ουδέτερες απαντήσεις"]
        .div(total_valid.where(total_valid != 0))
        .mul(100)
        .fillna(0)
        .round(2)
    )

    summary["Αρνητικές απαντήσεις %"] = (
        summary["Αρνητικές απαντήσεις"]
        .div(total_valid.where(total_valid != 0))
        .mul(100)
        .fillna(0)
        .round(2)
    )

    summary["_section_order"] = (
        summary["Ενότητα"]
        .apply(section_sort_value)
    )

    summary = summary.sort_values(
        "_section_order"
    ).drop(
        columns=["_section_order"]
    )

    return summary.reset_index(drop=True)


def create_priority_tables(question_statistics):
    """
    Δημιουργεί πίνακες με τις καλύτερες
    και τις χαμηλότερες ερωτήσεις.
    """

    if question_statistics.empty:
        empty = pd.DataFrame()

        return empty, empty

    valid_questions = question_statistics.loc[
        question_statistics["Μέσος όρος"].notna()
    ].copy()

    strongest = (
        valid_questions
        .sort_values(
            by=[
                "Μέσος όρος",
                "Θετικές απαντήσεις %",
            ],
            ascending=[
                False,
                False,
            ],
        )
        .head(10)
        .reset_index(drop=True)
    )

    improvement = (
        valid_questions
        .sort_values(
            by=[
                "Μέσος όρος",
                "Αρνητικές απαντήσεις %",
            ],
            ascending=[
                True,
                False,
            ],
        )
        .head(10)
        .reset_index(drop=True)
    )

    selected_columns = [
        "Ενότητα",
        "Ερώτηση",
        "Έγκυρες βαθμολογίες",
        "Μέσος όρος",
        "Θετικές απαντήσεις %",
        "Αρνητικές απαντήσεις %",
        "Δεν συμμετείχαν",
    ]

    return (
        strongest[selected_columns],
        improvement[selected_columns],
    )


def create_question_analytics():
    """
    Δημιουργεί το αρχείο QUESTION_ANALYTICS.xlsx.
    """

    raw_data = load_data()

    scored_data = build_scored_dataset()

    metadata = build_metadata()

    question_statistics = (
        calculate_question_statistics(
            raw_data=raw_data,
            scored_data=scored_data,
            metadata=metadata,
        )
    )

    section_summary = create_section_summary(
        question_statistics
    )

    strongest, improvement = (
        create_priority_tables(
            question_statistics
        )
    )

    OUTPUT_FILE.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    with pd.ExcelWriter(
        OUTPUT_FILE,
        engine="openpyxl",
    ) as writer:
        question_statistics.to_excel(
            writer,
            sheet_name="Question_Statistics",
            index=False,
        )

        section_summary.to_excel(
            writer,
            sheet_name="Section_Summary",
            index=False,
        )

        strongest.to_excel(
            writer,
            sheet_name="Strongest_Points",
            index=False,
        )

        improvement.to_excel(
            writer,
            sheet_name="Improvement_Priorities",
            index=False,
        )

    print(
        f"Question analytics δημιουργήθηκε: {OUTPUT_FILE}"
    )

    print(
        "Ερωτήσεις Likert που αναλύθηκαν: "
        f"{len(question_statistics)}"
    )

    print(
        "Ενότητες που αναλύθηκαν: "
        f"{len(section_summary)}"
    )


if __name__ == "__main__":
    create_question_analytics()