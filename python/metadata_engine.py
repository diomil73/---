"""
ΣΑΠ-ΦΘ
Metadata Engine
"""

from question_classifier import classify_questions
from import_data import load_data


SECTION_RULES = {

    "Δημογραφικά": [
        "ηλικ",
        "φύλο",
        "εξάμη",
        "πανεπιστ",
    ],

    "Β": [
        "οργάν",
        "πρόγραμμα",
    ],

    "Γ": [
        "κλιν",
        "θεραπε",
        "κρουσ",
        "υδροθεραπε",
        "ρομποτ",
        "ηλεκτροθεραπε",
        "laser",
    ],

    "Δ": [
        "υποδομ",
        "εξοπλισ",
        "χώρ",
    ],

    "Ε": [
        "ιατρ",
        "νοσηλευ",
        "συνεργ",
    ],

    "ΣΤ": [
        "υπεύθυ",
    ],

    "Ζ": [
        "προϊστ",
    ],

    "Η": [
        "γενικ",
        "ικανοπ",
    ],

    "Ι": [
        "σχό",
        "παρατήρ",
        "πρότα",
    ]
}


def detect_section(question):

    q = question.lower()

    for section, words in SECTION_RULES.items():

        for word in words:

            if word in q:
                return section

    return "Λοιπά"


def build_metadata():

    df = load_data()

    metadata = classify_questions(df.columns)

    metadata["Section"] = metadata["Question"].apply(detect_section)

    metadata["Scored"] = metadata["Type"].isin(
        [
            "LIKERT",
            "NPS",
        ]
    )

    return metadata


if __name__ == "__main__":

    meta = build_metadata()

    print(meta)
