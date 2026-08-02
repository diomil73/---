"""
ΣΑΠ-ΦΘ
Question Classifier
"""

import pandas as pd

LIKERT_KEYWORDS = [
    "ικανοπ",
    "οργάν",
    "υποδομ",
    "κλιν",
    "θεραπε",
    "συνεργ",
    "υπεύθυ",
    "προϊστ",
]

TEXT_KEYWORDS = [
    "σχό",
    "παρατήρ",
    "πρότα",
]

NPS_KEYWORDS = [
    "προτείνατε",
    "recommend",
]


def detect_question_type(question: str):

    q = question.lower()

    if any(k in q for k in TEXT_KEYWORDS):
        return "TEXT"

    if any(k in q for k in NPS_KEYWORDS):
        return "NPS"

    if any(k in q for k in LIKERT_KEYWORDS):
        return "LIKERT"

    return "OTHER"


def classify_questions(columns):

    output = []

    for col in columns:

        output.append(
            {
                "Question": col,
                "Type": detect_question_type(col),
            }
        )

    return pd.DataFrame(output)
