"""
ΣΑΠ-ΦΘ
Statistics Engine
"""

import pandas as pd


LIKERT_MAP = {
    "Πάρα πολύ": 5,
    "Πολύ": 4,
    "Μέτρια": 3,
    "Λίγο": 2,
    "Καθόλου": 1,
}


def convert_likert(df: pd.DataFrame) -> pd.DataFrame:
    """
    Μετατρέπει τις λεκτικές απαντήσεις Likert σε αριθμητικές τιμές.
    """
    converted = df.copy()

    for col in converted.columns:
        converted[col] = converted[col].replace(LIKERT_MAP)

    return converted


def question_statistics(df: pd.DataFrame) -> pd.DataFrame:
    """
    Υπολογισμός βασικών στατιστικών ανά ερώτηση.
    """

    numeric = df.select_dtypes(include="number")

    stats = pd.DataFrame({
        "Mean": numeric.mean(),
        "Median": numeric.median(),
        "Std": numeric.std(),
        "Min": numeric.min(),
        "Max": numeric.max(),
        "Responses": numeric.count()
    })

    return stats.round(2)
