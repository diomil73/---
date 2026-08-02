"""
Descriptive Statistics
"""

from scoring_engine import build_scored_dataset


def describe():

    df = build_scored_dataset()

    numeric = df.select_dtypes("number")

    return numeric.describe().T
