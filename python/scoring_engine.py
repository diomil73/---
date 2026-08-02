"""
ΣΑΠ-ΦΘ
Scoring Engine
"""

import pandas as pd

from import_data import load_data
from metadata_engine import build_metadata
from likert_engine import convert


def build_scored_dataset():

    df = load_data()

    meta = build_metadata()

    scored = df.copy()

    likert_questions = meta.loc[
        meta["Type"] == "LIKERT",
        "Question"
    ]

    for q in likert_questions:

        if q in scored.columns:

            scored[q] = convert(scored[q])

    return scored


if __name__ == "__main__":

    data = build_scored_dataset()

    print(data.head())
