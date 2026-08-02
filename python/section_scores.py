"""
ΣΑΠ-ΦΘ
Section Scores
"""

import pandas as pd

from scoring_engine import build_scored_dataset
from metadata_engine import build_metadata


def section_means():

    df = build_scored_dataset()

    meta = build_metadata()

    results = []

    for section in sorted(meta["Section"].unique()):

        questions = meta.loc[
            (meta["Section"] == section) &
            (meta["Type"] == "LIKERT"),
            "Question"
        ].tolist()

        cols = [c for c in questions if c in df.columns]

        if not cols:
            continue

        value = df[cols].mean().mean()

        results.append({

            "Section": section,

            "Mean": round(value, 2)

        })

    return pd.DataFrame(results)


if __name__ == "__main__":

    print(section_means())
