@'
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

    for section in sorted(meta["Section"].dropna().unique()):
        questions = meta.loc[
            (meta["Section"] == section)
            & (meta["Type"] == "LIKERT"),
            "Question",
        ].tolist()

        cols = [column for column in questions if column in df.columns]

        if not cols:
            continue

        numeric_values = df[cols].apply(
            lambda column: pd.to_numeric(column, errors="coerce")
        )

        value = numeric_values.stack().mean()

        results.append({
            "Section": section,
            "Mean": round(float(value), 2) if pd.notna(value) else None,
        })

    return pd.DataFrame(results)
'@ | Set-Content -Encoding utf8 python\section_scores.py
