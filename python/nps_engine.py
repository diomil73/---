"""
Net Promoter Score
"""

from scoring_engine import build_scored_dataset


def calculate_nps(column_name):

    df = build_scored_dataset()

    if column_name not in df.columns:
        raise ValueError(f"Η στήλη '{column_name}' δεν βρέθηκε.")

    scores = df[column_name].dropna()

    promoters = (scores >= 9).sum()
    passives = ((scores >= 7) & (scores <= 8)).sum()
    detractors = (scores <= 6).sum()

    total = len(scores)

    if total == 0:
        return None

    nps = ((promoters - detractors) / total) * 100

    return round(nps, 1)
