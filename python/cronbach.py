"""
Cronbach Alpha
"""

import numpy as np


def cronbach_alpha(df):

    df = df.dropna(axis=1, how="all")

    k = df.shape[1]

    if k < 2:
        return None

    item_var = df.var(axis=0, ddof=1)

    total_var = df.sum(axis=1).var(ddof=1)

    alpha = (k / (k - 1)) * (1 - item_var.sum() / total_var)

    return round(alpha, 3)
