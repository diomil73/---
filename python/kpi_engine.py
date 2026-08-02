"""
KPI Engine
"""

import pandas as pd


def overall_mean(df):

    numeric = df.select_dtypes("number")

    return numeric.mean().mean()


def response_count(df):

    return len(df)


def completion_rate(df):

    total = df.size

    missing = df.isna().sum().sum()

    return (total - missing) / total * 100
