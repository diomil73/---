"""
ΣΑΠ-ΦΘ
Import Google Forms data
"""

import pandas as pd

from config import GOOGLE_EXPORT


def load_data():

    df = pd.read_excel(GOOGLE_EXPORT)

    print(f"Εγγραφές : {len(df)}")
    print(f"Πεδία    : {len(df.columns)}")

    return df


if __name__ == "__main__":

    load_data()
