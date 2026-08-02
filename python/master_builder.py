"""
MASTER Builder
"""

from import_data import load_data
from question_classifier import classify_questions


def build():

    df = load_data()

    metadata = classify_questions(df.columns)

    print(metadata)

    return metadata


if __name__ == "__main__":

    build()
