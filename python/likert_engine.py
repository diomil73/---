"""
Likert Engine
"""

LIKERT = {
    "Πάρα πολύ": 5,
    "Πολύ": 4,
    "Μέτρια": 3,
    "Λίγο": 2,
    "Καθόλου": 1,
}


def convert(series):

    return series.replace(LIKERT)
