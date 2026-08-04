"""
ΣΑΠ-ΦΘ
Likert Engine
"""

import re
import unicodedata

import pandas as pd


LIKERT = {
    "παρα πολυ": 5,
    "πολυ": 4,
    "μετρια": 3,
    "λιγο": 2,
    "καθολου": 1,

    "συμφωνω απολυτα": 5,
    "συμφωνω": 4,
    "ουτε συμφωνω ουτε διαφωνω": 3,
    "διαφωνω": 2,
    "διαφωνω απολυτα": 1,

    "εξαιρετικα": 5,
    "πολυ καλα": 4,
    "καλα": 3,
    "μετρια": 3,
    "ανεπαρκως": 2,
    "πολυ ανεπαρκως": 1,
}


def normalize_text(value):
    text = str(value).strip().lower()

    text = unicodedata.normalize("NFD", text)
    text = "".join(
        char for char in text
        if unicodedata.category(char) != "Mn"
    )

    text = re.sub(r"\s+", " ", text)

    return text


def convert_value(value):

    if pd.isna(value):
        return pd.NA

    if isinstance(value, (int, float)):
        if 1 <= float(value) <= 5:
            return float(value)

    text = normalize_text(value)

    number_match = re.search(r"\b([1-5])\b", text)

    if number_match:
        return float(number_match.group(1))

    if text in LIKERT:
        return float(LIKERT[text])

    return pd.NA


def convert(series):

    return series.apply(convert_value).astype("Float64")