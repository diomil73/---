"""
ΣΑΠ-ΦΘ
Likert Engine
"""

import re
import unicodedata

import pandas as pd


LIKERT = {
    # Κλίμακα βαθμού / ποσότητας
    "παρα πολυ": 5,
    "πολυ": 4,
    "μετρια": 3,
    "λιγο": 2,
    "καθολου": 1,

    # Κλίμακα έντασης / συμμετοχής
    "πολυ υψηλη": 5,
    "υψηλη": 4,
    "χαμηλη": 2,
    "πολυ χαμηλη": 1,

    # Κλίμακα συμφωνίας
    "συμφωνω απολυτα": 5,
    "συμφωνω": 4,
    "ουτε συμφωνω ουτε διαφωνω": 3,
    "διαφωνω": 2,
    "διαφωνω απολυτα": 1,

    # Κλίμακα αξιολόγησης
    "εξαιρετικα": 5,
    "πολυ καλα": 4,
    "καλα": 3,
    "ανεπαρκως": 2,
    "πολυ ανεπαρκως": 1,
}


NON_APPLICABLE = {
    "δεν συμμετειχα",
    "δεν εφαρμοζεται",
    "δεν γνωριζω",
    "δεν ειχα εμπειρια",
}


def normalize_text(value):
    text = str(value).strip().lower()

    text = unicodedata.normalize("NFD", text)
    text = "".join(
        char
        for char in text
        if unicodedata.category(char) != "Mn"
    )

    text = re.sub(r"\s+", " ", text)

    return text


def is_non_applicable(value):
    if pd.isna(value):
        return True

    return normalize_text(value) in NON_APPLICABLE


def convert_value(value):
    if pd.isna(value):
        return pd.NA

    if isinstance(value, (int, float)):
        numeric_value = float(value)

        if 1 <= numeric_value <= 5:
            return numeric_value

    text = normalize_text(value)

    if text in NON_APPLICABLE:
        return pd.NA

    number_match = re.search(r"\b([1-5])\b", text)

    if number_match:
        return float(number_match.group(1))

    if text in LIKERT:
        return float(LIKERT[text])

    return pd.NA


def convert(series):
    return series.apply(convert_value).astype("Float64")