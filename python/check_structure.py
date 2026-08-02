"""
ΣΑΠ-ΦΘ
Έλεγχος δομής Google Forms
"""

from import_data import load_data


def main():

    df = load_data()

    print("\n--- Πεδία ερωτηματολογίου ---\n")

    for i, col in enumerate(df.columns, start=1):

        print(f"{i:02d}. {col}")


if __name__ == "__main__":

    main()
