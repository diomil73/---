"""
ΣΑΠ-ΦΘ
Main Entry Point
"""

from generator import create_master


def main():

    print("=" * 60)
    print("ΣΑΠ-ΦΘ")
    print("Σύστημα Αξιολόγησης και Ποιότητας Τμήματος Φυσικοθεραπείας")
    print("=" * 60)

    create_master()

    print("\nΟλοκληρώθηκε επιτυχώς.")


if __name__ == "__main__":
    main()
