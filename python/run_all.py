"""
ΣΑΠ-ΦΘ
Κεντρική εκτέλεση όλων των αναφορών.
"""

import sys
import time
import traceback
from pathlib import Path

from final_report import create_final_report
from master_excel import create_master
from question_analytics import create_question_analytics
from text_analytics import create_text_analytics


PROJECT_ROOT = Path(__file__).resolve().parent.parent
EXCEL_FOLDER = PROJECT_ROOT / "excel"


def print_separator():
    """
    Εκτυπώνει διαχωριστική γραμμή στο terminal.
    """

    print("=" * 70)


def run_step(step_number, title, function):
    """
    Εκτελεί ένα βήμα και εμφανίζει την πρόοδό του.
    """

    print_separator()
    print(f"ΒΗΜΑ {step_number}: {title}")
    print_separator()

    start_time = time.time()

    try:
        function()

        elapsed_time = time.time() - start_time

        print(
            f"\nΟλοκληρώθηκε επιτυχώς σε "
            f"{elapsed_time:.2f} δευτερόλεπτα."
        )

        return True

    except PermissionError:
        print("\nΣΦΑΛΜΑ ΠΡΟΣΒΑΣΗΣ")

        print(
            "Δεν ήταν δυνατή η αποθήκευση του αρχείου Excel."
        )

        print(
            "Κλείσε όλα τα αρχεία Excel που δημιουργεί "
            "το πρόγραμμα και εκτέλεσε ξανά την εντολή."
        )

        return False

    except FileNotFoundError as error:
        print("\nΔΕΝ ΒΡΕΘΗΚΕ ΑΡΧΕΙΟ")

        print(error)

        print(
            "\nΈλεγξε ότι υπάρχει το αρχείο:"
        )

        print(
            PROJECT_ROOT
            / "data"
            / "google_forms_export.xlsx"
        )

        return False

    except Exception as error:
        print("\nΠΑΡΟΥΣΙΑΣΤΗΚΕ ΣΦΑΛΜΑ")

        print(
            f"{type(error).__name__}: {error}"
        )

        print(
            "\nΑναλυτικές πληροφορίες:"
        )

        traceback.print_exc()

        return False


def list_created_files():
    """
    Εμφανίζει τα Excel αρχεία που υπάρχουν
    στον φάκελο εξόδου.
    """

    print_separator()
    print("ΑΡΧΕΙΑ ΣΤΟΝ ΦΑΚΕΛΟ EXCEL")
    print_separator()

    if not EXCEL_FOLDER.exists():
        print(
            "Ο φάκελος excel δεν υπάρχει."
        )

        return

    excel_files = sorted(
        EXCEL_FOLDER.glob("*.xlsx")
    )

    if not excel_files:
        print(
            "Δεν βρέθηκαν αρχεία Excel."
        )

        return

    for file_path in excel_files:
        file_size_kb = (
            file_path.stat().st_size / 1024
        )

        print(
            f"- {file_path.name} "
            f"({file_size_kb:.1f} KB)"
        )


def main():
    """
    Εκτελεί όλες τις αναφορές του ΣΑΠ-ΦΘ.
    """

    total_start_time = time.time()

    print()
    print_separator()
    print("ΣΑΠ-ΦΘ")
    print("ΔΗΜΙΟΥΡΓΙΑ ΟΛΩΝ ΤΩΝ ΑΝΑΦΟΡΩΝ")
    print_separator()

    EXCEL_FOLDER.mkdir(
        parents=True,
        exist_ok=True,
    )

    steps = [
        (
            1,
            "Δημιουργία MASTER Excel",
            create_master,
        ),
        (
            2,
            "Ανάλυση ερωτήσεων Likert",
            create_question_analytics,
        ),
        (
            3,
            "Ανάλυση ανοιχτών σχολίων",
            create_text_analytics,
        ),
        (
            4,
            "Δημιουργία τελικής αναφοράς",
            create_final_report,
        ),
    ]

    completed_steps = 0

    for step_number, title, function in steps:
        success = run_step(
            step_number=step_number,
            title=title,
            function=function,
        )

        if not success:
            print_separator()

            print(
                "Η διαδικασία σταμάτησε λόγω σφάλματος."
            )

            print(
                f"Ολοκληρώθηκαν {completed_steps} "
                f"από τα {len(steps)} βήματα."
            )

            sys.exit(1)

        completed_steps += 1

    total_elapsed_time = (
        time.time() - total_start_time
    )

    print()
    list_created_files()

    print()
    print_separator()
    print("Η ΔΙΑΔΙΚΑΣΙΑ ΟΛΟΚΛΗΡΩΘΗΚΕ ΕΠΙΤΥΧΩΣ")
    print_separator()

    print(
        f"Ολοκληρώθηκαν {completed_steps} "
        f"από τα {len(steps)} βήματα."
    )

    print(
        f"Συνολικός χρόνος: "
        f"{total_elapsed_time:.2f} δευτερόλεπτα."
    )

    print(
        f"Φάκελος αποτελεσμάτων: "
        f"{EXCEL_FOLDER}"
    )


if __name__ == "__main__":
    main()