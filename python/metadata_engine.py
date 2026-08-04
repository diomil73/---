"""
ΣΑΠ-ΦΘ
Metadata Engine
"""

import pandas as pd

from import_data import load_data


SECTION_QUESTIONS = {
    "Α": [
        "Υποδοχή",
        "Ενημέρωση",
    ],

    "Β": [
        "Οργάνωση προγράμματος",
        "Κατανομή περιστατικών",
        "Συνέπεια προγράμματος",
        "Διαθεσιμότητα υπευθύνων",
        "Συνολική οργάνωση",
    ],

    "Γ": [
        "Νευρολογική αποκατάσταση",
        "Ορθοπαιδική αποκατάσταση",
        "Γηριατρική αποκατάσταση",
        "Ρομποτική αποκατάσταση",
        "Θεραπευτική άσκηση",
        "Υδροθεραπεία εντός πισίνας",
        "Κρουστικά κύματα",
        "Εκπαίδευση στη βάδιση",
        "Αξιολόγηση ασθενών",
        "Καταγραφή στον φάκελο ασθενούς",
        "Παρουσίαση περιστατικών",
        "Συμμετοχή στη διεπιστημονική ομάδα",
        "Ανάπτυξη κλινικής σκέψης",
        "Εφαρμογή επιστημονικής γνώσης",
        "Ευκαιρίες πρακτικής εξάσκησης",
        "Εποικοδομητική ανατροφοδότηση",
        "Αυτονομία υπό επίβλεψη",
        "Απόκτηση νέων δεξιοτήτων",
    ],

    "Δ": [
        "Σύγχρονος εξοπλισμός",
        "Επάρκεια υλικών",
        "Καθαριότητα",
        "Ασφάλεια",
        "Οργάνωση χώρων",
        "Διαθεσιμότητα μηχανημάτων",
    ],

    "Ε": [
        "Φυσικοθεραπευτές",
        "Ιατροί",
        "Εργοθεραπευτές",
        "Λογοθεραπευτές",
        "Ψυχολόγοι",
        "Κοινωνικοί λειτουργοί",
        "Διοικητικό προσωπικό",
    ],

    "ΣΤ": [
        "Διαθεσιμότητα",
        "Καθοδήγηση",
        "Μεταδοτικότητα γνώσεων",
        "Ανατροφοδότηση",
        "Δικαιοσύνη",
        "Υποστήριξη",
        "Προσιτότητα",
        "Ενδιαφέρον για τους φοιτητές",
    ],

    "Ζ": [
        "Ηγεσία",
        "Οργάνωση",
        "Επαγγελματισμός",
        "Δημιουργία θετικού κλίματος",
    ],

    "Η": [
        "Ένιωσα μέλος της ομάδας.",
        "Με αντιμετώπισαν με σεβασμό.",
        "Μπορούσα να εκφράζω απορίες.",
        "Ένιωθα ασφαλής να κάνω λάθος και να μάθω.",
        "Η συνεργασία ήταν αποτελεσματική.",
    ],

    "Θ": [
        "Χώρος εργασίας",
        "Αποδυτήρια",
        "Χώρος διαλείμματος",
        "Υγιεινή",
        "Ωράριο",
        "Πρόσβαση σε εκπαιδευτικό υλικό",
    ],
}


DEMOGRAPHIC_QUESTIONS = {
    "Timestamp",
    "Πανεπιστήμιο",
    "Έτος σπουδών",
    "Διάρκεια πρακτικής",
    "Περίοδος πρακτικής",
}


TEXT_QUESTIONS = {
    "Ποια θεωρείτε ότι ήταν τα σημαντικότερα θετικά στοιχεία της πρακτικής σας άσκησης;",
    "Ποιοι τομείς χρειάζονται βελτίωση;",
    "Ποιες εκπαιδευτικές δραστηριότητες θα θέλατε να ενισχυθούν;",
    "Υπάρχουν δραστηριότητες που θα θέλατε να προστεθούν;",
    "Περιγράψτε μια εμπειρία που θεωρείτε ιδιαίτερα σημαντική κατά τη διάρκεια της πρακτικής σας.",
    "Επιπλέον σχόλια ή προτάσεις.",
}


def clean_question(question):
    """
    Αφαιρεί κενά και αγκύλες από τα ονόματα των ερωτήσεων.
    """

    return str(question).strip().strip("[]").strip()


def detect_section(question):
    """
    Επιστρέφει την ενότητα στην οποία ανήκει η ερώτηση.
    """

    cleaned = clean_question(question)

    if cleaned in DEMOGRAPHIC_QUESTIONS:
        return "Δημογραφικά"

    for section, questions in SECTION_QUESTIONS.items():
        if cleaned in questions:
            return section

    if cleaned in TEXT_QUESTIONS:
        return "Ι"

    if "Πόσο πιθανό είναι να προτείνατε" in cleaned:
        return "Γενική αξιολόγηση"

    return "Λοιπά"


def detect_type(question):
    """
    Καθορίζει τον τύπο της ερώτησης.
    """

    cleaned = clean_question(question)

    if cleaned in DEMOGRAPHIC_QUESTIONS:
        return "OTHER"

    if cleaned in TEXT_QUESTIONS:
        return "TEXT"

    if "Πόσο πιθανό είναι να προτείνατε" in cleaned:
        return "NPS"

    for questions in SECTION_QUESTIONS.values():
        if cleaned in questions:
            return "LIKERT"

    return "OTHER"


def build_metadata():
    """
    Δημιουργεί το metadata table για όλες τις στήλες.
    """

    df = load_data()

    records = []

    for question in df.columns:
        question_type = detect_type(question)
        section = detect_section(question)

        records.append({
            "Question": question,
            "CleanQuestion": clean_question(question),
            "Type": question_type,
            "Section": section,
            "Scored": question_type in {"LIKERT", "NPS"},
        })

    return pd.DataFrame(records)


if __name__ == "__main__":
    meta = build_metadata()

    print(
        meta.groupby(
            ["Section", "Type"]
        ).size().to_string()
    )

    print()

    print(
        meta[
            [
                "Section",
                "Type",
                "Question",
            ]
        ].to_string(index=False)
    )