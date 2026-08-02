"""
ΣΑΠ-ΦΘ
IQI
"""

from section_scores import section_means


def calculate_iqi():

    sections = section_means()

    return round(sections["Mean"].mean(), 2)


if __name__ == "__main__":

    print("IQI =", calculate_iqi())
