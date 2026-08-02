"""
Dashboard Tables
"""

from section_scores import section_means
from overall_iqi import calculate_iqi


def dashboard_data():

    sections = section_means()

    return {

        "IQI": calculate_iqi(),

        "Sections": sections

    }
