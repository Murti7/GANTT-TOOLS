"""Modelo de producto por capabilities."""

from enum import StrEnum


class Capability(StrEnum):
    """Capacidades funcionales principales del producto."""

    BUDGETING = "budgeting"
    PLANNING = "planning"
    BILLING = "billing"


class OutputFamily(StrEnum):
    """Familias de salida del filesystem."""

    DOCUMENTS = "documents"
    ANALYSIS = "analysis"
    CHARTS = "charts"
    DRAFTS = "drafts"
    ISSUED = "issued"
