"""Year conversion utilities."""

from ..exceptions import ValidationError

def convert_to_roc_year(western_year: int) -> int:
    """Convert Western year to ROC year format for MOPS API."""
    if western_year < 1912:
        raise ValidationError("Year must be 1912 or later (ROC establishment)")
    return western_year - 1911

def convert_to_western_year(roc_year: int) -> int:
    """Convert ROC year to Western year format."""
    if roc_year < 1:
        raise ValidationError("ROC year must be positive")
    return roc_year + 1911
