"""Input validation and conversion."""

import re
from typing import List, Union
from ..models import ValidatedParams
from ..exceptions import ValidationError
from ..utils.year_converter import convert_to_roc_year
from ..config import MIN_YEAR, MAX_YEAR
from datetime import datetime

class InputValidator:
    """Validates and converts user input parameters."""
    
    @staticmethod
    def validate_company_id(company_id: str) -> str:
        """Validate Taiwan stock company ID format."""
        if not company_id:
            raise ValidationError("Company ID cannot be empty")
        
        # Convert to string and strip whitespace
        company_id = str(company_id).strip()
        
        # Taiwan stock codes are typically 4 digits
        if not re.match(r'^\d{4}$', company_id):
            raise ValidationError(f"Company ID must be 4 digits, got: {company_id}")
        
        return company_id
    
    @staticmethod
    def validate_year(year: int) -> tuple[int, int]:
        """Validate year and return both Western and ROC years."""
        if not isinstance(year, int):
            try:
                year = int(year)
            except (ValueError, TypeError):
                raise ValidationError(f"Year must be an integer, got: {year}")
        
        current_year = datetime.now().year
        if year < MIN_YEAR or year > MAX_YEAR:
            raise ValidationError(f"Year must be between {MIN_YEAR} and {MAX_YEAR}, got: {year}")
        
        if year > current_year:
            raise ValidationError(f"Year cannot be in the future, got: {year}")
        
        roc_year = convert_to_roc_year(year)
        return year, roc_year
    
    @staticmethod
    def validate_quarter(quarter: Union[str, int, List[int]]) -> List[int]:
        """Validate quarter parameter and return list of quarters."""
        if quarter == "all" or quarter is None:
            return [1, 2, 3, 4]
        
        if isinstance(quarter, (int, str)):
            try:
                q = int(quarter)
                if q not in [1, 2, 3, 4]:
                    raise ValidationError(f"Quarter must be 1, 2, 3, 4, or 'all', got: {q}")
                return [q]
            except (ValueError, TypeError):
                raise ValidationError(f"Invalid quarter format: {quarter}")
        
        if isinstance(quarter, list):
            quarters = []
            for q in quarter:
                try:
                    q_int = int(q)
                    if q_int not in [1, 2, 3, 4]:
                        raise ValidationError(f"Quarter must be 1, 2, 3, 4, got: {q_int}")
                    quarters.append(q_int)
                except (ValueError, TypeError):
                    raise ValidationError(f"Invalid quarter in list: {q}")
            return sorted(list(set(quarters)))  # Remove duplicates and sort
        
        raise ValidationError(f"Invalid quarter type: {type(quarter)}")
    
    def validate_and_convert(self, company_id: str, year: int, quarter: Union[str, int] = "all") -> ValidatedParams:
        """Validate all inputs and return validated parameters."""
        validated_company_id = self.validate_company_id(company_id)
        western_year, roc_year = self.validate_year(year)
        quarters = self.validate_quarter(quarter)
        
        return ValidatedParams(
            company_id=validated_company_id,
            western_year=western_year,
            roc_year=roc_year,
            quarters=quarters
        )
