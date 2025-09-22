# Excel Advanced Analysis Module
# Contains advanced Excel file analysis features

from .formula_analysis import FormulaAnalyzer
from .formatting_comparison import FormattingComparator, compare_excel_formatting

__all__ = ['FormulaAnalyzer', 'FormattingComparator', 'compare_excel_formatting']