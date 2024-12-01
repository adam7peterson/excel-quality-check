import pandas as pd
import numpy as np
from pathlib import Path
from typing import Dict, List, Tuple

class ExcelQualityChecker:
    """A class to perform data quality checks on Excel files."""
    
    def __init__(self, file_path: str):
        """Initialize the checker with an Excel file path."""
        self.file_path = Path(file_path)
        if not self.file_path.exists():
            raise FileNotFoundError(f"Excel file not found: {file_path}")
        self.df = pd.read_excel(file_path)
        
    def check_null_values(self) -> Dict[str, float]:
        """Check for null values in each column."""
        null_percentages = (self.df.isnull().sum() / len(self.df) * 100).to_dict()
        return {col: round(pct, 2) for col, pct in null_percentages.items()}
    
    def check_duplicates(self) -> Dict[str, int]:
        """Check for duplicate rows and return statistics."""
        duplicates = self.df.duplicated()
        return {
            "total_rows": len(self.df),
            "duplicate_rows": duplicates.sum(),
            "duplicate_percentage": round(duplicates.sum() / len(self.df) * 100, 2)
        }
    
    def check_column_types(self) -> Dict[str, str]:
        """Get data types of each column."""
        return self.df.dtypes.astype(str).to_dict()
    
    def run_basic_checks(self) -> Dict[str, dict]:
        """Run all basic quality checks and return results."""
        return {
            "null_values": self.check_null_values(),
            "duplicates": self.check_duplicates(),
            "column_types": self.check_column_types()
        }

def main():
    """Example usage of the ExcelQualityChecker."""
    try:
        # Replace with your Excel file path
        checker = ExcelQualityChecker("your_file.xlsx")
        results = checker.run_basic_checks()
        
        # Print results in a formatted way
        print("\n=== Excel Data Quality Report ===\n")
        
        print("1. Null Values (% per column):")
        for col, pct in results["null_values"].items():
            print(f"   - {col}: {pct}%")
        
        print("\n2. Duplicate Rows:")
        for key, value in results["duplicates"].items():
            print(f"   - {key}: {value}")
        
        print("\n3. Column Data Types:")
        for col, dtype in results["column_types"].items():
            print(f"   - {col}: {dtype}")
            
    except Exception as e:
        print(f"Error: {str(e)}")

if __name__ == "__main__":
    main()