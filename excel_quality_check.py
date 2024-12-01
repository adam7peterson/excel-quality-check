import pandas as pd
import numpy as np
from pathlib import Path
from typing import Dict, List, Tuple
import matplotlib.pyplot as plt
import seaborn as sns
from scipy import stats

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

    def detect_outliers(self, method='zscore', threshold=3) -> Dict[str, List[float]]:
        """Detect outliers in numerical columns using z-score or IQR method."""
        outliers = {}
        numeric_cols = self.df.select_dtypes(include=[np.number]).columns
        
        for col in numeric_cols:
            if method == 'zscore':
                z_scores = np.abs(stats.zscore(self.df[col].dropna()))
                outliers[col] = self.df[col][z_scores > threshold].tolist()
            elif method == 'iqr':
                Q1 = self.df[col].quantile(0.25)
                Q3 = self.df[col].quantile(0.75)
                IQR = Q3 - Q1
                outliers[col] = self.df[col][
                    (self.df[col] < (Q1 - 1.5 * IQR)) | 
                    (self.df[col] > (Q3 + 1.5 * IQR))].tolist()
        return outliers

    def check_format_consistency(self) -> Dict[str, Dict[str, int]]:
        """Check formatting consistency in string columns."""
        format_issues = {}
        string_cols = self.df.select_dtypes(include=['object']).columns
        
        for col in string_cols:
            patterns = {
                'mixed_case': sum(s != s.lower() and s != s.upper() 
                                 for s in self.df[col].dropna()),
                'leading_trailing_spaces': sum(str(s).strip() != str(s) 
                                              for s in self.df[col].dropna()),
                'special_characters': sum(not str(s).isalnum() 
                                         for s in self.df[col].dropna())
            }
            format_issues[col] = patterns
        return format_issues

    def validate_formulas(self) -> Dict[str, List[str]]:
        """Check for common Excel formula errors."""
        # Note: This is a basic implementation that looks for common error strings
        error_patterns = ['#DIV/0!', '#N/A', '#NAME?', '#NULL!', '#NUM!', '#REF!', '#VALUE!']
        formula_errors = {}
        
        for col in self.df.columns:
            errors = []
            for pattern in error_patterns:
                if self.df[col].astype(str).str.contains(pattern).any():
                    errors.append(pattern)
            if errors:
                formula_errors[col] = errors
        return formula_errors

    def visualize_data_quality(self, save_path: str = None):
        """Create visualizations for data quality metrics."""
        # Set up the plotting style
        plt.style.use('seaborn')
        
        # Create a figure with multiple subplots
        fig = plt.figure(figsize=(15, 10))
        
        # 1. Null Values Heatmap
        plt.subplot(2, 2, 1)
        null_matrix = self.df.isnull()
        sns.heatmap(null_matrix, cbar=True)
        plt.title('Missing Values Heatmap')
        
        # 2. Data Types Distribution
        plt.subplot(2, 2, 2)
        dtype_counts = self.df.dtypes.value_counts()
        dtype_counts.plot(kind='bar')
        plt.title('Data Types Distribution')
        plt.xticks(rotation=45)
        
        # 3. Numeric Columns Distribution
        numeric_cols = self.df.select_dtypes(include=[np.number]).columns
        if len(numeric_cols) > 0:
            plt.subplot(2, 2, 3)
            self.df[numeric_cols].boxplot()
            plt.title('Numeric Columns Distribution')
            plt.xticks(rotation=45)
        
        plt.tight_layout()
        
        if save_path:
            plt.savefig(save_path)
            plt.close()
        else:
            plt.show()

    def generate_cleaning_recommendations(self) -> List[str]:
        """Generate recommendations for data cleaning based on the analysis."""
        recommendations = []
        
        # Check null values
        null_vals = self.check_null_values()
        for col, pct in null_vals.items():
            if pct > 0:
                if pct > 50:
                    recommendations.append(f"Consider removing column '{col}' due to high missing values ({pct}%)")
                else:
                    recommendations.append(f"Handle missing values in column '{col}' ({pct}%)")
        
        # Check duplicates
        dups = self.check_duplicates()
        if dups['duplicate_rows'] > 0:
            recommendations.append(f"Remove {dups['duplicate_rows']} duplicate rows")
        
        # Check formatting
        format_issues = self.check_format_consistency()
        for col, issues in format_issues.items():
            if issues['mixed_case'] > 0:
                recommendations.append(f"Standardize case formatting in column '{col}'")
            if issues['leading_trailing_spaces'] > 0:
                recommendations.append(f"Remove leading/trailing spaces in column '{col}'")
        
        return recommendations

    def run_advanced_checks(self) -> Dict[str, dict]:
        """Run all advanced quality checks and return results."""
        return {
            "basic_checks": self.run_basic_checks(),
            "outliers": self.detect_outliers(),
            "format_consistency": self.check_format_consistency(),
            "formula_errors": self.validate_formulas(),
            "cleaning_recommendations": self.generate_cleaning_recommendations()
        }

def main():
    """Example usage of the ExcelQualityChecker."""
    try:
        # Replace with your Excel file path
        checker = ExcelQualityChecker("your_file.xlsx")
        results = checker.run_advanced_checks()
        
        # Print results in a formatted way
        print("\n=== Excel Data Quality Report ===\n")
        
        # Basic checks
        print("1. Basic Checks:")
        print("\n   Null Values (% per column):")
        for col, pct in results["basic_checks"]["null_values"].items():
            print(f"   - {col}: {pct}%")
        
        print("\n   Duplicate Rows:")
        for key, value in results["basic_checks"]["duplicates"].items():
            print(f"   - {key}: {value}")
        
        # Advanced checks
        print("\n2. Outliers Detected:")
        for col, outliers in results["outliers"].items():
            if outliers:
                print(f"   - {col}: {len(outliers)} outliers found")
        
        print("\n3. Format Consistency Issues:")
        for col, issues in results["format_consistency"].items():
            print(f"   - {col}:")
            for issue, count in issues.items():
                if count > 0:
                    print(f"     * {issue}: {count} instances")
        
        print("\n4. Formula Errors:")
        if results["formula_errors"]:
            for col, errors in results["formula_errors"].items():
                print(f"   - {col}: {', '.join(errors)}")
        else:
            print("   No formula errors found")
        
        print("\n5. Cleaning Recommendations:")
        for i, rec in enumerate(results["cleaning_recommendations"], 1):
            print(f"   {i}. {rec}")
        
        # Generate visualizations
        checker.visualize_data_quality("data_quality_report.png")
            
    except Exception as e:
        print(f"Error: {str(e)}")

if __name__ == "__main__":
    main()