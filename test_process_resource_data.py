import unittest
import pandas as pd
import numpy as np
from datetime import datetime, date
import tempfile
import os
import sys

# Add the current directory to the path so we can import the main module
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from process_resource_data import expand_with_missing_dates, add_calculated_columns


class TestResourceDataProcessing(unittest.TestCase):
    """Test suite for resource data processing functionality."""
    
    def setUp(self):
        """Set up test data before each test."""
        # Create a minimal test dataset that matches the actual structure
        self.test_data = pd.DataFrame({
            'Employee_Number': ['TEST001', 'TEST001', 'TEST002'],
            'Team': ['Team A', 'Team B', 'Team A'],
            'Percent': [100.0, 0.0, 100.0],  # TEST001 exits, TEST002 stays
            'AsOfDate': ['2025-01-01', '2025-06-01', '2025-01-01'],
            'Employee_Name': ['John Doe', 'John Doe', 'Jane Smith'],
            'Squad_Org': ['Engineering', 'Engineering', 'Engineering'],
            'Squad': ['Platform', 'Platform', 'Platform'],
            'Group': ['Backend', 'Backend', 'Backend'],
            'Subgroup': ['API', 'API', 'API'],
            'Business_Title': ['Engineer', 'Engineer', 'Engineer'],
            'Role': ['Developer', 'Developer', 'Developer'],
            'Product_Area': ['Platform', 'Platform', 'Platform'],
            'Projects': ['API Gateway', 'API Gateway', 'API Gateway']
        })
    
    def test_basic_functionality(self):
        """Test basic functionality with simple test data."""
        df = self.test_data.copy()
        expanded_df = expand_with_missing_dates(df)
        
        # Basic checks
        self.assertIsInstance(expanded_df, pd.DataFrame)
        self.assertGreater(len(expanded_df), len(df))
        
        # Check that TEST001 (who exits) doesn't have records after June
        test001_records = expanded_df[expanded_df['Employee_Number'] == 'TEST001']
        if len(test001_records) > 0:
            june_records = test001_records[
                test001_records['AsOfDate'].dt.strftime('%Y-%m') == '2025-06'
            ]
            self.assertEqual(len(june_records), 0, "TEST001 should not have records in June (exit month)")
    
    def test_over_allocation_detection(self):
        """Test that over-allocation (>100%) is detected."""
        
        # Create test data with over-allocation
        over_allocated_data = pd.DataFrame({
            'Employee_Number': ['TEST001', 'TEST001'],
            'Team': ['Team A', 'Team B'],
            'Percent': [60.0, 50.0],  # 110% total
            'AsOfDate': ['2025-01-01', '2025-01-01'],
            'Employee_Name': ['John Doe', 'John Doe'],
            'Squad_Org': ['Engineering', 'Engineering'],
            'Squad': ['Platform', 'Platform'],
            'Group': ['Backend', 'Backend'],
            'Subgroup': ['API', 'API'],
            'Business_Title': ['Engineer', 'Engineer'],
            'Role': ['Developer', 'Developer'],
            'Product_Area': ['Platform', 'Platform'],
            'Projects': ['API Gateway', 'API Gateway']
        })
        
        expanded_df = expand_with_missing_dates(over_allocated_data)
        
        # Check for over-allocation
        test001_records = expanded_df[expanded_df['Employee_Number'] == 'TEST001']
        if len(test001_records) > 0:
            jan_records = test001_records[
                test001_records['AsOfDate'].dt.strftime('%Y-%m') == '2025-01'
            ]
            if len(jan_records) > 0:
                total_allocation = jan_records['Percent'].sum()
                self.assertGreater(total_allocation, 100.0, 
                                 f"Expected over-allocation, got {total_allocation}%")
    
    def test_fractional_allocation_validation(self):
        """Test that fractional allocations sum correctly."""
        
        # Create test data with proper fractional allocation
        fractional_data = pd.DataFrame({
            'Employee_Number': ['TEST001', 'TEST001'],
            'Team': ['Team A', 'Team B'],
            'Percent': [60.0, 40.0],  # 100% total
            'AsOfDate': ['2025-01-01', '2025-01-01'],
            'Employee_Name': ['John Doe', 'John Doe'],
            'Squad_Org': ['Engineering', 'Engineering'],
            'Squad': ['Platform', 'Platform'],
            'Group': ['Backend', 'Backend'],
            'Subgroup': ['API', 'API'],
            'Business_Title': ['Engineer', 'Engineer'],
            'Role': ['Developer', 'Developer'],
            'Product_Area': ['Platform', 'Platform'],
            'Projects': ['API Gateway', 'API Gateway']
        })
        
        expanded_df = expand_with_missing_dates(fractional_data)
        
        # Check that allocations sum to 100%
        test001_records = expanded_df[expanded_df['Employee_Number'] == 'TEST001']
        if len(test001_records) > 0:
            jan_records = test001_records[
                test001_records['AsOfDate'].dt.strftime('%Y-%m') == '2025-01'
            ]
            if len(jan_records) > 0:
                total_allocation = jan_records['Percent'].sum()
                self.assertAlmostEqual(total_allocation, 100.0, places=2,
                                     msg=f"Expected 100% allocation, got {total_allocation}%")
    
    def test_calculated_columns(self):
        """Test that calculated columns are added correctly."""
        df = self.test_data.copy()
        expanded_df = expand_with_missing_dates(df)
        result_df = add_calculated_columns(expanded_df)
        
        # Check that required columns exist
        required_columns = ['Year', 'YearMonth', 'MonthlyUnit', 'DayUnit']
        for col in required_columns:
            self.assertIn(col, result_df.columns, f"Missing required column: {col}")
        
        # Check that Year and YearMonth are correct for a sample record
        if len(result_df) > 0:
            sample_record = result_df.iloc[0]
            self.assertEqual(sample_record['Year'], '2025')
            self.assertEqual(sample_record['YearMonth'], '2025-01')
            
            # Check that MonthlyUnit equals Percent
            self.assertEqual(sample_record['MonthlyUnit'], sample_record['Percent'])
    
    def test_employee_exit_handling(self):
        """Test that employee exits are handled correctly."""
        df = self.test_data.copy()
        expanded_df = expand_with_missing_dates(df)
        
        # TEST001 should not have records after their exit date (June 2025)
        test001_records = expanded_df[expanded_df['Employee_Number'] == 'TEST001']
        
        if len(test001_records) > 0:
            # Check that no records exist after June
            later_months = ['2025-07', '2025-08', '2025-09', '2025-10', '2025-11', '2025-12']
            for month in later_months:
                month_records = test001_records[
                    test001_records['AsOfDate'].dt.strftime('%Y-%m') == month
                ]
                self.assertEqual(len(month_records), 0, 
                               f"TEST001 should not have records in {month} after exit")
    
    def test_data_validation_function(self):
        """Test the data validation function."""
        df = self.test_data.copy()
        expanded_df = expand_with_missing_dates(df)
        
        # Run validation
        errors = run_allocation_validation_tests(expanded_df)
        
        # Should not have any validation errors for this clean test data
        self.assertEqual(len(errors), 0, f"Found validation errors: {errors}")


def run_allocation_validation_tests(df):
    """
    Run validation tests on the expanded DataFrame to check for data quality issues.
    Returns a list of validation errors found.
    """
    errors = []
    
    # Check for over-allocation (>100%) per employee per month
    for employee in df['Employee_Number'].unique():
        employee_data = df[df['Employee_Number'] == employee]
        
        for month in employee_data['AsOfDate'].dt.strftime('%Y-%m').unique():
            month_data = employee_data[employee_data['AsOfDate'].dt.strftime('%Y-%m') == month]
            total_allocation = month_data['Percent'].sum()
            
            if total_allocation > 100.01:  # Allow small rounding errors
                errors.append(f"Employee {employee} over-allocated in {month}: {total_allocation:.2f}%")
    
    # Check for negative percentages
    negative_percentages = df[df['Percent'] < 0]
    if len(negative_percentages) > 0:
        errors.append(f"Found {len(negative_percentages)} records with negative percentages")
    
    # Check for missing required columns
    required_columns = ['Employee_Number', 'Team', 'Percent', 'AsOfDate']
    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        errors.append(f"Missing required columns: {missing_columns}")
    
    return errors


if __name__ == '__main__':
    # Run the tests
    unittest.main(verbosity=2)
