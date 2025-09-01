#!/usr/bin/env python3
"""
Validation script for production resource data.
This script runs validation tests on the actual production data to check for data quality issues.
"""

import pandas as pd
import sqlite3
import os
from datetime import datetime
import sys

# Add the current directory to the path so we can import the test module
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from test_process_resource_data import run_allocation_validation_tests


def validate_latest_production_data():
    """Validate the latest production data file."""
    
    # Find the most recent database file
    db_files = [f for f in os.listdir('.') if f.startswith('resources_') and f.endswith('.db')]
    
    if not db_files:
        print("No production database files found!")
        return
    
    # Get the most recent file
    latest_db = sorted(db_files)[-1]
    print(f"Validating production data from: {latest_db}")
    
    # Connect to the database
    conn = sqlite3.connect(latest_db)
    
    # Read the expanded data
    try:
        df = pd.read_sql_query("SELECT * FROM output_expanded", conn)
        print(f"Loaded {len(df)} records from output_expanded table")
    except Exception as e:
        print(f"Error reading output_expanded table: {e}")
        # Try reading the regular output table
        try:
            df = pd.read_sql_query("SELECT * FROM output", conn)
            print(f"Loaded {len(df)} records from output table")
        except Exception as e2:
            print(f"Error reading output table: {e2}")
            return
    
    conn.close()
    
    # Convert AsOfDate to datetime if it's not already
    if 'AsOfDate' in df.columns:
        df['AsOfDate'] = pd.to_datetime(df['AsOfDate'])
    
    # Run validation tests
    print("\nRunning validation tests...")
    errors = run_allocation_validation_tests(df)
    
    if errors:
        print(f"\n❌ Found {len(errors)} validation errors:")
        for error in errors:
            print(f"  - {error}")
    else:
        print("\n✅ No validation errors found!")
    
    # Additional analysis
    print("\n📊 Data Analysis:")
    
    # Employee count
    unique_employees = df['Employee_Number'].nunique()
    print(f"  - Total unique employees: {unique_employees}")
    
    # Date range
    if 'AsOfDate' in df.columns:
        min_date = df['AsOfDate'].min()
        max_date = df['AsOfDate'].max()
        print(f"  - Date range: {min_date.date()} to {max_date.date()}")
    
    # Team count
    if 'Team' in df.columns:
        unique_teams = df['Team'].nunique()
        print(f"  - Total unique teams: {unique_teams}")
    
    # Check for over-allocation by employee and month
    print("\n🔍 Checking for over-allocation patterns...")
    over_allocated_employees = set()
    
    for employee in df['Employee_Number'].unique():
        employee_data = df[df['Employee_Number'] == employee]
        
        for month in employee_data['AsOfDate'].dt.strftime('%Y-%m').unique():
            month_data = employee_data[employee_data['AsOfDate'].dt.strftime('%Y-%m') == month]
            total_allocation = month_data['Percent'].sum()
            
            if total_allocation > 100.01:  # Allow small rounding errors
                over_allocated_employees.add(employee)
                print(f"  - Employee {employee} over-allocated in {month}: {total_allocation:.2f}%")
    
    if over_allocated_employees:
        print(f"\n⚠️  Found {len(over_allocated_employees)} employees with over-allocation:")
        for emp in sorted(over_allocated_employees):
            print(f"    - {emp}")
    else:
        print("\n✅ No employees found with over-allocation!")
    
    # Check for employees with 0% allocation (exits)
    print("\n🔍 Checking for employee exits...")
    exit_records = df[df['Percent'] == 0]
    if len(exit_records) > 0:
        exit_employees = exit_records['Employee_Number'].unique()
        print(f"  - Found {len(exit_employees)} employees with exit records (0% allocation)")
        for emp in sorted(exit_employees):
            exit_dates = exit_records[exit_records['Employee_Number'] == emp]['AsOfDate'].dt.strftime('%Y-%m-%d').unique()
            print(f"    - {emp}: {', '.join(exit_dates)}")
    else:
        print("  - No exit records found")
    
    # Summary
    print(f"\n📋 Validation Summary:")
    print(f"  - Total records: {len(df)}")
    print(f"  - Validation errors: {len(errors)}")
    print(f"  - Over-allocated employees: {len(over_allocated_employees)}")
    
    if len(errors) == 0 and len(over_allocated_employees) == 0:
        print("\n🎉 All validation checks passed!")
    else:
        print("\n⚠️  Some validation issues found. Please review the data.")


if __name__ == '__main__':
    validate_latest_production_data()
