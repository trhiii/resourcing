import pandas as pd
import sqlite3
from datetime import datetime
import os

def process_resource_data():
    """
    Read the Excel file and create a SQLite database with worksheets starting with 'tbl'
    """
    # Excel file name
    excel_file = r"C:\Users\thockswender\OneDrive - InvestCloud\Tom\scratch\resources 1.0.xlsm"
    
    # Check if the Excel file exists
    if not os.path.exists(excel_file):
        print(f"Error: {excel_file} not found in the current directory")
        return
    
    # Generate datetime stamp for the database filename
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    db_filename = f"resources_{timestamp}.db"
    
    print(f"Processing Excel file: {excel_file}")
    print(f"Creating SQLite database: {db_filename}")
    
    try:
        # Read the Excel file to get worksheet names
        excel_data = pd.ExcelFile(excel_file)
        worksheet_names = excel_data.sheet_names
        
        print(f"Found {len(worksheet_names)} worksheets: {worksheet_names}")
        
        # Filter worksheets that start with 'tbl'
        tbl_worksheets = [name for name in worksheet_names if name.lower().startswith('tbl')]
        
        if not tbl_worksheets:
            print("No worksheets found starting with 'tbl'")
            print("Processing all worksheets instead...")
            tbl_worksheets = worksheet_names
        
        print(f"Processing {len(tbl_worksheets)} worksheets starting with 'tbl': {tbl_worksheets}")
        
        # Create SQLite database connection
        conn = sqlite3.connect(db_filename)
        
        # Process each worksheet starting with 'tbl'
        for worksheet_name in tbl_worksheets:
            try:
                print(f"Processing worksheet: {worksheet_name}")
                
                # Read the worksheet into a pandas DataFrame
                df = pd.read_excel(excel_file, sheet_name=worksheet_name)
                
                # Clean column names (remove spaces, special characters)
                df.columns = [str(col).replace(' ', '_').replace('-', '_').replace('(', '').replace(')', '') for col in df.columns]
                
                # Write the DataFrame to SQLite
                df.to_sql(worksheet_name, conn, if_exists='replace', index=False)
                
                print(f"  - Successfully added {len(df)} rows to table '{worksheet_name}'")
                
            except Exception as e:
                print(f"  - Error processing worksheet '{worksheet_name}': {str(e)}")
        
        # Close the database connection
        conn.close()
        
        print(f"\nDatabase creation completed successfully!")
        print(f"Database file: {db_filename}")
        
        # Show summary of created tables
        conn = sqlite3.connect(db_filename)
        cursor = conn.cursor()
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
        tables = cursor.fetchall()
        conn.close()
        
        print(f"Tables created: {[table[0] for table in tables]}")
        
        # Create joined output worksheet
        print("\nCreating joined output worksheet...")
        create_joined_output(excel_file, db_filename, timestamp)
        
    except Exception as e:
        print(f"Error processing Excel file: {str(e)}")

def create_joined_output(excel_file, db_filename, timestamp):
    """
    Create a joined output worksheet from the three main tables
    """
    try:
        # Connect to the database
        conn = sqlite3.connect(db_filename)
        
        # Create the joined query - left join from DW to UKG, PersonToTeam, Title Map, and TeamToBacklog
        join_query = """
        SELECT 
            dw.Employee_Number,
            ukg.*,
            pt.*,
            tm.*,
            tb.*
        FROM "tblDW" dw
        LEFT JOIN "tblUKG" ukg ON dw.Employee_Number = ukg.Employee_Number
        LEFT JOIN "tblPersonToTeam" pt ON dw.Employee_Number = pt.Employee_Number
        LEFT JOIN "tblTitleMap" tm ON ukg.Business_Title = tm.Business_Title
        LEFT JOIN "tblTeamToBacklog" tb ON pt.Team = tb.Team
        ORDER BY dw.Employee_Number
        """
        
        # First, get the count of rows in tblDW and all employee numbers to verify we don't lose any
        dw_count_query = 'SELECT COUNT(*) as count FROM "tblDW"'
        dw_count = pd.read_sql_query(dw_count_query, conn)
        dw_row_count = dw_count.iloc[0]['count']
        print(f"tblDW has {dw_row_count} rows")
        
        # Get all employee numbers from tblDW
        dw_employee_numbers = pd.read_sql_query('SELECT DISTINCT Employee_Number FROM "tblDW"', conn)
        
        # Execute the query and get the results
        joined_df = pd.read_sql_query(join_query, conn)
        conn.close()
        
        print(f"Joined data created with {len(joined_df)} rows")
        print(f"Joined DataFrame columns: {joined_df.columns.tolist()}")
        
        # Verify that all tblDW Employee_Number values are present in the joined output
        # Use the first Employee_Number column (from tblDW)
        joined_employee_numbers = joined_df.iloc[:, 0].dropna().unique()
        
        missing_employees = set(dw_employee_numbers['Employee_Number'].dropna().tolist()) - set(joined_employee_numbers)
        if missing_employees:
            print(f"ERROR: Missing {len(missing_employees)} employees from tblDW in joined output!")
            print(f"Missing Employee Numbers: {sorted(missing_employees)}")
        else:
            print("✓ All tblDW employees preserved in joined output")
        
        # Handle duplicate column names
        print("Checking for duplicate columns...")
        columns = joined_df.columns.tolist()
        seen_columns = {}
        new_columns = []
        
        for col in columns:
            if col in seen_columns:
                seen_columns[col] += 1
                new_col = f"{col}_{seen_columns[col]}"
                new_columns.append(new_col)
                print(f"  - Renamed duplicate column '{col}' to '{new_col}'")
            else:
                seen_columns[col] = 0
                new_columns.append(col)
        
        # Rename columns if there were duplicates
        if new_columns != columns:
            joined_df.columns = new_columns
            print(f"Renamed {len(new_columns) - len(columns)} duplicate columns")
        
        # Save the joined data to a table called "output" in the database
        conn = sqlite3.connect(db_filename)
        joined_df.to_sql("output", conn, if_exists='replace', index=False)
        print(f"Saved joined data to 'output' table in database")
        
        # Expand the DataFrame with missing calendar dates
        print("\nExpanding DataFrame with missing calendar dates...")
        expanded_df = expand_with_missing_dates(joined_df)
        
        # Add calculated columns to the expanded DataFrame
        print("\nAdding calculated columns to expanded data...")
        expanded_df = add_calculated_columns(expanded_df)
        
        # Save the expanded data to a table called "output_expanded" in the database
        expanded_df.to_sql("output_expanded", conn, if_exists='replace', index=False)
        conn.close()
        
        print(f"Saved expanded data to 'output_expanded' table in database")
        
        # Create a new Excel file with only the expanded data
        output_excel_file = f"resources_{timestamp}.xlsx"
        
        # Write only the expanded dataset to the Excel file
        try:
            # Use a more robust Excel writing approach
            expanded_df.to_excel(output_excel_file, sheet_name='output_expanded', index=False, engine='openpyxl')
            print(f"Successfully created Excel file: {output_excel_file}")
        except Exception as excel_error:
            print(f"Error writing Excel file: {str(excel_error)}")
            # Fallback: try writing as CSV
            csv_file = f"resources_{timestamp}.csv"
            expanded_df.to_csv(csv_file, index=False)
            print(f"Saved as CSV instead: {csv_file}")
        print(f"Expanded data has {len(expanded_df)} rows (added {len(expanded_df) - len(joined_df)} rows)")
        
    except Exception as e:
        print(f"Error creating joined output: {str(e)}")

def expand_with_missing_dates(df):
    """
    Expand the DataFrame using fill-down approach per employee.
    Each row is duplicated for all dates until the next row with a different date is encountered.
    Preserves the original sum of Percent values for each employee-date combination.
    """
    try:
        # Ensure the DataFrame is sorted by Employee_Number and AsOfDate
        df_sorted = df.sort_values(['Employee_Number', 'AsOfDate']).reset_index(drop=True)
        
        # Convert AsOfDate to datetime if it's not already
        df_sorted['AsOfDate'] = pd.to_datetime(df_sorted['AsOfDate'])
        
        # Get the date range for the entire dataset
        min_date = df_sorted['AsOfDate'].min()
        max_date = df_sorted['AsOfDate'].max()
        
        print(f"Date range: {min_date.date()} to {max_date.date()}")
        
        # Create a complete date range
        all_dates = pd.date_range(start=min_date, end=max_date, freq='D')
        
        expanded_rows = []
        
        # Get unique employee numbers
        employee_numbers = df_sorted['Employee_Number'].dropna().unique()
        total_employees = len(employee_numbers)
        
        for emp_idx, emp_num in enumerate(employee_numbers, 1):
            remaining_employees = total_employees - emp_idx
            print(f"Processing employee {emp_num}... ({remaining_employees} employees remaining)")
            
            # Get all records for this employee, sorted by date
            emp_records = df_sorted[df_sorted['Employee_Number'] == emp_num].copy().sort_values('AsOfDate')
            
            if len(emp_records) == 0:
                continue
            
            # For each date in the complete range
            for date in all_dates:
                # Find the most recent record before or on this date for this employee
                records_before_or_on = emp_records[emp_records['AsOfDate'] <= date]
                
                if not records_before_or_on.empty:
                    # Get the most recent date's records
                    most_recent_date = records_before_or_on['AsOfDate'].max()
                    most_recent_records = records_before_or_on[records_before_or_on['AsOfDate'] == most_recent_date]
                    
                    # Add all records for this date (preserve original percentages)
                    for _, record in most_recent_records.iterrows():
                        new_record = record.copy()
                        new_record['AsOfDate'] = date
                        expanded_rows.append(new_record.to_dict())
        
        # Create the expanded DataFrame
        expanded_df = pd.DataFrame(expanded_rows)
        
        # Sort by Employee_Number and AsOfDate
        expanded_df = expanded_df.sort_values(['Employee_Number', 'AsOfDate']).reset_index(drop=True)
        
        return expanded_df
        
    except Exception as e:
        print(f"Error expanding DataFrame with missing dates: {str(e)}")
        return df

def add_calculated_columns(df):
    """
    Add calculated columns based on AsOfDate and Percent values.
    """
    try:
        # Make a copy to avoid modifying the original DataFrame
        df_copy = df.copy()
        
        # Ensure AsOfDate is datetime
        df_copy['AsOfDate'] = pd.to_datetime(df_copy['AsOfDate'])
        
        # Add DayUnit column: 1 * percent
        print("  - Adding DayUnit column (1 * percent)...")
        df_copy['DayUnit'] = 1 * df_copy['Percent']
        
        # Add MonthlyUnit column: 1 / number of days in month * percent
        print("  - Adding MonthlyUnit column (1 / days_in_month * percent)...")
        # Get the number of days in each month
        days_in_month = df_copy['AsOfDate'].dt.days_in_month
        df_copy['MonthlyUnit'] = (1 / days_in_month) * df_copy['Percent']
        
        # Add Year column in YYYY format
        print("  - Adding Year column (YYYY format)...")
        df_copy['Year'] = df_copy['AsOfDate'].dt.strftime('%Y')
        
        # Add YearMonth column in YYYY-MM format
        print("  - Adding YearMonth column (YYYY-MM format)...")
        df_copy['YearMonth'] = df_copy['AsOfDate'].dt.strftime('%Y-%m')
        
        print(f"  - Successfully added 4 calculated columns to {len(df_copy)} rows")
        return df_copy
        
    except Exception as e:
        print(f"Error adding calculated columns: {str(e)}")
        return df

if __name__ == "__main__":
    process_resource_data()
