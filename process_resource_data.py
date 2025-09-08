import pandas as pd 
import sqlite3
from datetime import datetime
import os
import shutil
import xlwings as xw

def copy_data_worksheet_to_source(clean_output_file, source_file_path):
    """
    Copy the DATA worksheet from the clean output file to the source file using xlwings.
    This preserves all macros and file integrity.
    """
    try:
        print(f"Copying DATA worksheet from clean output to source file...")
        print(f"  - Source: {source_file_path}")
        print(f"  - Clean output: {clean_output_file}")
        
        # Open both files with xlwings (Excel doesn't need to be visible)
        app = xw.App(visible=False, add_book=False)
        
        try:
            # Open the clean output file (read-only)
            output_wb = app.books.open(clean_output_file, read_only=True)
            output_ws = output_wb.sheets['DATA']
            
            # Open the source file
            source_wb = app.books.open(source_file_path)
            
            # Remove existing DATA worksheet if it exists
            if 'DATA' in [sheet.name for sheet in source_wb.sheets]:
                source_wb.sheets['DATA'].delete()
                print("  - Removed existing DATA worksheet from source")
            
            # Copy the DATA worksheet from output to source
            output_ws.api.Copy(Before=source_wb.sheets[0].api)
            print("  - Copied DATA worksheet to source file")
            
            # Save the source file
            source_wb.save()
            print("  - Saved source file with DATA worksheet")
            
            # Close workbooks
            output_wb.close()
            source_wb.close()
            
            print("  - Successfully copied DATA worksheet while preserving all macros and worksheets")
            return True
            
        finally:
            # Always quit the Excel application
            app.quit()
            
    except Exception as e:
        print(f"  - Error copying DATA worksheet: {str(e)}")
        try:
            app.quit()
        except:
            pass
        return False

def process_resource_data():
    """
    Read the Excel file and create a SQLite database with worksheets starting with 'tbl'
    """
    # Excel file name
    excel_file = r"C:\Users\thockswender\OneDrive - InvestCloud\Tom\Resource Planning\resourcing.xlsm"

    # Check if the Excel file exists
    if not os.path.exists(excel_file):
        print(f"Error: {excel_file} not found in the current directory")
        return
    
    # Create output directory if it doesn't exist
    output_dir = "output"
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        print(f"Created output directory: {output_dir}")
    
    # Generate datetime stamp for the database filename
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    db_filename = os.path.join(output_dir, f"resources_{timestamp}.db")
    
    print(f"Processing Excel file: {excel_file}")
    print(f"Creating SQLite database: {db_filename}")
    
    try:
        # Copy the source file to the output directory to avoid locked file issues
        source_filename = os.path.basename(excel_file)
        # Rename the copied file with the same timestamp as output files
        name_without_ext = os.path.splitext(source_filename)[0]
        extension = os.path.splitext(source_filename)[1]
        copied_excel_file = os.path.join(output_dir, f"{name_without_ext}_{timestamp}{extension}")
        shutil.copy2(excel_file, copied_excel_file)
        print(f"Copied source file to: {copied_excel_file}")
        
        # Read the Excel file to get worksheet names
        excel_data = pd.ExcelFile(copied_excel_file)
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
                df = pd.read_excel(copied_excel_file, sheet_name=worksheet_name)
                
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
        create_joined_output(copied_excel_file, db_filename, timestamp, output_dir, excel_file)
        
    except Exception as e:
        print(f"Error processing Excel file: {str(e)}")



def create_joined_output(excel_file, db_filename, timestamp, output_dir, original_source_file):
    """
    Create a joined output worksheet from the main tables including tblRates
    """
    try:
        # Connect to the database
        conn = sqlite3.connect(db_filename)
        
        # Create the joined query - left join from DW to UKG, PersonToTeam, Title Map, TeamToBacklog, and Rates
        join_query = """
        SELECT 
            dw.Employee_Number,
            ukg.*,
            pt.*,
            tm.Role,
            tb.*,
            r.*
        FROM "tblDW" dw
        LEFT JOIN "tblUKG" ukg ON dw.Employee_Number = ukg.Employee_Number
        LEFT JOIN "tblPersonToTeam" pt ON dw.Employee_Number = pt.Employee_Number
        LEFT JOIN "tblTitleMap" tm ON ukg.Business_Title = tm.Business_Title
        LEFT JOIN "tblTeamToBacklog" tb ON pt.Team = tb.Team
        LEFT JOIN "tblRates" r ON ukg.Location_Country = r.Location_Country
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
        
        # Note: Field configuration will be applied at the end to filter final output
        print("Field configuration will be applied to final output")
        
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
        expanded_df = add_calculated_columns(expanded_df, db_filename)
        
        # Save the expanded data to a table called "output_expanded" in the database
        expanded_df.to_sql("output_expanded", conn, if_exists='replace', index=False)
        conn.close()
        
        print(f"Saved expanded data to 'output_expanded' table in database")
        
        # Apply field configuration to filter final output columns
        try:
            # Check if tblFieldConfig worksheet exists
            field_config_df = pd.read_excel(excel_file, sheet_name='tblFieldConfig')
            print("Applying field configuration to final output...")
            
            # Read the field configuration to get the order
            field_config_df = pd.read_excel(excel_file, sheet_name='tblFieldConfig')
            field_config_df.columns = [str(col).replace(' ', '_').replace('-', '_').replace('(', '').replace(')', '') for col in field_config_df.columns]
            
            # Find the table and field column names
            table_col = None
            field_col = None
            for col in field_config_df.columns:
                col_lower = col.lower()
                if 'table' in col_lower or 'tablename' in col_lower:
                    table_col = col
                elif 'field' in col_lower or 'column' in col_lower or 'fieldname' in col_lower:
                    field_col = col
            
            if table_col and field_col:
                # Get columns to keep in config order
                columns_to_keep = []
                missing_fields = []
                
                for _, row in field_config_df.iterrows():
                    field_name = row[field_col]
                    if pd.isna(field_name):
                        continue
                    
                    # Look for column match (handle spaces vs underscores)
                    found = False
                    for col in expanded_df.columns:
                        # Normalize both names for comparison
                        normalized_field = field_name.lower().replace(' ', '_').replace('-', '_')
                        normalized_col = col.lower()
                        if normalized_field == normalized_col:
                            columns_to_keep.append(col)
                            found = True
                            break
                    
                    if not found:
                        missing_fields.append(field_name)
                
                # Fail if any fields are missing
                if missing_fields:
                    error_msg = f"ERROR: Field configuration contains fields that don't exist:\n"
                    error_msg += "\n".join([f"  - {field}" for field in missing_fields])
                    error_msg += f"\n\nAvailable columns: {sorted(expanded_df.columns.tolist())}"
                    raise ValueError(error_msg)
                
                # Filter to only the specified columns in the specified order, plus derived fields
                if columns_to_keep:
                    # Determine the maximum number of levels in the organization for dynamic mgr fields
                    max_levels = get_max_org_levels(db_filename)
                    mgr_fields = [f'mgr_{i}' for i in range(1, max_levels + 1)]
                    derived_fields = ['Year', 'YearMonth', 'Employee_Name'] + mgr_fields + ['Level_From_Top']
                    final_columns = columns_to_keep + [col for col in derived_fields if col in expanded_df.columns]
                    expanded_df = expanded_df[final_columns]
                    print(f"Filtered final output to {len(final_columns)} columns: {final_columns}")
        except Exception as e:
            print(f"No field configuration found or error reading it: {str(e)}")
            print("Proceeding with all columns in output")
        
        # Create a new Excel file with only the expanded data
        output_excel_file = os.path.join(output_dir, f"resources_output_{timestamp}.xlsx")
        
        # Write only the expanded dataset to the Excel file
        try:
            # Use a more robust Excel writing approach
            expanded_df.to_excel(output_excel_file, sheet_name='DATA', index=False, engine='openpyxl')
            print(f"Successfully created Excel file: {output_excel_file}")
        except Exception as excel_error:
            print(f"Error writing Excel file: {str(excel_error)}")
            # Fallback: try writing as CSV
            csv_file = os.path.join(output_dir, f"resources_output_{timestamp}.csv")
            expanded_df.to_csv(csv_file, index=False)
            print(f"Saved as CSV instead: {csv_file}")
        print(f"Expanded data has {len(expanded_df)} rows (added {len(expanded_df) - len(joined_df)} rows)")
        
        # Copy the DATA worksheet from the clean output file to the source file
        copy_data_worksheet_to_source(output_excel_file, original_source_file)
        
    except Exception as e:
        print(f"Error creating joined output: {str(e)}")

def expand_with_missing_dates(df):
    """
    Expand the DataFrame to create one record per month per unique combination of grouping dimensions.
    Each unique combination of dimensions (Employee, Team, Squad, etc.) gets records only for its valid date range.
    When an employee goes to 0% (leaves organization), stop creating ANY records for that employee.
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
        
        # Create a complete month range for the entire dataset
        all_months = pd.date_range(start=min_date.replace(day=1), 
                                  end=max_date.replace(day=1), 
                                  freq='MS')
        
        print(f"Creating records for {len(all_months)} months: {min_date.strftime('%Y-%m')} to {max_date.strftime('%Y-%m')}")
        
        expanded_rows = []
        
        # Get all columns except AsOfDate
        dimension_columns = [col for col in df_sorted.columns if col != 'AsOfDate']
        
        # Create a unique key for each row by combining all dimension values
        df_sorted['unique_key'] = df_sorted[dimension_columns].astype(str).agg('|'.join, axis=1)
        
        # Get unique combinations
        unique_combinations = df_sorted[['unique_key'] + dimension_columns].drop_duplicates()
        
        print(f"Found {len(unique_combinations)} unique dimension combinations")
        
        # For each unique combination
        for idx, combination in unique_combinations.iterrows():
            if idx % 50 == 0:  # Progress indicator every 50 combinations
                print(f"Processing combination {idx + 1}/{len(unique_combinations)}...")
            
            # Find all records that match this combination, sorted by date
            matching_records = df_sorted[df_sorted['unique_key'] == combination['unique_key']].copy().sort_values('AsOfDate')
            
            if len(matching_records) == 0:
                continue
            
            # Get the employee number for this combination
            employee_num = combination['Employee_Number']
            
            # Check if this employee has any 0% records (exit records)
            employee_all_records = df_sorted[df_sorted['Employee_Number'] == employee_num].copy().sort_values('AsOfDate')
            employee_exit_records = employee_all_records[employee_all_records['Percent'] == 0]
            
            # Find the earliest exit date for this employee
            earliest_exit_date = None
            if not employee_exit_records.empty:
                earliest_exit_date = employee_exit_records['AsOfDate'].min()
                print(f"Employee {employee_num} exits on {earliest_exit_date.date()}")
            
            # For each month in the complete range
            for month_start in all_months:
                # If this employee has exited before this month, skip (but allow the exit month itself)
                if earliest_exit_date and month_start > earliest_exit_date.replace(day=1):
                    continue
                
                # Find the most recent record before or on the end of this month for this combination
                month_end = month_start.replace(day=1) + pd.offsets.MonthEnd(0)
                records_before_or_on = matching_records[matching_records['AsOfDate'] <= month_end]
                
                if not records_before_or_on.empty:
                    # Get the most recent date's record
                    most_recent_record = records_before_or_on.iloc[-1]
                    most_recent_percent = most_recent_record['Percent']
                    
                    # Create a record for this month (including exit records with Percent = 0)
                    new_record = most_recent_record.copy()
                    new_record['AsOfDate'] = month_start
                    if 'unique_key' in new_record:
                        del new_record['unique_key']
                    expanded_rows.append(new_record.to_dict())
        
        # Create the expanded DataFrame
        expanded_df = pd.DataFrame(expanded_rows)
        
        # Sort by Employee_Number and AsOfDate
        expanded_df = expanded_df.sort_values(['Employee_Number', 'AsOfDate']).reset_index(drop=True)
        
        print(f"Created {len(expanded_df)} monthly records (vs {len(df_sorted)} original records)")
        print(f"Preserved all {len(unique_combinations)} unique dimension combinations")
        return expanded_df
        
    except Exception as e:
        print(f"Error expanding DataFrame with missing dates: {str(e)}")
        return df

def add_calculated_columns(df, db_filename=None):
    """
    Add calculated columns based on AsOfDate and Percent values.
    Now optimized for monthly records instead of daily.
    """
    try:
        # Make a copy to avoid modifying the original DataFrame
        df_copy = df.copy()
        
        # Ensure AsOfDate is datetime
        df_copy['AsOfDate'] = pd.to_datetime(df_copy['AsOfDate'])
        
        # Add Year column in YYYY format
        print("  - Adding Year column (YYYY format)...")
        df_copy['Year'] = df_copy['AsOfDate'].dt.strftime('%Y')
        
        # Add YearMonth column in YYYY-MM format
        print("  - Adding YearMonth column (YYYY_MM format)...")
        df_copy['YearMonth'] = df_copy['AsOfDate'].dt.strftime('%Y_%m')
        
        # Add Employee_Name column (lastname, firstname format)
        print("  - Adding Employee_Name column (lastname, firstname format)...")
        # Handle missing values by filling with empty string
        last_name = df_copy['Last_Name'].fillna('')
        first_name = df_copy['First_Name'].fillna('')
        df_copy['Employee_Name'] = last_name + ', ' + first_name
        # Clean up any cases where we have ", " with no name
        df_copy['Employee_Name'] = df_copy['Employee_Name'].str.replace('^, ', '', regex=True)
        df_copy['Employee_Name'] = df_copy['Employee_Name'].str.replace(', $', '', regex=True)
        
        # Add manager hierarchy columns if database filename is provided
        if db_filename:
            # Determine the maximum number of levels in the organization
            max_levels = get_max_org_levels(db_filename)
            print(f"  - Adding manager hierarchy columns (mgr_1 through mgr_{max_levels})...")
            
            for level in range(1, max_levels + 1):
                column_name = f'mgr_{level}'
                df_copy[column_name] = df_copy['Employee_Number'].apply(
                    lambda x: get_manager_at_level_from_top(x, level, db_filename) if pd.notna(x) else None
                )
            
            # Add employee level from top column
            print("  - Adding employee level from top column...")
            df_copy['Level_From_Top'] = df_copy['Employee_Number'].apply(
                lambda x: get_employee_level_from_top(x, db_filename) if pd.notna(x) else None
            )
            print(f"  - Successfully added {max_levels + 4} calculated columns to {len(df_copy)} rows")
        else:
            print(f"  - Successfully added 3 calculated columns to {len(df_copy)} rows")
        
        return df_copy
        
    except Exception as e:
        print(f"Error adding calculated columns: {str(e)}")
        return df


# Global cache for supervisor mapping
_supervisor_cache = None
_employee_names_cache = None


def _build_supervisor_cache(db_filename):
    """
    Build a cached mapping of employee number to supervisor number from the UKG table.
    This is called once and cached for performance.
    """
    global _supervisor_cache, _employee_names_cache
    
    if _supervisor_cache is not None:
        return  # Already cached
    
    try:
        # Connect to the database
        conn = sqlite3.connect(db_filename)
        
        # Get employee number to supervisor number mapping
        supervisor_query = """
        SELECT Employee_Number, Supervisor_Number 
        FROM tblUKG 
        WHERE Employee_Number IS NOT NULL AND Supervisor_Number IS NOT NULL
        """
        
        supervisor_df = pd.read_sql_query(supervisor_query, conn)
        
        # Get employee number to name mapping
        names_query = """
        SELECT Employee_Number, First_Name, Last_Name 
        FROM tblUKG 
        WHERE Employee_Number IS NOT NULL
        """
        
        names_df = pd.read_sql_query(names_query, conn)
        conn.close()
        
        # Build supervisor cache dictionary
        _supervisor_cache = {}
        for _, row in supervisor_df.iterrows():
            emp_num = row['Employee_Number']
            sup_num = row['Supervisor_Number']
            if pd.notna(emp_num) and pd.notna(sup_num):
                # Convert to clean string format (remove .0 if present)
                emp_str = str(int(emp_num)) if pd.notna(emp_num) else str(emp_num)
                sup_str = str(int(sup_num)) if pd.notna(sup_num) else str(sup_num)
                _supervisor_cache[emp_str] = sup_str
        
        # Build employee names cache dictionary
        _employee_names_cache = {}
        for _, row in names_df.iterrows():
            emp_num = row['Employee_Number']
            first_name = row['First_Name'] if pd.notna(row['First_Name']) else ""
            last_name = row['Last_Name'] if pd.notna(row['Last_Name']) else ""
            if pd.notna(emp_num):
                # Convert to clean string format (remove .0 if present)
                emp_str = str(int(emp_num)) if pd.notna(emp_num) else str(emp_num)
                full_name = f"{first_name} {last_name}".strip()
                _employee_names_cache[emp_str] = full_name if full_name else None
        
        print(f"Built supervisor cache with {len(_supervisor_cache)} employee-supervisor mappings")
        print(f"Built names cache with {len(_employee_names_cache)} employee names")
        
    except Exception as e:
        print(f"Error building supervisor cache: {str(e)}")
        _supervisor_cache = {}
        _employee_names_cache = {}


def get_manager_at_level_from_top(employee_number, level, db_filename):
    """
    Get the manager name at a specified level from the top of the org chart.
    mgr_1 = top executive (same for everyone), mgr_2 = level 2 managers, etc.
    Level represents distance from the top (1 = top, 2 = one level down, etc.)
    
    Args:
        employee_number: The employee number to look up
        level: Level from the top (1 = top executive, 2 = one level down, etc.)
        db_filename: The SQLite database filename
    
    Returns:
        The manager name at the specified level from the top, or None if not found
    """
    try:
        # Build cache if not already built
        _build_supervisor_cache(db_filename)
        
        # Validate inputs
        if level <= 0:
            return None
        
        # Convert to clean string format (remove .0 if present)
        employee_str = str(int(employee_number)) if pd.notna(employee_number) else str(employee_number)
        
        # Check if employee exists
        if employee_str not in _supervisor_cache and employee_str not in _employee_names_cache:
            return None
        
        # First, find the top of the org chart by traversing up until we can't go further
        current_employee = employee_str
        visited = set()  # To detect circular references
        path_to_top = [current_employee]
        
        while current_employee in _supervisor_cache:
            # Check for circular references
            if current_employee in visited:
                break
            
            visited.add(current_employee)
            current_employee = _supervisor_cache[current_employee]
            path_to_top.append(current_employee)
        
        # Reverse the path so we go from top down
        path_from_top = list(reversed(path_to_top))
        
        # Get the manager at the specified level from the top
        if level <= len(path_from_top):
            manager_employee = path_from_top[level - 1]  # level 1 = index 0
            if manager_employee in _employee_names_cache:
                return _employee_names_cache[manager_employee]
        
        return None
            
    except Exception as e:
        print(f"Error getting manager at level {level} from top: {str(e)}")
        return None


def get_max_org_levels(db_filename):
    """
    Find the maximum number of levels in the organization by checking all employees.
    
    Args:
        db_filename: The SQLite database filename
    
    Returns:
        The maximum number of levels in the organization
    """
    try:
        # Build cache if not already built
        _build_supervisor_cache(db_filename)
        
        max_levels = 0
        
        # Check all employees to find the deepest level
        for employee_str in _supervisor_cache.keys():
            current_employee = employee_str
            visited = set()
            level_count = 0
            
            while current_employee in _supervisor_cache:
                if current_employee in visited:
                    break
                
                visited.add(current_employee)
                current_employee = _supervisor_cache[current_employee]
                level_count += 1
            
            max_levels = max(max_levels, level_count)
        
        return max_levels
        
    except Exception as e:
        print(f"Error getting max org levels: {str(e)}")
        return 5  # Default fallback


def get_employee_level_from_top(employee_number, db_filename):
    """
    Get the number of levels an employee is down from the top supervisor.
    
    Args:
        employee_number: The employee number to look up
        db_filename: The SQLite database filename
    
    Returns:
        The number of levels down from the top, or None if not found
    """
    try:
        # Build cache if not already built
        _build_supervisor_cache(db_filename)
        
        # Convert to clean string format (remove .0 if present)
        employee_str = str(int(employee_number)) if pd.notna(employee_number) else str(employee_number)
        
        # Check if employee exists
        if employee_str not in _supervisor_cache and employee_str not in _employee_names_cache:
            return None
        
        current_employee = employee_str
        level_count = 0
        visited = set()  # To detect circular references
        
        # Traverse up the supervisor chain until we reach the top
        while current_employee in _supervisor_cache:
            # Check for circular references
            if current_employee in visited:
                return None
            
            visited.add(current_employee)
            current_employee = _supervisor_cache[current_employee]
            level_count += 1
        
        return level_count
        
    except Exception as e:
        print(f"Error getting employee level from top: {str(e)}")
        return None







if __name__ == "__main__":
    process_resource_data()
