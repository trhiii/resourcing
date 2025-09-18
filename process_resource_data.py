import pandas as pd 
import sqlite3
from datetime import datetime
import os
import shutil
import platform

def get_source_file_path():
    """
    Get the correct source file path based on the operating system.
    Returns the path to the resourcing.xlsm file.
    """
    system = platform.system()
    
    if system == "Darwin":  # macOS
        # macOS path provided by user
        base_path = "/Users/tom/Library/CloudStorage/OneDrive-InvestCloud/Tom/Resource Planning"
        excel_file = os.path.join(base_path, "resourcing.xlsm")
    elif system == "Windows":
        # Windows path (original)
        base_path = r"C:\Users\thockswender\OneDrive - InvestCloud\Tom\Resource Planning"
        excel_file = os.path.join(base_path, "resourcing.xlsm")
    else:
        # Fallback for other systems (Linux, etc.)
        print(f"Warning: Unsupported operating system '{system}'. Using default path.")
        base_path = "/Users/tom/Library/CloudStorage/OneDrive-InvestCloud/Tom/Resource Planning"
        excel_file = os.path.join(base_path, "resourcing.xlsm")
    
    # Check if the base directory exists
    if not os.path.exists(base_path):
        print(f"Error: Base directory not found: {base_path}")
        print(f"Please ensure the OneDrive directory is accessible on {system}")
        return None
    
    # Check if the file exists
    if not os.path.exists(excel_file):
        print(f"Error: Source file not found at: {excel_file}")
        print(f"Please ensure the file exists at the correct location for {system}")
        print(f"Expected filename: resourcing.xlsm")
        return None
    
    # Check if the file is readable
    if not os.access(excel_file, os.R_OK):
        print(f"Error: Source file is not readable: {excel_file}")
        print(f"Please check file permissions on {system}")
        return None
    
    print(f"Using source file: {excel_file}")
    return excel_file

def print_platform_help():
    """
    Print helpful information about file locations for different platforms.
    """
    system = platform.system()
    print("\n" + "=" * 60)
    print("PLATFORM-SPECIFIC FILE LOCATION HELP")
    print("=" * 60)
    
    if system == "Darwin":  # macOS
        print("macOS (Darwin) detected:")
        print("Expected file location: /Users/tom/Library/CloudStorage/OneDrive-InvestCloud/Tom/Resource Planning/resourcing.xlsm")
        print("If the file is not found:")
        print("1. Check if OneDrive is synced and accessible")
        print("2. Verify the file exists in the OneDrive folder")
        print("3. Check file permissions (should be readable)")
    elif system == "Windows":
        print("Windows detected:")
        print("Expected file location: C:\\Users\\thockswender\\OneDrive - InvestCloud\\Tom\\Resource Planning\\resourcing.xlsm")
        print("If the file is not found:")
        print("1. Check if OneDrive is synced and accessible")
        print("2. Verify the file exists in the OneDrive folder")
        print("3. Check file permissions (should be readable)")
    else:
        print(f"Unsupported platform: {system}")
        print("Please ensure the file is accessible and update the path in get_source_file_path() function")
    
    print("=" * 60)

def create_database_from_excel(excel_file, output_dir, timestamp):
    """
    Read the Excel file and create a SQLite database with worksheets starting with 'tbl'
    Returns the database filename and copied Excel file path
    """
    # Check if the Excel file exists
    if not os.path.exists(excel_file):
        print(f"Error: {excel_file} not found in the current directory")
        return None, None
    
    # Create output directory if it doesn't exist
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        print(f"Created output directory: {output_dir}")
    
    # Generate datetime stamp for the database filename
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
        
        # Append tblTBH data to tblUKG if both tables exist
        try:
            # Check if both tables exist
            cursor = conn.cursor()
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name IN ('tblUKG', 'tblTBH');")
            existing_tables = [row[0] for row in cursor.fetchall()]
            
            if 'tblUKG' in existing_tables and 'tblTBH' in existing_tables:
                print("Appending tblTBH data to tblUKG...")
                
                # Get the current row count in tblUKG
                cursor.execute("SELECT COUNT(*) FROM tblUKG")
                ukg_before_count = cursor.fetchone()[0]
                
                # Get the row count in tblTBH
                cursor.execute("SELECT COUNT(*) FROM tblTBH")
                tbh_count = cursor.fetchone()[0]
                
                # Append tblTBH to tblUKG
                cursor.execute("INSERT INTO tblUKG SELECT * FROM tblTBH")
                conn.commit()
                
                # Get the new row count in tblUKG
                cursor.execute("SELECT COUNT(*) FROM tblUKG")
                ukg_after_count = cursor.fetchone()[0]
                
                print(f"  - Appended {tbh_count} rows from tblTBH to tblUKG")
                print(f"  - tblUKG now has {ukg_after_count} rows (was {ukg_before_count})")
            else:
                print("Skipping tblTBH append - one or both tables missing")
                
        except Exception as e:
            print(f"  - Error appending tblTBH to tblUKG: {str(e)}")
        
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
        
        return db_filename, copied_excel_file
        
    except Exception as e:
        print(f"Error processing Excel file: {str(e)}")
        return None, None


def create_joined_dataframe(db_filename):
    """
    Create a joined DataFrame from the main tables
    Returns the joined DataFrame
    """
    try:
        # Connect to the database
        conn = sqlite3.connect(db_filename)
        
        # Create the joined query - left join from PersonToTeam to UKG, TeamToBacklog, Rates, and RoleOverride
        # Note: Role will be added programmatically after the join with override logic
        join_query = """
        SELECT 
            pt.Employee_Number,
            ukg.*,
            pt.*,
            tb.*,
            r.*,
            ro.Role as Override_Role
        FROM "tblPersonToTeam" pt
        LEFT JOIN "tblUKG" ukg ON pt.Employee_Number = ukg.Employee_Number
        LEFT JOIN "tblTeamToBacklog" tb ON pt.Team = tb.Team
        LEFT JOIN "tblRates" r ON ukg.Location_Country = r.Location_Country
        LEFT JOIN "tblRoleOverride" ro ON pt.Employee_Number = ro.Employee_Number
        ORDER BY pt.Employee_Number
        """
        
        # First, get the count of rows in tblPersonToTeam and all employee numbers to verify we don't lose any
        pt_count_query = 'SELECT COUNT(*) as count FROM "tblPersonToTeam"'
        pt_count = pd.read_sql_query(pt_count_query, conn)
        pt_row_count = pt_count.iloc[0]['count']
        print(f"tblPersonToTeam has {pt_row_count} rows")
        
        # Get all employee numbers from tblPersonToTeam
        pt_employee_numbers = pd.read_sql_query('SELECT DISTINCT Employee_Number FROM "tblPersonToTeam"', conn)
        
        # Execute the query and get the results
        joined_df = pd.read_sql_query(join_query, conn)
        
        # Add Role column programmatically using tblTitleMap lookup with override logic
        print("Adding Role column programmatically with override logic...")
        role_mapping_df = pd.read_sql_query('SELECT Business_Title, Role FROM "tblTitleMap"', conn)
        role_mapping = dict(zip(role_mapping_df['Business_Title'], role_mapping_df['Role']))
        
        # First map Business_Title to Role, defaulting to "?" if not found
        joined_df['Role'] = joined_df['Business_Title'].map(role_mapping).fillna('?')
        
        # Apply role overrides where available (non-null Override_Role)
        override_mask = joined_df['Override_Role'].notna()
        joined_df.loc[override_mask, 'Role'] = joined_df.loc[override_mask, 'Override_Role']
        
        # Drop the temporary Override_Role column
        joined_df = joined_df.drop('Override_Role', axis=1)
        
        override_count = override_mask.sum()
        print(f"Role mapping applied: {len(role_mapping)} titles mapped, {joined_df['Role'].isna().sum()} records defaulted to '?'")
        print(f"Role overrides applied: {override_count} records overridden from tblRoleOverride")
        
        conn.close()
        
        print(f"Joined data created with {len(joined_df)} rows")
        print(f"Joined DataFrame columns: {joined_df.columns.tolist()}")
        
        # Verify that all tblPersonToTeam Employee_Number values are present in the joined output
        # Use the first Employee_Number column (from tblPersonToTeam)
        joined_employee_numbers = joined_df.iloc[:, 0].dropna().unique()
        
        missing_employees = set(pt_employee_numbers['Employee_Number'].dropna().tolist()) - set(joined_employee_numbers)
        if missing_employees:
            print(f"ERROR: Missing {len(missing_employees)} employees from tblPersonToTeam in joined output!")
            print(f"Missing Employee Numbers: {sorted(missing_employees)}")
        else:
            print("✓ All tblPersonToTeam employees preserved in joined output")
        
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
        
        return joined_df
        
    except Exception as e:
        print(f"Error creating joined DataFrame: {str(e)}")
        return None


def expand_dataframe_with_dates(joined_df):
    """
    Expand the DataFrame with daily records using fill-down logic
    Returns the expanded DataFrame
    """
    print("\nExpanding DataFrame with daily records...")
    expanded_df = expand_with_missing_dates(joined_df)
    return expanded_df


def add_calculated_fields(expanded_df, db_filename):
    """
    Add calculated columns to the expanded DataFrame
    Returns the DataFrame with calculated fields
    """
    print("\nAdding calculated columns to expanded data...")
    final_df = add_calculated_columns(expanded_df, db_filename)
    return final_df


def apply_field_configuration(final_df, excel_file, db_filename):
    """
    Apply field configuration to filter final output columns
    Returns the filtered DataFrame
    """
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
                for col in final_df.columns:
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
                error_msg += f"\n\nAvailable columns: {sorted(final_df.columns.tolist())}"
                raise ValueError(error_msg)
            
            # Filter to only the specified columns in the specified order, plus derived fields
            if columns_to_keep:
                # Determine the maximum number of levels in the organization for dynamic mgr fields
                max_levels = get_max_org_levels(db_filename)
                mgr_fields = [f'mgr_{i}' for i in range(1, max_levels + 1)]
                derived_fields = ['Year', 'YearMonth', 'Sprint', 'Sprint_Allocation', 'RunDate', 'Employee_Name', 'Allocation'] + mgr_fields + ['Level_From_Top']
                final_columns = columns_to_keep + [col for col in derived_fields if col in final_df.columns]
                final_df = final_df[final_columns]
                print(f"Filtered final output to {len(final_columns)} columns: {final_columns}")
        return final_df
    except Exception as e:
        print(f"No field configuration found or error reading it: {str(e)}")
        print("Proceeding with all columns in output")
        return final_df


def create_output_files(final_df, output_dir, timestamp, original_source_file):
    """
    Create the final output files
    """
    # Create a new Excel file with only the expanded data
    output_excel_file = os.path.join(output_dir, f"resources_output_{timestamp}.xlsx")
    
    # Write only the expanded dataset to the Excel file
    try:
        # Use a more robust Excel writing approach with proper datetime formatting
        with pd.ExcelWriter(output_excel_file, engine='openpyxl', datetime_format='YYYY-MM-DD') as writer:
            final_df.to_excel(writer, sheet_name='DATA', index=False)
        print(f"Successfully created Excel file: {output_excel_file}")
    except Exception as excel_error:
        print(f"Error writing Excel file: {str(excel_error)}")
        # Fallback: try writing as CSV
        csv_file = os.path.join(output_dir, f"resources_output_{timestamp}.csv")
        final_df.to_csv(csv_file, index=False)
        print(f"Saved as CSV instead: {csv_file}")
    
    # Create OUTPUT file in the source directory
    source_dir = os.path.dirname(original_source_file)
    output_file_in_source_dir = os.path.join(source_dir, "OUTPUT.xlsx")
    
    # Copy the clean output file to the source directory as OUTPUT.xlsx
    shutil.copy2(output_excel_file, output_file_in_source_dir)
    print(f"Created OUTPUT file in source directory: {output_file_in_source_dir}")


def process_resource_data():
    """
    Main function that orchestrates the entire data processing pipeline
    """
    # Get the correct Excel file path based on the operating system
    excel_file = get_source_file_path()
    if not excel_file:
        print("ERROR: Could not locate source file. Exiting.")
        print_platform_help()
        return
    
    # Create output directory
    output_dir = "output"
    
    # Generate datetime stamp for all output files
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    
    print("=" * 60)
    print("RESOURCE DATA PROCESSING PIPELINE")
    print(f"Running on: {platform.system()} {platform.release()}")
    print("=" * 60)
    
    # Step 1: Create database from Excel
    print("\nSTEP 1: Creating database from Excel file...")
    db_filename, copied_excel_file = create_database_from_excel(excel_file, output_dir, timestamp)
    if not db_filename:
        print("ERROR: Failed to create database. Exiting.")
        return
    
    # Step 2: Create joined DataFrame
    print("\nSTEP 2: Creating joined DataFrame...")
    joined_df = create_joined_dataframe(db_filename)
    if joined_df is None:
        print("ERROR: Failed to create joined DataFrame. Exiting.")
        return
    
    # Step 3: Expand DataFrame with daily records
    print("\nSTEP 3: Expanding DataFrame with daily records...")
    expanded_df = expand_dataframe_with_dates(joined_df)
    
    # Step 4: Add calculated fields
    print("\nSTEP 4: Adding calculated fields...")
    final_df = add_calculated_fields(expanded_df, db_filename)
    
    # Step 5: Apply field configuration
    print("\nSTEP 5: Applying field configuration...")
    final_df = apply_field_configuration(final_df, copied_excel_file, db_filename)
    
    # Step 6: Create output files
    print("\nSTEP 6: Creating output files...")
    create_output_files(final_df, output_dir, timestamp, excel_file)
    
    print(f"\nFinal data has {len(final_df)} rows")
    print("=" * 60)
    print("PROCESSING COMPLETED SUCCESSFULLY!")
    print("=" * 60)




def expand_with_missing_dates(df):
    """
    Expand the DataFrame to create one record per day per unique combination of grouping dimensions.
    The repeating group dimension is Employee + Group + Subgroup + Team (excluding Percent and AsOfDate).
    For each unique team assignment, track how the percentage changes over time and create daily records
    with the appropriate percentage for each period.
    """
    try:
        # Ensure the DataFrame is sorted by grouping dimensions and AsOfDate
        df_sorted = df.sort_values(['Employee_Number', 'Group', 'Subgroup', 'Team', 'AsOfDate']).reset_index(drop=True)
        
        # Convert AsOfDate to datetime if it's not already
        df_sorted['AsOfDate'] = pd.to_datetime(df_sorted['AsOfDate'])
        
        # Get the date range for the entire dataset
        min_date = df_sorted['AsOfDate'].min()
        max_date = df_sorted['AsOfDate'].max()
        
        print(f"Date range: {min_date.date()} to {max_date.date()}")
        print(f"Total days: {(max_date - min_date).days + 1}")
        
        expanded_rows = []
        
        # Get all unique team assignments (Employee + Group + Subgroup + Team)
        grouping_cols = ['Employee_Number', 'Group', 'Subgroup', 'Team']
        unique_team_assignments = df_sorted[grouping_cols].drop_duplicates()
        print(f"Processing {len(unique_team_assignments)} unique team assignments...")
        
        # For each unique team assignment
        for assignment_idx, (_, assignment_row) in enumerate(unique_team_assignments.iterrows()):
            if assignment_idx % 50 == 0:  # Progress indicator every 50 assignments
                emp_num = assignment_row['Employee_Number']
                team = assignment_row['Team']
                print(f"Processing assignment {assignment_idx + 1}/{len(unique_team_assignments)}: Employee {emp_num}, Team {team}")
            
            # Create a mask for this specific team assignment
            mask = True
            for col in grouping_cols:
                mask = mask & (df_sorted[col] == assignment_row[col])
            
            # Get all records for this team assignment, sorted by AsOfDate
            team_assignment_records = df_sorted[mask].copy().sort_values('AsOfDate')
            
            if len(team_assignment_records) == 0:
                continue
            
            # Get the employee number for this assignment
            employee_num = assignment_row['Employee_Number']
            
            # Get ALL AsOfDates for this employee across ALL team assignments
            employee_all_dates = df_sorted[df_sorted['Employee_Number'] == employee_num]['AsOfDate'].unique()
            employee_all_dates = sorted(employee_all_dates)
            
            # Process each period for this team assignment
            for period_idx, (_, period_record) in enumerate(team_assignment_records.iterrows()):
                period_start_date = period_record['AsOfDate']
                period_percent = period_record['Percent']
                
                # Find the end date for this period
                period_start_idx = list(employee_all_dates).index(period_start_date)
                if period_start_idx < len(employee_all_dates) - 1:
                    # There's a next AsOfDate - this period ends the day before it
                    next_asof_date = employee_all_dates[period_start_idx + 1]
                    period_end_date = next_asof_date - pd.Timedelta(days=1)
                else:
                    # This is the last AsOfDate for this employee - continue to dataset max
                    period_end_date = max_date
                
                # Only create daily records if the period has a valid date range and non-zero percentage
                if period_start_date <= period_end_date and period_percent > 0:
                    # Create daily date range for this period
                    period_dates = pd.date_range(start=period_start_date, end=period_end_date, freq='D')
                    
                    # For each day in this period, create a record with the period's percentage
                    for current_date in period_dates:
                        # Create a new record for this date with the team assignment data
                        new_record = period_record.copy()
                        new_record['AsOfDate'] = current_date
                        new_record['Percent'] = period_percent
                        expanded_rows.append(new_record.to_dict())
        
        # Create the expanded DataFrame
        expanded_df = pd.DataFrame(expanded_rows)
        
        # Ensure AsOfDate is datetime type
        expanded_df['AsOfDate'] = pd.to_datetime(expanded_df['AsOfDate'])
        
        # Sort by Employee_Number, Team, and AsOfDate
        expanded_df = expanded_df.sort_values(['Employee_Number', 'Team', 'AsOfDate']).reset_index(drop=True)
        
        print(f"Created {len(expanded_df)} daily records (vs {len(df_sorted)} original records)")
        print(f"Expansion ratio: {len(expanded_df) / len(df_sorted):.1f}x")
        return expanded_df
        
    except Exception as e:
        print(f"Error expanding DataFrame with missing dates: {str(e)}")
        return df

def get_sprint_info(asof_date):
    """
    Calculate sprint information based on AsOfDate.
    Sprint starts on Wednesday, January 1, 2025 and each sprint is 14 days.
    Returns sprint text in format "Sprint ##, DD-MMM" where DD-MMM is the Tuesday end date.
    """
    try:
        # Convert to datetime if not already
        if not isinstance(asof_date, pd.Timestamp):
            asof_date = pd.to_datetime(asof_date)
        
        # Sprint start date: Wednesday, January 1, 2025
        sprint_start = pd.Timestamp('2025-01-01')
        
        # Calculate days since sprint start
        days_since_start = (asof_date - sprint_start).days
        
        # If before sprint start, return None or handle as needed
        if days_since_start < 0:
            return None
        
        # Calculate sprint number (1-based, each sprint is 14 days)
        sprint_number = (days_since_start // 14) + 1
        
        # Cap at 26 sprints per year (26 * 14 = 364 days, leaving 1 day for year end)
        sprint_number = min(sprint_number, 26)
        
        # Calculate the end date of this sprint (Tuesday, 13 days after sprint start)
        sprint_end_date = sprint_start + pd.Timedelta(days=(sprint_number - 1) * 14 + 13)
        
        # Format the end date as DD-MMM
        end_date_formatted = sprint_end_date.strftime('%d-%b')
        
        # Format as "Sprint ##, DD-MMM"
        return f"Sprint {sprint_number:02d}, {end_date_formatted}"
        
    except Exception as e:
        print(f"Error calculating sprint for date {asof_date}: {str(e)}")
        return None


def add_calculated_columns(df, db_filename=None):
    """
    Add calculated columns based on AsOfDate and Percent values.
    Now optimized for daily records.
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
        
        # Add Sprint column
        print("  - Adding Sprint column (Sprint ##, DD-MMM format)...")
        df_copy['Sprint'] = df_copy['AsOfDate'].apply(get_sprint_info)
        
        # Add Sprint Allocation column (always 1/14)
        print("  - Adding Sprint Allocation column (1/14)...")
        df_copy['Sprint_Allocation'] = 1.0 / 14.0
        
        # Add RunDate column (datetime when script is run)
        print("  - Adding RunDate column (script execution datetime)...")
        df_copy['RunDate'] = datetime.now()
        
        # Add Allocation column (daily portion of monthly percentage)
        print("  - Adding Allocation column (Percent / days_in_month for daily records)...")
        # Calculate days in month for each date
        df_copy['days_in_month'] = df_copy['AsOfDate'].dt.days_in_month
        # Allocation = daily portion of the monthly percentage
        df_copy['Allocation'] = df_copy['Percent'] / df_copy['days_in_month']
        
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
            print(f"  - Successfully added {max_levels + 7} calculated columns to {len(df_copy)} rows")
        else:
            print(f"  - Successfully added 6 calculated columns to {len(df_copy)} rows")
        
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
