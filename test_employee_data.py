import sqlite3
import pandas as pd

# Connect to the most recent database
db_file = 'output/resources_20250912_140723.db'
conn = sqlite3.connect(db_file)

# Check what tables exist
tables = pd.read_sql_query("SELECT name FROM sqlite_master WHERE type='table'", conn)
print('Available tables:')
print(tables)

# Look at the structure of tblPersonToTeam
print('\n=== tblPersonToTeam structure ===')
structure = pd.read_sql_query('PRAGMA table_info(tblPersonToTeam)', conn)
print(structure)

# Get all records for employee 20010207
print('\n=== Employee 20010207 records ===')
emp_data = pd.read_sql_query('SELECT * FROM tblPersonToTeam WHERE Employee_Number = 20010207 ORDER BY AsOfDate', conn)
print(emp_data)

# Show the detailed breakdown
print('\n=== Detailed breakdown for Employee 20010207 ===')
for _, row in emp_data.iterrows():
    print(f"AsOfDate: {row['AsOfDate']}, Group: {row['Group']}, Subgroup: {row['Subgroup']}, Team: {row['Team']}, Percent: {row['Percent']}")

# Let's also check what the unique combinations look like for this employee
print('\n=== Unique combinations for Employee 20010207 ===')
unique_combos = emp_data.groupby(['Employee_Number', 'Group', 'Subgroup', 'Team', 'AsOfDate']).size().reset_index(name='count')
print(unique_combos)

# Check if there are multiple records with same Employee+Group+Subgroup+Team but different AsOfDate
print('\n=== Records by Employee+Group+Subgroup+Team (ignoring AsOfDate) ===')
team_assignments = emp_data.groupby(['Employee_Number', 'Group', 'Subgroup', 'Team']).agg({
    'AsOfDate': ['min', 'max', 'count'],
    'Percent': ['min', 'max', 'mean']
}).round(2)
print(team_assignments)

conn.close()
