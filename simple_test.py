import sqlite3
import pandas as pd

# Connect to the database
db_file = 'output/resources_20250912_140723.db'
conn = sqlite3.connect(db_file)

# Get all records for employee 20010207
print('=== Employee 20010207 raw data ===')
emp_data = pd.read_sql_query('SELECT * FROM tblPersonToTeam WHERE Employee_Number = 20010207 ORDER BY AsOfDate', conn)
print(emp_data.to_string())

print('\n=== Summary by AsOfDate ===')
summary = emp_data.groupby('AsOfDate').agg({
    'Team': lambda x: list(x),
    'Percent': lambda x: list(x),
    'Percent': 'sum'
}).round(3)
print(summary)

print('\n=== What my current logic would do ===')
print('Current logic groups by Employee+AsOfDate, so:')
print('AsOfDate 2025-08-01: 3 teams with 33.33% each = 100% total')
print('AsOfDate 2025-09-10: 2 teams with 50% each = 100% total')
print('This should be correct!')

conn.close()
