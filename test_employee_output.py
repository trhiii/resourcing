import pandas as pd

# Read the output Excel file
output_file = 'output/resources_output_20250912_152836.xlsx'
df = pd.read_excel(output_file)

# Filter for employee 20010207
emp_20010207 = df[df['Employee_Number'] == 20010207].copy()

print('=== Employee 20010207 Output Data ===')
print(f'Total records: {len(emp_20010207)}')

# Group by Team and show date ranges and percentages
print('\n=== Team Assignments by Date Range ===')
team_summary = emp_20010207.groupby('Team').agg({
    'AsOfDate': ['min', 'max'],
    'Percent': ['min', 'max', 'mean'],
    'Allocation': 'sum'
}).round(4)

print(team_summary)

# Check monthly totals
print('\n=== Monthly Totals (should be 100% each month) ===')
monthly_totals = emp_20010207.groupby('YearMonth')['Percent'].sum().round(2)
print(monthly_totals)

# Show sample records for each team
print('\n=== Sample Records by Team ===')
for team in emp_20010207['Team'].unique():
    team_data = emp_20010207[emp_20010207['Team'] == team].head(3)
    print(f'\n--- {team} ---')
    print(team_data[['AsOfDate', 'Percent', 'Allocation']].to_string(index=False))

# Check for any overlapping periods
print('\n=== Checking for Overlapping Periods ===')
overlaps = []
teams = emp_20010207['Team'].unique()
for i, team1 in enumerate(teams):
    for team2 in teams[i+1:]:
        team1_data = emp_20010207[emp_20010207['Team'] == team1]
        team2_data = emp_20010207[emp_20010207['Team'] == team2]
        
        # Check if date ranges overlap
        team1_start = team1_data['AsOfDate'].min()
        team1_end = team1_data['AsOfDate'].max()
        team2_start = team2_data['AsOfDate'].min()
        team2_end = team2_data['AsOfDate'].max()
        
        if not (team1_end < team2_start or team2_end < team1_start):
            overlaps.append(f'{team1} ({team1_start.date()} to {team1_end.date()}) overlaps with {team2} ({team2_start.date()} to {team2_end.date()})')

if overlaps:
    print('Found overlapping periods:')
    for overlap in overlaps:
        print(f'  - {overlap}')
else:
    print('No overlapping periods found - good!')

