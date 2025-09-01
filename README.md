# Resource Data Processing System

A comprehensive system for processing and validating resource allocation data, with support for employee lifecycle management, team transitions, fractional allocations, and organizational hierarchy analysis.

## Overview

This system processes Excel-based resource data and creates monthly-expanded datasets suitable for crosstab/pivot table analysis. It handles:

- **Employee Lifecycle Management**: Hire, team changes, fractional allocations, and exits
- **Transaction-Based Processing**: Each record represents a change/transaction
- **Monthly Expansion**: Creates one record per month per unique dimension combination
- **Organizational Hierarchy**: Manager chain analysis with dynamic level detection
- **Field Configuration**: Configurable output columns via Excel worksheet
- **Data Validation**: Comprehensive validation for over-allocation and data quality issues

## Key Features

### ✅ **Employee Lifecycle Support**
- **Hiring**: New employees added to teams
- **Team Transitions**: Movement between teams/groups
- **Fractional Allocations**: Employees working across multiple teams
- **Exits**: Proper handling when employees leave (Percent = 0)

### ✅ **Organizational Hierarchy Analysis**
- **Dynamic Manager Levels**: Automatically detects organizational depth (up to 8 levels)
- **Manager Chain Fields**: `mgr_1` through `mgr_8` representing distance from top
- **Top-Down Analysis**: `mgr_1` = top executive (same for everyone)
- **Level Detection**: `Level_From_Top` shows employee's depth in organization
- **Cached Performance**: Optimized lookups for large datasets

### ✅ **Field Configuration System**
- **Configurable Output**: Use `tblFieldConfig` worksheet to specify output columns
- **Order Preservation**: Maintains exact column order from configuration
- **Strict Validation**: Fails if configured fields don't exist
- **Derived Fields**: Always includes calculated fields (Year, YearMonth, Employee_Name, mgr_X, Level_From_Top)

### ✅ **Data Integrity**
- **Over-Allocation Detection**: Identifies employees allocated >100%
- **Transaction Accuracy**: Respects actual change dates
- **Exit Handling**: Stops creating records after employee exits
- **Dimension Preservation**: Maintains all grouping dimensions

### ✅ **Performance Optimized**
- **Monthly-Level Processing**: Much more efficient than daily records
- **Reduced File Sizes**: Significantly smaller output files
- **Crosstab Ready**: Perfect for monthly analysis and reporting
- **Output Organization**: Timestamped files in dedicated output directory

## File Structure

```
├── process_resource_data.py      # Main processing script
├── test_process_resource_data.py # Comprehensive test suite
├── validate_production_data.py   # Production data validation
├── requirements.txt              # Python dependencies
├── README.md                     # This file
└── output/                       # Generated output files
    ├── resources_YYYYMMDD_HHMMSS.xlsx    # Processed Excel output
    ├── resources_YYYYMMDD_HHMMSS.db      # SQLite database
    └── resources 1.0_YYYYMMDD_HHMMSS.xlsm # Copied source file
```

## Installation

1. **Clone the repository**:
   ```bash
   git clone <repository-url>
   cd resourcing
   ```

2. **Create virtual environment**:
   ```bash
   python -m venv venv
   ```

3. **Activate virtual environment**:
   ```bash
   # Windows
   venv\Scripts\Activate.ps1
   
   # macOS/Linux
   source venv/bin/activate
   ```

4. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

### Processing Resource Data

```bash
python process_resource_data.py
```

This will:
1. Create an `output` directory if it doesn't exist
2. Copy the source Excel file to avoid locked file issues
3. Process all worksheets starting with 'tbl'
4. Create a SQLite database with joined data
5. Expand the data to monthly records
6. Add calculated columns including manager hierarchy
7. Apply field configuration if `tblFieldConfig` exists
8. Generate timestamped output files

### Running Tests

```bash
# Run all tests
python test_process_resource_data.py

# Run with verbose output
python -m unittest test_process_resource_data -v
```

### Validating Production Data

```bash
python validate_production_data.py
```

This validates the latest production data and reports:
- Data quality issues
- Over-allocation patterns
- Employee exit records
- Overall data statistics

## Manager Hierarchy System

### **Manager Fields**
- **`mgr_1`**: Top executive (same for all employees)
- **`mgr_2`**: Level 2 managers (direct reports to mgr_1)
- **`mgr_3`**: Level 3 managers (direct reports to mgr_2)
- **...**: Continues to deepest organizational level
- **`Level_From_Top`**: Number of levels employee is down from top

### **Example Hierarchy**
```
mgr_1: Jeffery Yabuki (CEO - same for everyone)
mgr_2: James Young (CTO - same for all tech employees)
mgr_3: Ian Peckett (VP Engineering - same for engineering)
mgr_4: [Department Head]
mgr_5: [Team Lead]
Level_From_Top: 5 (for individual contributors)
```

### **Dynamic Level Detection**
- Automatically detects organizational depth
- Creates appropriate number of mgr fields (1-8 levels)
- Handles circular references and missing data gracefully
- Optimized with caching for performance

## Field Configuration

### **Configuration Worksheet**
Create a worksheet named `tblFieldConfig` in your source Excel file:

| Table_Name | Field_Name |
|------------|------------|
| tblUKG | Employee_Number |
| tblUKG | First_Name |
| tblUKG | Last_Name |
| tblPersonToTeam | Team |
| tblPersonToTeam | Percent |

### **Configuration Rules**
- **Strict Matching**: Field names must exactly match (case-insensitive)
- **Order Preservation**: Output columns appear in configuration order
- **Derived Fields**: Always included regardless of configuration
- **Validation**: Process fails if any configured field is missing

### **Derived Fields (Always Included)**
- `Year`: Year from AsOfDate (YYYY format)
- `YearMonth`: Year-Month from AsOfDate (YYYY_MM format)
- `Employee_Name`: Concatenated "Last_Name, First_Name"
- `mgr_1` through `mgr_8`: Manager hierarchy fields
- `Level_From_Top`: Employee's level in organization

## Test Suite

The test suite covers all critical scenarios:

### 🧪 **Core Functionality Tests**
- **Basic Functionality**: Simple data processing
- **Calculated Columns**: Year, YearMonth, Employee_Name
- **Manager Hierarchy**: mgr_X fields and Level_From_Top
- **Data Validation**: Over-allocation detection

### 🧪 **Employee Lifecycle Tests**
- **Complete Lifecycle**: Hire → Team Changes → Fractional → Exit
- **Employee Exit Handling**: Proper exit date handling
- **Fractional Allocation**: 60% Team A + 40% Team B = 100%

### 🧪 **Data Quality Tests**
- **Over-Allocation Detection**: Identifies >100% allocations
- **Validation Function**: Comprehensive data quality checks
- **Edge Cases**: Single records, immediate exits

### 🧪 **Test Scenarios Covered**

1. **Employee Addition**: New person joins a team
2. **Team Movement**: Employee moves between teams
3. **Fractional Allocation**: Employee works across multiple teams
4. **Allocation Reduction**: Back to single team assignment
5. **Employee Exit**: Person leaves the organization
6. **Over-Allocation**: Employee allocated >100% (allowed but flagged)

## Data Structure

### Input Data (Excel)
- **tblDW**: Employee master data
- **tblUKG**: HR system data (includes supervisor relationships)
- **tblPersonToTeam**: Team assignments
- **tblTitleMap**: Role mappings
- **tblTeamToBacklog**: Team metadata
- **tblFieldConfig**: Output field configuration (optional)

### Output Data
- **Monthly Records**: One record per month per unique combination
- **Calculated Columns**: Year, YearMonth, Employee_Name, mgr_X, Level_From_Top
- **Preserved Dimensions**: All original grouping dimensions maintained
- **Manager Hierarchy**: Complete organizational chain analysis

## Output File Organization

### **Timestamped Files**
All output files share the same timestamp for easy identification:
- `resources_20250901_143022.xlsx` - Main processed output
- `resources_20250901_143022.db` - SQLite database
- `resources 1.0_20250901_143022.xlsm` - Copied source file

### **Output Directory**
- Automatically created if it doesn't exist
- Contains all generated files
- Prevents file locking issues
- Organized by processing session

## Validation Rules

### ✅ **Data Quality Checks**
- No negative percentages
- Required columns present
- Valid date ranges
- Employee number consistency
- Manager hierarchy integrity

### ⚠️ **Over-Allocation Detection**
- **Allowed**: Employees can be allocated >100%
- **Flagged**: System detects and reports over-allocation
- **Use Case**: Temporary assignments, special projects

### 🚫 **Exit Handling**
- **0% Records**: Mark employee exits
- **No Forward Records**: Stop creating records after exit
- **Clean Data**: No phantom records for exited employees

## Performance Metrics

### **File Size Optimization**
- **Efficient Processing**: Monthly vs daily records
- **Organized Output**: Timestamped files in dedicated directory
- **Cached Lookups**: Manager hierarchy optimized with caching
- **Memory Efficient**: Handles large datasets effectively

### **Data Volume**
- **Original**: Transaction-based records
- **Expanded**: Monthly records for all valid combinations
- **Manager Fields**: Dynamic hierarchy analysis
- **Calculated Columns**: Automated derived fields

## Best Practices

### **Data Input**
1. Ensure all required columns are present
2. Use consistent date formats (YYYY-MM-DD)
3. Validate employee numbers across tables
4. Check for duplicate records
5. Include supervisor relationships in UKG data

### **Processing**
1. Run tests before processing production data
2. Validate output data quality
3. Check for over-allocation patterns
4. Review employee exit records
5. Verify manager hierarchy accuracy

### **Analysis**
1. Use YearMonth for monthly grouping
2. Sum Percent values for allocation totals
3. Filter by active employees (no 0% records)
4. Monitor over-allocation trends
5. Analyze organizational structure with mgr fields

## Troubleshooting

### **Common Issues**

1. **Missing Excel File**: Check file path in `process_resource_data.py`
2. **Column Errors**: Ensure all required columns are present
3. **Date Format Issues**: Use consistent YYYY-MM-DD format
4. **Memory Issues**: Process smaller datasets or increase system memory
5. **Field Configuration Errors**: Verify field names match exactly

### **Validation Errors**

1. **Over-Allocation**: Review employee assignments
2. **Missing Data**: Check source Excel files
3. **Date Range Issues**: Verify AsOfDate values
4. **Employee Exits**: Confirm 0% records are intentional
5. **Manager Hierarchy**: Check supervisor relationships in UKG data

### **Output Issues**

1. **Locked Files**: Output directory prevents this automatically
2. **Missing Files**: Check output directory for timestamped files
3. **Field Configuration**: Verify tblFieldConfig worksheet format
4. **Manager Fields**: Ensure UKG data has supervisor relationships

## Contributing

1. **Add Tests**: New features should include comprehensive tests
2. **Follow Patterns**: Use existing test structure and naming
3. **Validate Data**: Run validation on test data
4. **Document Changes**: Update README for new features
5. **Test Hierarchy**: Verify manager chain functionality

## License

[Add your license information here]

## Support

For issues or questions:
1. Check the test suite for similar scenarios
2. Run validation on your data
3. Review the troubleshooting section
4. Verify field configuration format
5. Check manager hierarchy data integrity
6. Create an issue with detailed information
