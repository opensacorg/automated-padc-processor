# Sample Data Files

This directory contains example data files for testing and demonstration purposes.

## Files

### sample_monthly_attendance.xlsx
- **Description**: Example monthly attendance summary report from school system
- **Purpose**: Input file for the ADA audit process
- **Format**: Excel format with student attendance data organized by program, month, and grade level
- **Usage**: Use this as a template or test file when running the ADA Audit GUI

### sample_ada_reconciliation.xlsx
- **Description**: Example ADA reconciliation template
- **Purpose**: Output file template for processed audit results
- **Format**: Excel format with pre-formatted cells for ADA calculations
- **Usage**: Select this as the output file when running the audit process

## Using Sample Files

### Quick Test
```bash
python ADA_Audit_GUI.py
```

1. Click "Select Input File" and choose `sample_monthly_attendance.xlsx`
2. Click "Select Output File" and choose `sample_ada_reconciliation.xlsx`
3. Click "Load & Analyze" to auto-detect program boundaries
4. Click "Run Audit" to process the data

### Command Line Testing
```bash
python test_ada_audit.py
```

The unit tests automatically use these sample files to verify functionality.

## Notes

⚠️ **Data Privacy**: These files contain sample/anonymized data for testing purposes only. Do not commit real student data to the repository.

✅ **Version Control**: These sample files are tracked in Git to provide working examples for users.

## File Structure

```
sample_data/
├── README.md                          # This file
├── sample_monthly_attendance.xlsx     # Input example
└── sample_ada_reconciliation.xlsx     # Output template
```
