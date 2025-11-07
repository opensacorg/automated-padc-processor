"""
Unit tests for ADA Audit Tool - PADC Processor

Tests the core functionality of the ADA audit processing system using real data files.
"""

import unittest
import pandas as pd
import os
from pathlib import Path
from ADA_Audit_25_26_IMPROVED import (
    find_rows_containing_program_name,
    find_rows_containing_month_number,
    find_program_boundary_rows,
    extract_student_attendance_data
)


class TestADAuditFunctions(unittest.TestCase):
    """Test suite for ADA audit processing functions"""
    
    @classmethod
    def setUpClass(cls):
        """Load test data files once for all tests"""
        cls.input_file = r"C:\Users\Shawn\Downloads\PrintMonthlyAttendanceSummaryTotals_20251105_092651_658012e.xlsx"
        cls.reconciliation_file = r"C:\Users\Shawn\Downloads\2025-2026_I4C_ADA_Reconciliation.xlsx"
        
        # Check if files exist
        if not os.path.exists(cls.input_file):
            raise FileNotFoundError(f"Input file not found: {cls.input_file}")
        if not os.path.exists(cls.reconciliation_file):
            raise FileNotFoundError(f"Reconciliation file not found: {cls.reconciliation_file}")
        
        # Load the data
        print(f"\nLoading test data from: {cls.input_file}")
        cls.student_data = pd.read_excel(cls.input_file, header=None)
        print(f"Loaded {len(cls.student_data)} rows of data")
        
    def test_file_loading(self):
        """Test that data files load correctly"""
        self.assertIsNotNone(self.student_data)
        self.assertGreater(len(self.student_data), 0)
        print(f"✓ File loaded successfully with {len(self.student_data)} rows")
        
    def test_find_program_c_rows(self):
        """Test finding Program C Charter Resident rows"""
        program_name = "Program C Charter Resident"
        rows = find_rows_containing_program_name(self.student_data, program_name)
        
        self.assertIsInstance(rows, list)
        print(f"✓ Found {len(rows)} rows containing '{program_name}'")
        
        if len(rows) > 0:
            print(f"  First occurrence at row: {rows[0]}")
            print(f"  Last occurrence at row: {rows[-1]}")
            
    def test_find_program_n_rows(self):
        """Test finding Program N Non-Resident Charter rows"""
        program_name = "Program N Non-Resident Charter"
        rows = find_rows_containing_program_name(self.student_data, program_name)
        
        self.assertIsInstance(rows, list)
        print(f"✓ Found {len(rows)} rows containing '{program_name}'")
        
    def test_find_month_numbers(self):
        """Test finding rows for each month (1-12)"""
        for month in range(1, 13):
            rows = find_rows_containing_month_number(self.student_data, month)
            self.assertIsInstance(rows, list)
            if len(rows) > 0:
                print(f"✓ Month {month:2d}: Found {len(rows)} occurrences")
                
    def test_find_program_boundaries(self):
        """Test boundary detection for major programs"""
        test_programs = [
            "Program C Charter Resident",
            "Program N Non-Resident Charter",
            "Program J Independent Study Charter Resident",
            "Program K Independent Study Charter Non-Resident"
        ]
        
        print("\n" + "="*60)
        print("Program Boundary Detection:")
        print("="*60)
        
        for program_name in test_programs:
            rows = find_rows_containing_program_name(self.student_data, program_name)
            start, stop = find_program_boundary_rows(rows)
            
            if start is not None and stop is not None:
                self.assertLessEqual(start, stop)
                print(f"✓ {program_name}")
                print(f"  Start: Row {start}, Stop: Row {stop}, Span: {stop - start + 1} rows")
            else:
                print(f"⚠ {program_name}: Not found in data")
                
    def test_extract_attendance_data(self):
        """Test extracting attendance data for a program"""
        # Find Program C boundaries
        program_name = "Program C Charter Resident"
        prog_c_rows = find_rows_containing_program_name(self.student_data, program_name)
        start, stop = find_program_boundary_rows(prog_c_rows)
        
        if start is None or stop is None:
            self.skipTest(f"Program C not found in test data")
            
        # Build month attendance mapping for just month 1
        monthly_attendance = {
            1: find_rows_containing_month_number(self.student_data, 1)
        }
        
        # Build program boundary info
        program_boundaries = {
            "Prog_C": {"start": start, "stop": stop}
        }
        
        # Extract data
        attendance_data = extract_student_attendance_data(
            monthly_attendance,
            program_boundaries,
            self.student_data
        )
        
        self.assertIsInstance(attendance_data, dict)
        print(f"\n✓ Extracted {len(attendance_data)} attendance data points")
        
        # Show sample of extracted data
        if len(attendance_data) > 0:
            print("\nSample extracted data:")
            for i, (key, value) in enumerate(list(attendance_data.items())[:5]):
                print(f"  {key}{value}")
                
    def test_boundary_validation(self):
        """Test that boundaries don't overlap"""
        program_configs = {
            "Prog_C": "Program C Charter Resident",
            "Prog_N": "Program N Non-Resident Charter",
            "Prog_J": "Program J Independent Study Charter Resident",
            "Prog_K": "Program K Independent Study Charter Non-Resident"
        }
        
        boundaries = {}
        for code, name in program_configs.items():
            rows = find_rows_containing_program_name(self.student_data, name)
            start, stop = find_program_boundary_rows(rows)
            if start is not None:
                boundaries[code] = {"start": start, "stop": stop}
                
        # Check for overlaps
        boundary_list = sorted(boundaries.items(), key=lambda x: x[1]["start"])
        
        print("\n" + "="*60)
        print("Boundary Overlap Check:")
        print("="*60)
        
        for i in range(len(boundary_list) - 1):
            curr_code, curr_bounds = boundary_list[i]
            next_code, next_bounds = boundary_list[i + 1]
            
            gap = next_bounds["start"] - curr_bounds["stop"]
            
            print(f"{curr_code}: {curr_bounds['start']}-{curr_bounds['stop']}")
            print(f"  Gap to {next_code}: {gap} rows")
            
            # No overlap means next starts after current ends
            self.assertGreater(next_bounds["start"], curr_bounds["stop"],
                             f"Overlap detected between {curr_code} and {next_code}")
                             
        print("✓ No overlaps detected")
        
    def test_data_quality(self):
        """Test basic data quality checks"""
        print("\n" + "="*60)
        print("Data Quality Checks:")
        print("="*60)
        
        # Check for empty DataFrame
        self.assertFalse(self.student_data.empty, "DataFrame should not be empty")
        print(f"✓ DataFrame has {len(self.student_data)} rows")
        
        # Check number of columns (should have column AJ which is index 35)
        num_columns = len(self.student_data.columns)
        self.assertGreaterEqual(num_columns, 36, "Should have at least 36 columns for column AJ")
        print(f"✓ DataFrame has {num_columns} columns")
        
        # Check if column B (index 1) has program names
        column_b_values = self.student_data.iloc[:, 1].dropna()
        self.assertGreater(len(column_b_values), 0, "Column B should have data")
        print(f"✓ Column B has {len(column_b_values)} non-empty cells")
        
        # Check if column C (index 2) has month numbers
        column_c_values = self.student_data.iloc[:, 2].dropna()
        self.assertGreater(len(column_c_values), 0, "Column C should have data")
        print(f"✓ Column C has {len(column_c_values)} non-empty cells")


class TestBoundaryConfiguration(unittest.TestCase):
    """Test saved boundary configuration files"""
    
    def test_boundary_files_exist(self):
        """Test that example boundary configuration exists"""
        boundary_dir = Path(r"C:\Users\Shawn\Desktop\GCC_AI\automated-padc-processor\boundary_settings")
        
        self.assertTrue(boundary_dir.exists(), "Boundary settings directory should exist")
        
        # Check for .gitkeep
        gitkeep = boundary_dir / ".gitkeep"
        self.assertTrue(gitkeep.exists(), ".gitkeep file should exist")
        
        # Check for example configuration
        example_config = boundary_dir / "example_configuration.json"
        self.assertTrue(example_config.exists(), "Example configuration should exist")
        
        print(f"\n✓ Boundary settings directory exists: {boundary_dir}")
        print(f"✓ Example configuration exists: {example_config}")


def run_tests():
    """Run all tests with verbose output"""
    # Create test suite
    loader = unittest.TestLoader()
    suite = unittest.TestSuite()
    
    # Add all test classes
    suite.addTests(loader.loadTestsFromTestCase(TestADAuditFunctions))
    suite.addTests(loader.loadTestsFromTestCase(TestBoundaryConfiguration))
    
    # Run with verbose output
    runner = unittest.TextTestRunner(verbosity=2)
    result = runner.run(suite)
    
    # Print summary
    print("\n" + "="*60)
    print("TEST SUMMARY")
    print("="*60)
    print(f"Tests run: {result.testsRun}")
    print(f"Successes: {result.testsRun - len(result.failures) - len(result.errors)}")
    print(f"Failures: {len(result.failures)}")
    print(f"Errors: {len(result.errors)}")
    print("="*60)
    
    return result.wasSuccessful()


if __name__ == "__main__":
    success = run_tests()
    exit(0 if success else 1)
