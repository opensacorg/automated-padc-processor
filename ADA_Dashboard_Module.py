"""
ADA Dashboard Module - Integration version for GUI
Adapted from ADA Dashboard_v2 (1).py to work with the GUI boundary system
"""

import os
import pandas as pd
import openpyxl
from tqdm import tqdm
import time
from prettytable import PrettyTable
import tkinter as tk
from tkinter import messagebox, filedialog, simpledialog


def find_row_with_value(df, target_value):
    """
    Finds row numbers where the target value is located in the DataFrame.
    """
    rows = []
    for index, value in enumerate(df.iloc[:, 1], start=1):
        if value == target_value:
            rows.append(index)
    return rows


def find_occurrences_of_number(df, number):
    """
    Finds occurrences of a specified number [Months] in the third column of the DataFrame.
    """
    rows = []
    for index, value in enumerate(df.iloc[:, 2], start=1):
        if pd.isna(value):
            continue
        try:
            if int(value) == number:
                rows.append(index)
        except ValueError:
            continue
    return rows


def find_start_stop_indices(rows):
    """
    Finds the lowest and highest row numbers from the given list of row numbers.
    """
    if not rows:
        return None, None
    return min(rows), max(rows)


def check_occurrences_and_create_fields(number_occurrences, target_indices, df):
    """
    Check occurrences [Months] and create fields based on their values.
    """
    created_fields = {}
    for number, occurrences in number_occurrences.items():
        for row_number in occurrences:
            for program, indices in target_indices.items():
                if indices["start"] is not None and indices["stop"] is not None:
                    if indices["start"] <= row_number <= indices["stop"]:
                        col_value = df.iloc[row_number-1, 4]  # Grade level column
                        Month_value = df.iloc[row_number-1, 2]  # Month column
                        APA_value = df.iloc[row_number-1, 39]  # APA column
                        ADA_Perc = df.iloc[row_number-1, 47]  # ADA percentage column
                        field_name = f"{program}_Month_{Month_value}_{col_value}: "
                        created_fields[field_name] = APA_value, ADA_Perc
    
    return created_fields


def parse_data_to_csv(data, school_year=None, location=None, school_name=None, output_dir=None):
    """
    Parse the extracted data to CSV format with proper structure.
    """
    csv_data = []

    # Define the grade prefix mapping
    grade_prefix_mapping = {
        "TK-3": "1 Grade",
        "4-6": "2 Grade", 
        "7-8": "3 Grade",
        "9-12": "4 Grade"
    }

    # Process each field in the data
    for field_name, values in data.items():
        apa_value = values[0]  # APA_value is first in the tuple
        ada_percentage = values[1]  # ADA_Perc is second in the tuple

        # Parse the field name to extract Month, Program, and Grade Level
        split_field = field_name.split("_")
        
        # Extract program (position 1)
        program = split_field[1]  # E.g., 'C', 'N', 'J', 'K'
        
        # Find month number - look for "Month" and get the next element
        month = None
        for i, part in enumerate(split_field):
            if part == "Month" and i + 1 < len(split_field):
                month = split_field[i + 1]
                break
        
        # Extract grade level - it's the last part before the colon, clean it up
        grade_level_raw = split_field[-1].rstrip(': ')
        grade_level = grade_level_raw
        
        # Determine TK indicator
        tk_indicator = "Y" if "TK" in field_name and "Prog_" + program + "_TK_" in field_name else "N"

        # Append each row to csv_data (only if month is valid)
        if month and month.isdigit():
            csv_data.append({
                "Year": school_year,
                "School": school_name,
                "Location": location,
                "Month": f"M{int(month):02d}",
                "Program": program,
                "TK": tk_indicator,
                "Grade Level": f"{grade_prefix_mapping.get(grade_level, 'Unknown Grade')} {grade_level}",
                "ADA %": f"{float(ada_percentage) * 100:.2f}%" if ada_percentage else "0.00%",
                "Total ADA": f"{float(apa_value):.2f}" if apa_value else "0.00"
            })

    # Create a DataFrame and export to CSV
    if csv_data:
        df = pd.DataFrame(csv_data)
        
        # Create output directory if specified
        if output_dir:
            os.makedirs(output_dir, exist_ok=True)
            
        # Create a timestamped CSV filename
        timestamp = time.strftime("%Y%m%d_%H%M%S")
        csv_filename = f"ada_dashboard_output_{timestamp}.csv"
        
        # Set full path if output directory is specified
        if output_dir:
            csv_filename = os.path.join(output_dir, csv_filename)
            standard_filename = os.path.join(output_dir, "ada_dashboard_output.csv")
        else:
            standard_filename = "ada_dashboard_output.csv"
        
        # Write to both timestamped and standard files
        df.to_csv(csv_filename, index=False)
        df.to_csv(standard_filename, index=False)
        
        return csv_filename, len(csv_data), df
    else:
        return None, 0, None


def run_ada_dashboard_with_boundaries(input_file_path, program_boundaries, program_mappings, 
                                    school_year=None, location=None, school_name=None,
                                    output_dir=None, progress_callback=None, log_callback=None):
    """
    Run the ADA Dashboard process using the provided boundaries from the GUI.
    
    Args:
        input_file_path: Path to the Excel attendance file
        program_boundaries: Dictionary of program boundaries from GUI
        program_mappings: Dictionary of program name mappings from GUI
        school_year: School year for the dashboard
        location: Location for the dashboard
        school_name: School name for the dashboard
        output_dir: Directory to save output files
        progress_callback: Function to call for progress updates
        log_callback: Function to call for log messages
        
    Returns:
        Dictionary with results including CSV file path and data summary
    """
    
    def log(message, msg_type='info'):
        if log_callback:
            log_callback(message, msg_type)
        else:
            print(message)
    
    def update_progress(value):
        if progress_callback:
            progress_callback(value)
    
    try:
        log("🚀 Starting ADA Dashboard process...")
        update_progress(10)
        
        # Validate input file
        if not os.path.exists(input_file_path):
            raise FileNotFoundError(f"Input file not found: {input_file_path}")
        
        log(f"📄 Reading data from: {os.path.basename(input_file_path)}")
        
        # Read the Excel file
        df = pd.read_excel(input_file_path, header=None)
        update_progress(20)
        
        # Convert GUI boundaries to the format expected by dashboard functions
        target_indices = {}
        for short_code, boundaries in program_boundaries.items():
            target_indices[short_code] = {
                "start": boundaries.get("start"),
                "stop": boundaries.get("stop")
            }
        
        log(f"📊 Using {len(target_indices)} program boundaries from GUI")
        update_progress(30)
        
        # Find month occurrences for all 12 months
        log("📅 Finding month occurrences in data...")
        number_occurrences = {}
        for i in range(1, 13):
            occurrences = find_occurrences_of_number(df, i)
            number_occurrences[i] = occurrences
            log(f"  Month {i}: Found in {len(occurrences)} rows")
        
        update_progress(50)
        
        # Extract attendance data using boundaries
        log("📈 Extracting attendance data based on program boundaries...")
        created_fields = check_occurrences_and_create_fields(number_occurrences, target_indices, df)
        
        log(f"✅ Extracted {len(created_fields)} attendance data fields")
        update_progress(70)
        
        # Get configuration from user if not provided
        if not school_year:
            school_year = "2024-2025"  # Default
        if not location:
            location = "TK-12"  # Default
        if not school_name:
            school_name = "CCCS"  # Default
            
        log(f"📋 Dashboard Configuration - Year: {school_year}, Location: {location}, School: {school_name}")
        
        # Generate CSV output
        log("💾 Generating CSV dashboard output...")
        csv_file, record_count, df_output = parse_data_to_csv(
            created_fields, school_year, location, school_name, output_dir
        )
        
        update_progress(90)
        
        if csv_file:
            log(f"✅ Dashboard CSV created: {os.path.basename(csv_file)}")
            log(f"📊 Total records: {record_count}")
            
            # Create summary table for display
            if df_output is not None and not df_output.empty:
                summary_by_program = df_output.groupby(['Program', 'Month']).agg({
                    'Total ADA': lambda x: sum(float(val) for val in x),
                    'Grade Level': 'count'
                }).reset_index()
                
                log("📈 Summary by Program and Month:")
                for _, row in summary_by_program.head(10).iterrows():
                    log(f"  Program {row['Program']} {row['Month']}: {row['Total ADA']:.2f} ADA ({row['Grade Level']} grade levels)")
        else:
            log("❌ No data was extracted - check boundaries and input file", 'error')
            
        update_progress(100)
        
        return {
            'success': True,
            'csv_file': csv_file,
            'record_count': record_count,
            'data_fields': len(created_fields),
            'output_data': df_output,
            'message': f'Dashboard completed successfully with {record_count} records'
        }
        
    except Exception as e:
        log(f"❌ Dashboard process failed: {str(e)}", 'error')
        return {
            'success': False,
            'error': str(e),
            'message': f'Dashboard failed: {str(e)}'
        }


def get_dashboard_configuration_from_user():
    """
    Get dashboard configuration from user via dialog boxes.
    Returns tuple of (school_year, location, school_name)
    """
    
    # Create a simple dialog for configuration
    root = tk.Tk()
    root.withdraw()  # Hide main window
    
    try:
        school_year = simpledialog.askstring(
            "Dashboard Configuration",
            "Enter School Year (e.g., 2024-2025):",
            initialvalue="2024-2025"
        )
        
        location = simpledialog.askstring(
            "Dashboard Configuration", 
            "Enter Location (e.g., TK-8, Elementary, High):",
            initialvalue="TK-12"
        )
        
        school_name = simpledialog.askstring(
            "Dashboard Configuration",
            "Enter School Name (e.g., CCCS):",
            initialvalue="CCCS"
        )
        
        # Use defaults if user cancels
        if not school_year:
            school_year = "2024-2025"
        if not location:
            location = "TK-12"
        if not school_name:
            school_name = "CCCS"
            
        return school_year, location, school_name
        
    except Exception:
        # Return defaults if any error occurs
        return "2024-2025", "TK-12", "CCCS"
    finally:
        root.destroy()


def validate_boundaries_for_dashboard(program_boundaries):
    """
    Validate that the program boundaries are suitable for dashboard processing.
    Returns (is_valid, message, is_warning)
    where is_warning indicates if this is a warning that can be ignored
    """
    
    valid_programs = 0
    total_programs = len(program_boundaries)
    missing_programs = []
    
    for program, boundaries in program_boundaries.items():
        start = boundaries.get("start")
        stop = boundaries.get("stop")
        
        if start is not None and stop is not None:
            valid_programs += 1
        else:
            missing_programs.append(program)
    
    if valid_programs == 0:
        return False, "No valid program boundaries found. Please load and analyze data first.", False
    
    if valid_programs < total_programs / 2:
        warning_message = f"Warning: Only {valid_programs} out of {total_programs} programs have valid boundaries.\n\n"
        warning_message += f"Missing boundaries for: {', '.join(missing_programs[:5])}"
        if len(missing_programs) > 5:
            warning_message += f" (and {len(missing_programs) - 5} more)"
        warning_message += "\n\nThe dashboard will only process programs with valid boundaries."
        warning_message += "\n\nDo you want to continue anyway?"
        
        return False, warning_message, True  # This is a warning that can be ignored
    
    return True, f"Boundaries validated: {valid_programs}/{total_programs} programs ready for dashboard processing", False


if __name__ == "__main__":
    # This allows the module to be run standalone for testing
    print("ADA Dashboard Module - Use this module by importing it into the GUI")
    print("For standalone testing, run the original ADA Dashboard_v2 (1).py file")