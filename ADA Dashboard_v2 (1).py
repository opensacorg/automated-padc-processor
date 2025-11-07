import os
import pandas as pd
import openpyxl
from tqdm import tqdm
import time
from prettytable import PrettyTable


# Define the functions

def find_row_with_value(df, target_value):
    """
    Finds row numbers where the target value is located in the DataFrame.

    Args:
    - df (DataFrame): DataFrame to search in.
    - target_value (str): Target value to search for.
    

    Returns:
    - list of int: List of row numbers where the target value is located.
    """
    rows = []
    for index, value in enumerate(df.iloc[:, 1], start=1):  # Using positional indexing for the second column
        if value == target_value:
            rows.append(index)
    return rows

def find_occurrences_of_number(df, number):
    """
    Finds occurrences of a specified number [Months] in the second column of the DataFrame.

    Args:
    - df (DataFrame): DataFrame to search in.
    - number (int): Number to search for.

    Returns:
    - list of int: List of row numbers where the specified number is found in the second column.
    """
    rows = []
    for index, value in enumerate(df.iloc[:, 2], start=1):  # Using positional indexing for the second column
        if pd.isna(value):
            continue  # Skip NaN values
        try:
            if int(value) == number:
                rows.append(index)
        except ValueError:
            continue  # Skip non-numeric values
    return rows


def find_start_stop_indices(rows):
    """
    Finds the lowest and highest row numbers from the given list of row numbers. This allows you to know where the program codes begin and end

    Args:
    - rows (list of int): List of row numbers.

    Returns:
    - int: Lowest row number.
    - int: Highest row number.
    """
    if not rows:
        return None, None
    return min(rows), max(rows)

def check_occurrences_and_create_fields(number_occurrences, target_indices, df):
    """
    Check occurrences [Months] and create fields based on their values.

    Args:
    - number_occurrences (dict): Dictionary containing row numbers where each number is found.
    - target_indices (dict): Dictionary containing start and stop indices for each target program.
    - df (DataFrame): DataFrame containing the data.

    Returns:
    - dict: Dictionary containing created fields.
    """
    created_fields = {}
    for number, occurrences in number_occurrences.items():
        for row_number in occurrences:
            for program, indices in target_indices.items():
                if indices["start"] is not None and indices["stop"] is not None:
                    if indices["start"] <= row_number <= indices["stop"]:
                        col_value = df.iloc[row_number-1, 4]  # Assuming you want to check the column +1 over
                        Month_value= df.iloc[row_number-1, 2]  # Assuming you want to check the column +1 over
                        APA_value = df.iloc[row_number-1, 39]
                        ADA_Perc = df.iloc[row_number-1, 47]
                        field_name = f"{program}_Month_{Month_value}_{col_value}: "
                        created_fields[field_name] = APA_value,ADA_Perc
  
    return created_fields


    
def batch_load_values(created_fields, excel_output_path, excel_sheet_name):
    cell_value_list_VA = [
        #Program C Placements
        ('E10', created_fields.get("Prog_C_Month_1_K-3: ", 0)),
        ('E11', created_fields.get("Prog_C_Month_1_4-6: ", 0)),
        ('E12', created_fields.get("Prog_C_Month_1_7-8: ", 0)),
        ('E13', created_fields.get("Prog_C_Month_1_9-12: ", 0)),
        ('E9', created_fields.get("Prog_C_TK_Month_1_K-3: ", 0)),
        ('F10', created_fields.get("Prog_C_Month_2_K-3: ", 0)),
        ('F11', created_fields.get("Prog_C_Month_2_4-6: ", 0)),
        ('F12', created_fields.get("Prog_C_Month_2_7-8: ", 0)),
        ('F13', created_fields.get("Prog_C_Month_2_9-12: ", 0)),
        ('F9', created_fields.get("Prog_C_TK_Month_2_K-3: ", 0)),
        ('G10', created_fields.get("Prog_C_Month_3_K-3: ", 0)),
        ('G11', created_fields.get("Prog_C_Month_3_4-6: ", 0)),
        ('G12', created_fields.get("Prog_C_Month_3_7-8: ", 0)),
        ('G13', created_fields.get("Prog_C_Month_3_9-12: ", 0)),
        ('G9', created_fields.get("Prog_C_TK_Month_3_K-3: ", 0)),
        ('H10', created_fields.get("Prog_C_Month_4_K-3: ", 0)),
        ('H11', created_fields.get("Prog_C_Month_4_4-6: ", 0)),
        ('H12', created_fields.get("Prog_C_Month_4_7-8: ", 0)),
        ('H13', created_fields.get("Prog_C_Month_4_9-12: ", 0)),
        ('H9', created_fields.get("Prog_C_TK_Month_4_K-3: ", 0)),
        ('I10', created_fields.get("Prog_C_Month_5_K-3: ", 0)),
        ('I11', created_fields.get("Prog_C_Month_5_4-6: ", 0)),
        ('I12', created_fields.get("Prog_C_Month_5_7-8: ", 0)),
        ('I13', created_fields.get("Prog_C_Month_5_9-12: ", 0)),
        ('I9', created_fields.get("Prog_C_TK_Month_5_K-3: ", 0)),
        ('J10', created_fields.get("Prog_C_Month_6_K-3: ", 0)),
        ('J11', created_fields.get("Prog_C_Month_6_4-6: ", 0)),
        ('J12', created_fields.get("Prog_C_Month_6_7-8: ", 0)),
        ('J13', created_fields.get("Prog_C_Month_6_9-12: ", 0)),
        ('J9', created_fields.get("Prog_C_TK_Month_6_K-3: ", 0)),
        ('K10', created_fields.get("Prog_C_Month_7_K-3: ", 0)),
        ('K11', created_fields.get("Prog_C_Month_7_4-6: ", 0)),
        ('K12', created_fields.get("Prog_C_Month_7_7-8: ", 0)),
        ('K13', created_fields.get("Prog_C_Month_7_9-12: ", 0)),
        ('K9', created_fields.get("Prog_C_TK_Month_7_K-3: ", 0)),
        ('L10', created_fields.get("Prog_C_Month_8_K-3: ", 0)),
        ('L11', created_fields.get("Prog_C_Month_8_4-6: ", 0)),
        ('L12', created_fields.get("Prog_C_Month_8_7-8: ", 0)),
        ('L13', created_fields.get("Prog_C_Month_8_9-12: ", 0)),
        ('L9', created_fields.get("Prog_C_TK_Month_8_K-3: ", 0)),
        ('M10', created_fields.get("Prog_C_Month_9_K-3: ", 0)),
        ('M11', created_fields.get("Prog_C_Month_9_4-6: ", 0)),
        ('M12', created_fields.get("Prog_C_Month_9_7-8: ", 0)),
        ('M13', created_fields.get("Prog_C_Month_9_9-12: ", 0)),
        ('M9', created_fields.get("Prog_C_TK_Month_9_K-3: ", 0)),
        ('O10', created_fields.get("Prog_C_Month_10_K-3: ", 0)),
        ('O11', created_fields.get("Prog_C_Month_10_4-6: ", 0)),
        ('O12', created_fields.get("Prog_C_Month_10_7-8: ", 0)),
        ('O13', created_fields.get("Prog_C_Month_10_9-12: ", 0)),
        ('O9', created_fields.get("Prog_C_TK_Month_10_K-3: ", 0)),
        ('P10', created_fields.get("Prog_C_Month_11_K-3: ", 0)),
        ('P11', created_fields.get("Prog_C_Month_11_4-6: ", 0)),
        ('P12', created_fields.get("Prog_C_Month_11_7-8: ", 0)),
        ('P13', created_fields.get("Prog_C_Month_11_9-12: ", 0)),
        ('P9', created_fields.get("Prog_C_TK_Month_11_K-3: ", 0)),
        #Program N Placements
        ('E52', created_fields.get("Prog_N_Month_1_K-3: ", 0)),
        ('E53', created_fields.get("Prog_N_Month_1_4-6: ", 0)),
        ('E54', created_fields.get("Prog_N_Month_1_7-8: ", 0)),
        ('E55', created_fields.get("Prog_N_Month_1_9-12: ", 0)),
        ('E51', created_fields.get("Prog_N_TK_Month_1_K-3: ", 0)),
        ('F52', created_fields.get("Prog_N_Month_2_K-3: ", 0)),
        ('F53', created_fields.get("Prog_N_Month_2_4-6: ", 0)),
        ('F54', created_fields.get("Prog_N_Month_2_7-8: ", 0)),
        ('F55', created_fields.get("Prog_N_Month_2_9-12: ", 0)),
        ('F51', created_fields.get("Prog_N_TK_Month_2_K-3: ", 0)),
        ('G52', created_fields.get("Prog_N_Month_3_K-3: ", 0)),
        ('G53', created_fields.get("Prog_N_Month_3_4-6: ", 0)),
        ('G54', created_fields.get("Prog_N_Month_3_7-8: ", 0)),
        ('G55', created_fields.get("Prog_N_Month_3_9-12: ", 0)),
        ('G51', created_fields.get("Prog_N_TK_Month_3_K-3: ", 0)),
        ('H52', created_fields.get("Prog_N_Month_4_K-3: ", 0)),
        ('H53', created_fields.get("Prog_N_Month_4_4-6: ", 0)),
        ('H54', created_fields.get("Prog_N_Month_4_7-8: ", 0)),
        ('H55', created_fields.get("Prog_N_Month_4_9-12: ", 0)),
        ('H51', created_fields.get("Prog_N_TK_Month_4_K-3: ", 0)),
        ('I52', created_fields.get("Prog_N_Month_5_K-3: ", 0)),
        ('I53', created_fields.get("Prog_N_Month_5_4-6: ", 0)),
        ('I54', created_fields.get("Prog_N_Month_5_7-8: ", 0)),
        ('I55', created_fields.get("Prog_N_Month_5_9-12: ", 0)),
        ('I51', created_fields.get("Prog_N_TK_Month_5_K-3: ", 0)),
        ('J52', created_fields.get("Prog_N_Month_6_K-3: ", 0)),
        ('J53', created_fields.get("Prog_N_Month_6_4-6: ", 0)),
        ('J54', created_fields.get("Prog_N_Month_6_7-8: ", 0)),
        ('J55', created_fields.get("Prog_N_Month_6_9-12: ", 0)),
        ('J51', created_fields.get("Prog_N_TK_Month_6_K-3: ", 0)),
        ('K52', created_fields.get("Prog_N_Month_7_K-3: ", 0)),
        ('K53', created_fields.get("Prog_N_Month_7_4-6: ", 0)),
        ('K54', created_fields.get("Prog_N_Month_7_7-8: ", 0)),
        ('K55', created_fields.get("Prog_N_Month_7_9-12: ", 0)),
        ('K51', created_fields.get("Prog_N_TK_Month_7_K-3: ", 0)),
        ('L52', created_fields.get("Prog_N_Month_8_K-3: ", 0)),
        ('L53', created_fields.get("Prog_N_Month_8_4-6: ", 0)),
        ('L54', created_fields.get("Prog_N_Month_8_7-8: ", 0)),
        ('L55', created_fields.get("Prog_N_Month_8_9-12: ", 0)),
        ('L51', created_fields.get("Prog_N_TK_Month_8_K-3: ", 0)),
        ('M52', created_fields.get("Prog_N_Month_9_K-3: ", 0)),
        ('M53', created_fields.get("Prog_N_Month_9_4-6: ", 0)),
        ('M54', created_fields.get("Prog_N_Month_9_7-8: ", 0)),
        ('M55', created_fields.get("Prog_N_Month_9_9-12: ", 0)),
        ('M51', created_fields.get("Prog_N_TK_Month_9_K-3: ", 0)),
        ('O52', created_fields.get("Prog_N_Month_10_K-3: ", 0)),
        ('O53', created_fields.get("Prog_N_Month_10_4-6: ", 0)),
        ('O54', created_fields.get("Prog_N_Month_10_7-8: ", 0)),
        ('O55', created_fields.get("Prog_N_Month_10_9-12: ", 0)),
        ('O51', created_fields.get("Prog_N_TK_Month_10_K-3: ", 0)),
        ('P52', created_fields.get("Prog_N_Month_11_K-3: ", 0)),
        ('P53', created_fields.get("Prog_N_Month_11_4-6: ", 0)),
        ('P54', created_fields.get("Prog_N_Month_11_7-8: ", 0)),
        ('P55', created_fields.get("Prog_N_Month_11_9-12: ", 0)),
        ('P51', created_fields.get("Prog_N_TK_Month_11_K-3: ", 0)),
        #Program J Placements
        ('E34', created_fields.get("Prog_J_Month_1_K-3: ", 0)),
        ('E35', created_fields.get("Prog_J_Month_1_4-6: ", 0)),
        ('E36', created_fields.get("Prog_J_Month_1_7-8: ", 0)),
        ('E37', created_fields.get("Prog_J_Month_1_9-12: ", 0)),
        ('E33', created_fields.get("Prog_J_TK_Month_1_K-3: ", 0)),
        ('F34', created_fields.get("Prog_J_Month_2_K-3: ", 0)),
        ('F35', created_fields.get("Prog_J_Month_2_4-6: ", 0)),
        ('F36', created_fields.get("Prog_J_Month_2_7-8: ", 0)),
        ('F37', created_fields.get("Prog_J_Month_2_9-12: ", 0)),
        ('F33', created_fields.get("Prog_J_TK_Month_2_K-3: ", 0)),
        ('G34', created_fields.get("Prog_J_Month_3_K-3: ", 0)),
        ('G35', created_fields.get("Prog_J_Month_3_4-6: ", 0)),
        ('G36', created_fields.get("Prog_J_Month_3_7-8: ", 0)),
        ('G37', created_fields.get("Prog_J_Month_3_9-12: ", 0)),
        ('G33', created_fields.get("Prog_J_TK_Month_3_K-3: ", 0)),
        ('H34', created_fields.get("Prog_J_Month_4_K-3: ", 0)),
        ('H35', created_fields.get("Prog_J_Month_4_4-6: ", 0)),
        ('H36', created_fields.get("Prog_J_Month_4_7-8: ", 0)),
        ('H37', created_fields.get("Prog_J_Month_4_9-12: ", 0)),
        ('H33', created_fields.get("Prog_J_TK_Month_4_K-3: ", 0)),
        ('I34', created_fields.get("Prog_J_Month_5_K-3: ", 0)),
        ('I35', created_fields.get("Prog_J_Month_5_4-6: ", 0)),
        ('I36', created_fields.get("Prog_J_Month_5_7-8: ", 0)),
        ('I37', created_fields.get("Prog_J_Month_5_9-12: ", 0)),
        ('I33', created_fields.get("Prog_J_TK_Month_5_K-3: ", 0)),
        ('J34', created_fields.get("Prog_J_Month_6_K-3: ", 0)),
        ('J35', created_fields.get("Prog_J_Month_6_4-6: ", 0)),
        ('J36', created_fields.get("Prog_J_Month_6_7-8: ", 0)),
        ('J37', created_fields.get("Prog_J_Month_6_9-12: ", 0)),
        ('J33', created_fields.get("Prog_J_TK_Month_6_K-3: ", 0)),
        ('K34', created_fields.get("Prog_J_Month_7_K-3: ", 0)),
        ('K35', created_fields.get("Prog_J_Month_7_4-6: ", 0)),
        ('K36', created_fields.get("Prog_J_Month_7_7-8: ", 0)),
        ('K37', created_fields.get("Prog_J_Month_7_9-12: ", 0)),
        ('K33', created_fields.get("Prog_J_TK_Month_7_K-3: ", 0)),
        ('L34', created_fields.get("Prog_J_Month_8_K-3: ", 0)),
        ('L35', created_fields.get("Prog_J_Month_8_4-6: ", 0)),
        ('L36', created_fields.get("Prog_J_Month_8_7-8: ", 0)),
        ('L37', created_fields.get("Prog_J_Month_8_9-12: ", 0)),
        ('L33', created_fields.get("Prog_J_TK_Month_8_K-3: ", 0)),
        ('M34', created_fields.get("Prog_J_Month_9_K-3: ", 0)),
        ('M35', created_fields.get("Prog_J_Month_9_4-6: ", 0)),
        ('M36', created_fields.get("Prog_J_Month_9_7-8: ", 0)),
        ('M37', created_fields.get("Prog_J_Month_9_9-12: ", 0)),
        ('M33', created_fields.get("Prog_J_TK_Month_9_K-3: ", 0)),
        ('O34', created_fields.get("Prog_J_Month_10_K-3: ", 0)),
        ('O35', created_fields.get("Prog_J_Month_10_4-6: ", 0)),
        ('O36', created_fields.get("Prog_J_Month_10_7-8: ", 0)),
        ('O37', created_fields.get("Prog_J_Month_10_9-12: ", 0)),
        ('O33', created_fields.get("Prog_J_TK_Month_10_K-3: ", 0)),
        ('P34', created_fields.get("Prog_J_Month_11_K-3: ", 0)),
        ('P35', created_fields.get("Prog_J_Month_11_4-6: ", 0)),
        ('P36', created_fields.get("Prog_J_Month_11_7-8: ", 0)),
        ('P37', created_fields.get("Prog_J_Month_11_9-12: ", 0)),
        ('P33', created_fields.get("Prog_J_TK_Month_11_K-3: ", 0)),
        #Program k Placements
        ('E76', created_fields.get("Prog_K_Month_1_K-3: ", 0)),
        ('E77', created_fields.get("Prog_K_Month_1_4-6: ", 0)),
        ('E78', created_fields.get("Prog_K_Month_1_7-8: ", 0)),
        ('E79', created_fields.get("Prog_K_Month_1_9-12: ", 0)),
        ('E75', created_fields.get("Prog_K_TK_Month_1_K-3: ", 0)),
        ('F76', created_fields.get("Prog_K_Month_2_K-3: ", 0)),
        ('F77', created_fields.get("Prog_K_Month_2_4-6: ", 0)),
        ('F78', created_fields.get("Prog_K_Month_2_7-8: ", 0)),
        ('F79', created_fields.get("Prog_K_Month_2_9-12: ", 0)),
        ('F75', created_fields.get("Prog_K_TK_Month_2_K-3: ", 0)),
        ('G76', created_fields.get("Prog_K_Month_3_K-3: ", 0)),
        ('G77', created_fields.get("Prog_K_Month_3_4-6: ", 0)),
        ('G78', created_fields.get("Prog_K_Month_3_7-8: ", 0)),
        ('G79', created_fields.get("Prog_K_Month_3_9-12: ", 0)),
        ('G75', created_fields.get("Prog_K_TK_Month_3_K-3: ", 0)),
        ('H76', created_fields.get("Prog_K_Month_4_K-3: ", 0)),
        ('H77', created_fields.get("Prog_K_Month_4_4-6: ", 0)),
        ('H78', created_fields.get("Prog_K_Month_4_7-8: ", 0)),
        ('H79', created_fields.get("Prog_K_Month_4_9-12: ", 0)),
        ('H75', created_fields.get("Prog_K_TK_Month_4_K-3: ", 0)),
        ('I76', created_fields.get("Prog_K_Month_5_K-3: ", 0)),
        ('I77', created_fields.get("Prog_K_Month_5_4-6: ", 0)),
        ('I78', created_fields.get("Prog_K_Month_5_7-8: ", 0)),
        ('I79', created_fields.get("Prog_K_Month_5_9-12: ", 0)),
        ('I75', created_fields.get("Prog_K_TK_Month_5_K-3: ", 0)),
        ('J76', created_fields.get("Prog_K_Month_6_K-3: ", 0)),
        ('J77', created_fields.get("Prog_K_Month_6_4-6: ", 0)),
        ('J78', created_fields.get("Prog_K_Month_6_7-8: ", 0)),
        ('J79', created_fields.get("Prog_K_Month_6_9-12: ", 0)),
        ('J75', created_fields.get("Prog_K_TK_Month_6_K-3: ", 0)),
        ('K76', created_fields.get("Prog_K_Month_7_K-3: ", 0)),
        ('K77', created_fields.get("Prog_K_Month_7_4-6: ", 0)),
        ('K78', created_fields.get("Prog_K_Month_7_7-8: ", 0)),
        ('K79', created_fields.get("Prog_K_Month_7_9-12: ", 0)),
        ('K75', created_fields.get("Prog_K_TK_Month_7_K-3: ", 0)),
        ('L76', created_fields.get("Prog_K_Month_8_K-3: ", 0)),
        ('L77', created_fields.get("Prog_K_Month_8_4-6: ", 0)),
        ('L78', created_fields.get("Prog_K_Month_8_7-8: ", 0)),
        ('L79', created_fields.get("Prog_K_Month_8_9-12: ", 0)),
        ('L75', created_fields.get("Prog_K_TK_Month_8_K-3: ", 0)),
        ('M76', created_fields.get("Prog_K_Month_9_K-3: ", 0)),
        ('M77', created_fields.get("Prog_K_Month_9_4-6: ", 0)),
        ('M78', created_fields.get("Prog_K_Month_9_7-8: ", 0)),
        ('M79', created_fields.get("Prog_K_Month_9_9-12: ", 0)),
        ('M75', created_fields.get("Prog_K_TK_Month_9_K-3: ", 0)),
        ('O76', created_fields.get("Prog_K_Month_10_K-3: ", 0)),
        ('O77', created_fields.get("Prog_K_Month_10_4-6: ", 0)),
        ('O78', created_fields.get("Prog_K_Month_10_7-8: ", 0)),
        ('O79', created_fields.get("Prog_K_Month_10_9-12: ", 0)),
        ('O75', created_fields.get("Prog_K_TK_Month_10_K-3: ", 0)),
        ('P76', created_fields.get("Prog_K_Month_11_K-3: ", 0)),
        ('P77', created_fields.get("Prog_K_Month_11_4-6: ", 0)),
        ('P78', created_fields.get("Prog_K_Month_11_7-8: ", 0)),
        ('P79', created_fields.get("Prog_K_Month_11_9-12: ", 0)),
        ('P75', created_fields.get("Prog_K_TK_Month_11_K-3: ", 0)),
                
        
        
]
    # Define the list of cell references and values to be written
    cell_value_list = [
        #Program C Placements
        ('E10', created_fields.get("Prog_C_Month_1_TK-3: ", 0)),
        ('E11', created_fields.get("Prog_C_Month_1_4-6: ", 0)),
        ('E12', created_fields.get("Prog_C_Month_1_7-8: ", 0)),
        ('E13', created_fields.get("Prog_C_Month_1_9-12: ", 0)),
        ('E9', created_fields.get("Prog_C_TK_Month_1_TK-3: ", 0)),
        ('F10', created_fields.get("Prog_C_Month_2_TK-3: ", 0)),
        ('F11', created_fields.get("Prog_C_Month_2_4-6: ", 0)),
        ('F12', created_fields.get("Prog_C_Month_2_7-8: ", 0)),
        ('F13', created_fields.get("Prog_C_Month_2_9-12: ", 0)),
        ('F9', created_fields.get("Prog_C_TK_Month_2_TK-3: ", 0)),
        ('G10', created_fields.get("Prog_C_Month_3_TK-3: ", 0)),
        ('G11', created_fields.get("Prog_C_Month_3_4-6: ", 0)),
        ('G12', created_fields.get("Prog_C_Month_3_7-8: ", 0)),
        ('G13', created_fields.get("Prog_C_Month_3_9-12: ", 0)),
        ('G9', created_fields.get("Prog_C_TK_Month_3_TK-3: ", 0)),
        ('H10', created_fields.get("Prog_C_Month_4_TK-3: ", 0)),
        ('H11', created_fields.get("Prog_C_Month_4_4-6: ", 0)),
        ('H12', created_fields.get("Prog_C_Month_4_7-8: ", 0)),
        ('H13', created_fields.get("Prog_C_Month_4_9-12: ", 0)),
        ('H9', created_fields.get("Prog_C_TK_Month_4_TK-3: ", 0)),
        ('I10', created_fields.get("Prog_C_Month_5_TK-3: ", 0)),
        ('I11', created_fields.get("Prog_C_Month_5_4-6: ", 0)),
        ('I12', created_fields.get("Prog_C_Month_5_7-8: ", 0)),
        ('I13', created_fields.get("Prog_C_Month_5_9-12: ", 0)),
        ('I9', created_fields.get("Prog_C_TK_Month_5_TK-3: ", 0)),
        ('J10', created_fields.get("Prog_C_Month_6_TK-3: ", 0)),
        ('J11', created_fields.get("Prog_C_Month_6_4-6: ", 0)),
        ('J12', created_fields.get("Prog_C_Month_6_7-8: ", 0)),
        ('J13', created_fields.get("Prog_C_Month_6_9-12: ", 0)),
        ('J9', created_fields.get("Prog_C_TK_Month_6_TK-3: ", 0)),
        ('K10', created_fields.get("Prog_C_Month_7_TK-3: ", 0)),
        ('K11', created_fields.get("Prog_C_Month_7_4-6: ", 0)),
        ('K12', created_fields.get("Prog_C_Month_7_7-8: ", 0)),
        ('K13', created_fields.get("Prog_C_Month_7_9-12: ", 0)),
        ('K9', created_fields.get("Prog_C_TK_Month_7_TK-3: ", 0)),
        ('L10', created_fields.get("Prog_C_Month_8_TK-3: ", 0)),
        ('L11', created_fields.get("Prog_C_Month_8_4-6: ", 0)),
        ('L12', created_fields.get("Prog_C_Month_8_7-8: ", 0)),
        ('L13', created_fields.get("Prog_C_Month_8_9-12: ", 0)),
        ('L9', created_fields.get("Prog_C_TK_Month_8_TK-3: ", 0)),
        ('M10', created_fields.get("Prog_C_Month_9_TK-3: ", 0)),
        ('M11', created_fields.get("Prog_C_Month_9_4-6: ", 0)),
        ('M12', created_fields.get("Prog_C_Month_9_7-8: ", 0)),
        ('M13', created_fields.get("Prog_C_Month_9_9-12: ", 0)),
        ('M9', created_fields.get("Prog_C_TK_Month_9_TK-3: ", 0)),
        ('N10', created_fields.get("Prog_C_Month_10_TK-3: ", 0)),
        ('N11', created_fields.get("Prog_C_Month_10_4-6: ", 0)),
        ('N12', created_fields.get("Prog_C_Month_10_7-8: ", 0)),
        ('N13', created_fields.get("Prog_C_Month_10_9-12: ", 0)),
        ('N9', created_fields.get("Prog_C_TK_Month_10_TK-3: ", 0)),
        ('O10', created_fields.get("Prog_C_Month_11_TK-3: ", 0)),
        ('O11', created_fields.get("Prog_C_Month_11_4-6: ", 0)),
        ('O12', created_fields.get("Prog_C_Month_11_7-8: ", 0)),
        ('O13', created_fields.get("Prog_C_Month_11_9-12: ", 0)),
        ('O9', created_fields.get("Prog_C_TK_Month_11_TK-3: ", 0)),
        ('P10', created_fields.get("Prog_C_Month_12_TK-3: ", 0)),
        ('P11', created_fields.get("Prog_C_Month_12_4-6: ", 0)),
        ('P12', created_fields.get("Prog_C_Month_12_7-8: ", 0)),
        ('P13', created_fields.get("Prog_C_Month_12_9-12: ", 0)),
        ('P9', created_fields.get("Prog_C_TK_Month_12_TK-3: ", 0)),
        #Program N Placements
        ('E52', created_fields.get("Prog_N_Month_1_TK-3: ", 0)),
        ('E53', created_fields.get("Prog_N_Month_1_4-6: ", 0)),
        ('E54', created_fields.get("Prog_N_Month_1_7-8: ", 0)),
        ('E55', created_fields.get("Prog_N_Month_1_9-12: ", 0)),
        ('E51', created_fields.get("Prog_N_TK_Month_1_TK-3: ", 0)),
        ('F52', created_fields.get("Prog_N_Month_2_TK-3: ", 0)),
        ('F53', created_fields.get("Prog_N_Month_2_4-6: ", 0)),
        ('F54', created_fields.get("Prog_N_Month_2_7-8: ", 0)),
        ('F55', created_fields.get("Prog_N_Month_2_9-12: ", 0)),
        ('F51', created_fields.get("Prog_N_TK_Month_2_TK-3: ", 0)),
        ('G52', created_fields.get("Prog_N_Month_3_TK-3: ", 0)),
        ('G53', created_fields.get("Prog_N_Month_3_4-6: ", 0)),
        ('G54', created_fields.get("Prog_N_Month_3_7-8: ", 0)),
        ('G55', created_fields.get("Prog_N_Month_3_9-12: ", 0)),
        ('G51', created_fields.get("Prog_N_TK_Month_3_TK-3: ", 0)),
        ('H52', created_fields.get("Prog_N_Month_4_TK-3: ", 0)),
        ('H53', created_fields.get("Prog_N_Month_4_4-6: ", 0)),
        ('H54', created_fields.get("Prog_N_Month_4_7-8: ", 0)),
        ('H55', created_fields.get("Prog_N_Month_4_9-12: ", 0)),
        ('H51', created_fields.get("Prog_N_TK_Month_4_TK-3: ", 0)),
        ('I52', created_fields.get("Prog_N_Month_5_TK-3: ", 0)),
        ('I53', created_fields.get("Prog_N_Month_5_4-6: ", 0)),
        ('I54', created_fields.get("Prog_N_Month_5_7-8: ", 0)),
        ('I55', created_fields.get("Prog_N_Month_5_9-12: ", 0)),
        ('I51', created_fields.get("Prog_N_TK_Month_5_TK-3: ", 0)),
        ('J52', created_fields.get("Prog_N_Month_6_TK-3: ", 0)),
        ('J53', created_fields.get("Prog_N_Month_6_4-6: ", 0)),
        ('J54', created_fields.get("Prog_N_Month_6_7-8: ", 0)),
        ('J55', created_fields.get("Prog_N_Month_6_9-12: ", 0)),
        ('J51', created_fields.get("Prog_N_TK_Month_6_TK-3: ", 0)),
        ('K52', created_fields.get("Prog_N_Month_7_TK-3: ", 0)),
        ('K53', created_fields.get("Prog_N_Month_7_4-6: ", 0)),
        ('K54', created_fields.get("Prog_N_Month_7_7-8: ", 0)),
        ('K55', created_fields.get("Prog_N_Month_7_9-12: ", 0)),
        ('K51', created_fields.get("Prog_N_TK_Month_7_TK-3: ", 0)),
        ('L52', created_fields.get("Prog_N_Month_8_TK-3: ", 0)),
        ('L53', created_fields.get("Prog_N_Month_8_4-6: ", 0)),
        ('L54', created_fields.get("Prog_N_Month_8_7-8: ", 0)),
        ('L55', created_fields.get("Prog_N_Month_8_9-12: ", 0)),
        ('L51', created_fields.get("Prog_N_TK_Month_8_TK-3: ", 0)),
        ('M52', created_fields.get("Prog_N_Month_9_TK-3: ", 0)),
        ('M53', created_fields.get("Prog_N_Month_9_4-6: ", 0)),
        ('M54', created_fields.get("Prog_N_Month_9_7-8: ", 0)),
        ('M55', created_fields.get("Prog_N_Month_9_9-12: ", 0)),
        ('M51', created_fields.get("Prog_N_TK_Month_9_TK-3: ", 0)),
        ('N52', created_fields.get("Prog_N_Month_10_TK-3: ", 0)),
        ('N53', created_fields.get("Prog_N_Month_10_4-6: ", 0)),
        ('N54', created_fields.get("Prog_N_Month_10_7-8: ", 0)),
        ('N55', created_fields.get("Prog_N_Month_10_9-12: ", 0)),
        ('N51', created_fields.get("Prog_N_TK_Month_10_TK-3: ", 0)),
        ('O52', created_fields.get("Prog_N_Month_11_TK-3: ", 0)),
        ('O53', created_fields.get("Prog_N_Month_11_4-6: ", 0)),
        ('O54', created_fields.get("Prog_N_Month_11_7-8: ", 0)),
        ('O55', created_fields.get("Prog_N_Month_11_9-12: ", 0)),
        ('O51', created_fields.get("Prog_N_TK_Month_11_TK-3: ", 0)),
        ('P52', created_fields.get("Prog_N_Month_12_TK-3: ", 0)),
        ('P53', created_fields.get("Prog_N_Month_12_4-6: ", 0)),
        ('P54', created_fields.get("Prog_N_Month_12_7-8: ", 0)),
        ('P55', created_fields.get("Prog_N_Month_12_9-12: ", 0)),
        ('P51', created_fields.get("Prog_N_TK_Month_12_TK-3: ", 0)),
    
        #Program J Placements
        ('E34', created_fields.get("Prog_J_Month_1_TK-3: ", 0)),
        ('E35', created_fields.get("Prog_J_Month_1_4-6: ", 0)),
        ('E36', created_fields.get("Prog_J_Month_1_7-8: ", 0)),
        ('E37', created_fields.get("Prog_J_Month_1_9-12: ", 0)),
        ('E33', created_fields.get("Prog_J_TK_Month_1_TK-3: ", 0)),
        ('F34', created_fields.get("Prog_J_Month_2_TK-3: ", 0)),
        ('F35', created_fields.get("Prog_J_Month_2_4-6: ", 0)),
        ('F36', created_fields.get("Prog_J_Month_2_7-8: ", 0)),
        ('F37', created_fields.get("Prog_J_Month_2_9-12: ", 0)),
        ('F33', created_fields.get("Prog_J_TK_Month_2_TK-3: ", 0)),
        ('G34', created_fields.get("Prog_J_Month_3_TK-3: ", 0)),
        ('G35', created_fields.get("Prog_J_Month_3_4-6: ", 0)),
        ('G36', created_fields.get("Prog_J_Month_3_7-8: ", 0)),
        ('G37', created_fields.get("Prog_J_Month_3_9-12: ", 0)),
        ('G33', created_fields.get("Prog_J_TK_Month_3_TK-3: ", 0)),
        ('H34', created_fields.get("Prog_J_Month_4_TK-3: ", 0)),
        ('H35', created_fields.get("Prog_J_Month_4_4-6: ", 0)),
        ('H36', created_fields.get("Prog_J_Month_4_7-8: ", 0)),
        ('H37', created_fields.get("Prog_J_Month_4_9-12: ", 0)),
        ('H33', created_fields.get("Prog_J_TK_Month_4_TK-3: ", 0)),
        ('I34', created_fields.get("Prog_J_Month_5_TK-3: ", 0)),
        ('I35', created_fields.get("Prog_J_Month_5_4-6: ", 0)),
        ('I36', created_fields.get("Prog_J_Month_5_7-8: ", 0)),
        ('I37', created_fields.get("Prog_J_Month_5_9-12: ", 0)),
        ('I33', created_fields.get("Prog_J_TK_Month_5_TK-3: ", 0)),
        ('J34', created_fields.get("Prog_J_Month_6_TK-3: ", 0)),
        ('J35', created_fields.get("Prog_J_Month_6_4-6: ", 0)),
        ('J36', created_fields.get("Prog_J_Month_6_7-8: ", 0)),
        ('J37', created_fields.get("Prog_J_Month_6_9-12: ", 0)),
        ('J33', created_fields.get("Prog_J_TK_Month_6_TK-3: ", 0)),
        ('K34', created_fields.get("Prog_J_Month_7_TK-3: ", 0)),
        ('K35', created_fields.get("Prog_J_Month_7_4-6: ", 0)),
        ('K36', created_fields.get("Prog_J_Month_7_7-8: ", 0)),
        ('K37', created_fields.get("Prog_J_Month_7_9-12: ", 0)),
        ('K33', created_fields.get("Prog_J_TK_Month_7_TK-3: ", 0)),
        ('L34', created_fields.get("Prog_J_Month_8_TK-3: ", 0)),
        ('L35', created_fields.get("Prog_J_Month_8_4-6: ", 0)),
        ('L36', created_fields.get("Prog_J_Month_8_7-8: ", 0)),
        ('L37', created_fields.get("Prog_J_Month_8_9-12: ", 0)),
        ('L33', created_fields.get("Prog_J_TK_Month_8_TK-3: ", 0)),
        ('M34', created_fields.get("Prog_J_Month_9_TK-3: ", 0)),
        ('M35', created_fields.get("Prog_J_Month_9_4-6: ", 0)),
        ('M36', created_fields.get("Prog_J_Month_9_7-8: ", 0)),
        ('M37', created_fields.get("Prog_J_Month_9_9-12: ", 0)),
        ('M33', created_fields.get("Prog_J_TK_Month_9_TK-3: ", 0)),
        ('N34', created_fields.get("Prog_J_Month_10_TK-3: ", 0)),
        ('N35', created_fields.get("Prog_J_Month_10_4-6: ", 0)),
        ('N36', created_fields.get("Prog_J_Month_10_7-8: ", 0)),
        ('N37', created_fields.get("Prog_J_Month_10_9-12: ", 0)),
        ('N33', created_fields.get("Prog_J_TK_Month_10_TK-3: ", 0)),
        ('O34', created_fields.get("Prog_J_Month_11_TK-3: ", 0)),
        ('O35', created_fields.get("Prog_J_Month_11_4-6: ", 0)),
        ('O36', created_fields.get("Prog_J_Month_11_7-8: ", 0)),
        ('O37', created_fields.get("Prog_J_Month_11_9-12: ", 0)),
        ('O33', created_fields.get("Prog_J_TK_Month_11_TK-3: ", 0)),
        ('P34', created_fields.get("Prog_J_Month_12_TK-3: ", 0)),
        ('P35', created_fields.get("Prog_J_Month_12_4-6: ", 0)),
        ('P36', created_fields.get("Prog_J_Month_12_7-8: ", 0)),
        ('P37', created_fields.get("Prog_J_Month_12_9-12: ", 0)),
        ('P33', created_fields.get("Prog_J_TK_Month_12_TK-3: ", 0)),
        #Program k Placements
        ('E76', created_fields.get("Prog_K_Month_1_TK-3: ", 0)),
        ('E77', created_fields.get("Prog_K_Month_1_4-6: ", 0)),
        ('E78', created_fields.get("Prog_K_Month_1_7-8: ", 0)),
        ('E79', created_fields.get("Prog_K_Month_1_9-12: ", 0)),
        ('E75', created_fields.get("Prog_K_TK_Month_1_TK-3: ", 0)),
        ('F76', created_fields.get("Prog_K_Month_2_TK-3: ", 0)),
        ('F77', created_fields.get("Prog_K_Month_2_4-6: ", 0)),
        ('F78', created_fields.get("Prog_K_Month_2_7-8: ", 0)),
        ('F79', created_fields.get("Prog_K_Month_2_9-12: ", 0)),
        ('F75', created_fields.get("Prog_K_TK_Month_2_TK-3: ", 0)),
        ('G76', created_fields.get("Prog_K_Month_3_TK-3: ", 0)),
        ('G77', created_fields.get("Prog_K_Month_3_4-6: ", 0)),
        ('G78', created_fields.get("Prog_K_Month_3_7-8: ", 0)),
        ('G79', created_fields.get("Prog_K_Month_3_9-12: ", 0)),
        ('G75', created_fields.get("Prog_K_TK_Month_3_TK-3: ", 0)),
        ('H76', created_fields.get("Prog_K_Month_4_TK-3: ", 0)),
        ('H77', created_fields.get("Prog_K_Month_4_4-6: ", 0)),
        ('H78', created_fields.get("Prog_K_Month_4_7-8: ", 0)),
        ('H79', created_fields.get("Prog_K_Month_4_9-12: ", 0)),
        ('H75', created_fields.get("Prog_K_TK_Month_4_TK-3: ", 0)),
        ('I76', created_fields.get("Prog_K_Month_5_TK-3: ", 0)),
        ('I77', created_fields.get("Prog_K_Month_5_4-6: ", 0)),
        ('I78', created_fields.get("Prog_K_Month_5_7-8: ", 0)),
        ('I79', created_fields.get("Prog_K_Month_5_9-12: ", 0)),
        ('I75', created_fields.get("Prog_K_TK_Month_5_TK-3: ", 0)),
        ('J76', created_fields.get("Prog_K_Month_6_TK-3: ", 0)),
        ('J77', created_fields.get("Prog_K_Month_6_4-6: ", 0)),
        ('J78', created_fields.get("Prog_K_Month_6_7-8: ", 0)),
        ('J79', created_fields.get("Prog_K_Month_6_9-12: ", 0)),
        ('J75', created_fields.get("Prog_K_TK_Month_6_TK-3: ", 0)),
        ('K76', created_fields.get("Prog_K_Month_7_TK-3: ", 0)),
        ('K77', created_fields.get("Prog_K_Month_7_4-6: ", 0)),
        ('K78', created_fields.get("Prog_K_Month_7_7-8: ", 0)),
        ('K79', created_fields.get("Prog_K_Month_7_9-12: ", 0)),
        ('K75', created_fields.get("Prog_K_TK_Month_7_TK-3: ", 0)),
        ('L76', created_fields.get("Prog_K_Month_8_TK-3: ", 0)),
        ('L77', created_fields.get("Prog_K_Month_8_4-6: ", 0)),
        ('L78', created_fields.get("Prog_K_Month_8_7-8: ", 0)),
        ('L79', created_fields.get("Prog_K_Month_8_9-12: ", 0)),
        ('L75', created_fields.get("Prog_K_TK_Month_8_TK-3: ", 0)),
        ('M76', created_fields.get("Prog_K_Month_9_TK-3: ", 0)),
        ('M77', created_fields.get("Prog_K_Month_9_4-6: ", 0)),
        ('M78', created_fields.get("Prog_K_Month_9_7-8: ", 0)),
        ('M79', created_fields.get("Prog_K_Month_9_9-12: ", 0)),
        ('M75', created_fields.get("Prog_K_TK_Month_9_TK-3: ", 0)),
        ('N76', created_fields.get("Prog_K_Month_10_TK-3: ", 0)),
        ('N77', created_fields.get("Prog_K_Month_10_4-6: ", 0)),
        ('N78', created_fields.get("Prog_K_Month_10_7-8: ", 0)),
        ('N79', created_fields.get("Prog_K_Month_10_9-12: ", 0)),
        ('N75', created_fields.get("Prog_K_TK_Month_10_TK-3: ", 0)),
        ('O76', created_fields.get("Prog_K_Month_11_TK-3: ", 0)),
        ('O77', created_fields.get("Prog_K_Month_11_4-6: ", 0)),
        ('O78', created_fields.get("Prog_K_Month_11_7-8: ", 0)),
        ('O79', created_fields.get("Prog_K_Month_11_9-12: ", 0)),
        ('O75', created_fields.get("Prog_K_TK_Month_11_TK-3: ", 0)),
        ('P76', created_fields.get("Prog_K_Month_12_TK-3: ", 0)),
        ('P77', created_fields.get("Prog_K_Month_12_4-6: ", 0)),
        ('P78', created_fields.get("Prog_K_Month_12_7-8: ", 0)),
        ('P79', created_fields.get("Prog_K_Month_12_9-12: ", 0)),
        ('P75', created_fields.get("Prog_K_TK_Month_12_TK-3: ", 0)),
                
        
        
]




    # Batch load values into cells
    for cell, value in cell_value_list:
        sheet[cell] = value



def parse_data_to_csv(data, school_year=None, location=None, school_name=None):
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
        # Field format: "Prog_C_Month_1_TK-3: " or "Prog_C_TK_Month_1_TK-3: "
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
        grade_level_raw = split_field[-1].rstrip(': ')  # Remove ': ' from end
        grade_level = grade_level_raw  # E.g., 'TK-3', '4-6', '7-8', '9-12'
        
        # Determine TK indicator
        tk_indicator = "Y" if "TK" in field_name and "Prog_" + program + "_TK_" in field_name else "N"
        ada_percentage_str = ada_percentage

        # Append each row to csv_data (only if month is valid)
        if month and month.isdigit():
            csv_data.append({
                "Year": school_year,
                "School": school_name,
                "Location": location,
                "Month": f"M{int(month):02d}",  # Format as M01, M02, etc.
                "Program": program,
                "TK": tk_indicator,
                "Grade Level": f"{grade_prefix_mapping.get(grade_level, 'Unknown Grade')} {grade_level}",
                "ADA %": f"{float(ada_percentage) * 100:.2f}%" if ada_percentage else "0.00%",
                "Total ADA": f"{float(apa_value):.2f}" if apa_value else "0.00"
            })

    # Create a DataFrame and export to CSV
    if csv_data:
        df = pd.DataFrame(csv_data)
        
        # Create a timestamped CSV filename to avoid caching issues
        timestamp = time.strftime("%Y%m%d_%H%M%S")
        csv_filename = f"output_{timestamp}.csv"
        
        # Write to both the timestamped file and the standard output.csv
        df.to_csv(csv_filename, index=False)
        df.to_csv("output.csv", index=False)
        
        print(f"\n✅ CSV files created successfully:")
        print(f"   📄 {csv_filename}")
        print(f"   📄 output.csv")
        print(f"   📊 Total records: {len(csv_data)}")
        
        # Check file sizes
        import os
        if os.path.exists(csv_filename):
            size = os.path.getsize(csv_filename)
            print(f"   💾 File size: {size} bytes")
        
    else:
        print("\n❌ No data to export to CSV. Check the data processing logic.")
        return

    # Display the CSV structure for verification
    table = PrettyTable(["Year", "School", "Location", "Month", "Program", "TK", "Grade Level", "ADA %", "Total ADA"])
    for row in csv_data:
        table.add_row([row["Year"], row["School"], row["Location"], row["Month"], row["Program"], row["TK"], row["Grade Level"], row["ADA %"], row["Total ADA"]])
    
    print("\nGenerated CSV Data:")
    print(table)


def main():
    """
    Main function to execute the program.
    """
    # Get user input for Location, School Year, and School Name
    print("=== ADA Dashboard Configuration ===")
    location = input("Enter the Location (e.g., TK-8, Elementary, Middle, High): ").strip()
    school_year = input("Enter the School Year (e.g., 2024-2025, 2023-2024): ").strip()
    school_name = input("Enter the School Name (e.g., CCCS, Lincoln Elementary): ").strip()
    
    # Validate inputs
    if not location:
        location = "TK-8"  # Default value
        print(f"No location provided, using default: {location}")
    
    if not school_year:
        school_year = "2024-2025"  # Default value
        print(f"No school year provided, using default: {school_year}")
    
    if not school_name:
        school_name = "CCCS"  # Default value
        print(f"No school name provided, using default: {school_name}")
    
    print(f"Configuration set - Location: {location}, School Year: {school_year}, School: {school_name}")
    print("=" * 50)

    # Try to find the attendance file automatically
    import glob
    desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
    downloads_path = os.path.join(os.path.expanduser("~"), "Downloads")
    
    # Look for attendance files in common locations
    search_patterns = [
        os.path.join(desktop_path, "PrintMonthlyAttendanceSummaryTotals*.xlsx"),
        os.path.join(downloads_path, "PrintMonthlyAttendanceSummaryTotals*.xlsx"),
        "PrintMonthlyAttendanceSummaryTotals*.xlsx"
    ]
    
    file_path = None
    for pattern in search_patterns:
        files = glob.glob(pattern)
        if files:
            file_path = files[0]  # Use the first match
            print(f"Found attendance file: {file_path}")
            break
    
    if not file_path:
        print("\nNo attendance file found automatically.")
        file_path = input("C:\\Users\\Shawn\\Downloads\\PrintMonthlyAttendanceSummaryTotals_20251027_102942_46646e5.xlsx").strip()
        
        if not os.path.exists(file_path):
            print(f"ERROR: File not found at {file_path}")
            print("Please check the file path and try again.")
            return
    
    print(f"Using attendance file: {file_path}")

    target_values = {
        "Program C Charter Resident": "Prog_C",
        "Program C Charter Resident -  Transitional Kindergarten(TK)": "Prog_C_TK",
        "Program N Non-Resident Charter": "Prog_N",
        "Program N Non-Resident Charter -  Transitional Kindergarten(TK)": "Prog_N_TK",
        "Program J Indep Study Charter Resident": "Prog_J",
        "Program J Indep Study Charter Non-Resident -  Transitional Kindergarten(TK)": "Prog_J_TK",
        "Program K Indep Study Charter Non-Resident": "Prog_K",
        "Program K Indep Study Charter Non-Resident -  Transitional Kindergarten(TK)": "Prog_K_TK",
    }

    # Read the Excel file into a DataFrame
    df = pd.read_excel(file_path, header=None)  # Specify header=None to indicate no header

    # Initialize dictionaries to store start and stop indices for each target value
    target_indices = {value: {"start": None, "stop": None} for value in target_values.values()}
    # Find row numbers for each target value and update dictionaries with start and stop indices
    for target_value, dict_name in target_values.items():
        rows = find_row_with_value(df, target_value)
        start, stop = find_start_stop_indices(rows)
        target_indices[dict_name]["start"] = start
        target_indices[dict_name]["stop"] = stop

    prog_c_start = target_indices["Prog_C_TK"]["start"]
    prog_c_tk_stop = target_indices["Prog_N"]["start"]
    if prog_c_start is not None and prog_c_tk_stop is not None:
        target_indices["Prog_C"]["stop"] = prog_c_start - 1
        
    if prog_c_tk_stop is not None:
        target_indices["Prog_C_TK"]["stop"] = prog_c_tk_stop - 1

    prog_n_tk_start = target_indices["Prog_N_TK"]["start"]
    if prog_n_tk_start is not None:
        target_indices["Prog_N"]["stop"] = prog_n_tk_start - 1

    programs = ["Prog_N_TK", "Prog_J", "Prog_K"]
    for i in range(len(programs) - 1):
        start = target_indices[programs[i]]["start"]
        next_start = target_indices[programs[i + 1]]["start"]
        if start is not None and next_start is not None:
            target_indices[programs[i]]["stop"] = next_start - 1
         
    print("Current start and stop indices of programs:")
    for program, indices in target_indices.items():
        start = indices.get("start", "Not defined")
        stop = indices.get("stop", "Not defined")
        print(f"Program: {program}, Start: {start}, Stop: {stop}")

    for program in target_indices.keys():
        confirm = input(f"Are the saved start and stop indices for {program} correct? (yes/no): ").lower()
        if confirm == "no":
            while True:
                user_input = input(f"Enter new start and stop indices for {program} separated by a comma (e.g., 'start, stop'): ")
                new_indices = user_input.split(',')
                
                if len(new_indices) == 2:
                    start = int(new_indices[0].strip()) if new_indices[0].strip().lower() != "none" else None
                    stop = int(new_indices[1].strip()) if new_indices[1].strip().lower() != "none" else None
                    target_indices[program]["start"] = start
                    target_indices[program]["stop"] = stop
                    break
                else:
                    print("Invalid input. Please enter valid integer start and stop indices separated by a comma.")
    
    for target_value, indices in target_indices.items():
        print(f"{target_value}: {indices}")

    number_occurrences = {}
    for i in range(1, 13):
        occurrences = find_occurrences_of_number(df, i)
        number_occurrences[i] = occurrences
        print(f"Occurrences of Month '{i}' in rows:")
        print(occurrences)

    created_fields = check_occurrences_and_create_fields(number_occurrences, target_indices, df)
    
    print(f"\nTotal created fields: {len(created_fields)}")
    print("\nCreated Fields:")
    table = PrettyTable(["Field Name", "APA Value", "ADA Percentage"])
    for field_name, values in created_fields.items():
        apa_value, ada_percentage = values
        table.add_row([field_name, apa_value, ada_percentage])

    print(table)
    
    # Debug: Show some sample field names
    if created_fields:
        print(f"\nSample field names:")
        for i, field_name in enumerate(list(created_fields.keys())[:5]):
            print(f"  {i+1}: {field_name}")
    else:
        print("\n⚠️  WARNING: No fields were created. Check if:")
        print("  1. The Excel file exists and can be read")
        print("  2. The program names in target_values match what's in the Excel file")
        print("  3. The column indices (2, 4, 39, 47) are correct for your Excel structure")

    # Generate CSV from created_fields
    parse_data_to_csv(created_fields, school_year, location, school_name)

if __name__ == "__main__":
    main()
    desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
    file_path = os.path.join(desktop_path, "output.csv")
    os.startfile(file_path)

