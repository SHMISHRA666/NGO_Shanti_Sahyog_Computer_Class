import pandas as pd
import re

def standardize_columns(df, sheet_name, file_name):
    """
    Standardize column names across all sheets
    """
    # Create a mapping for column name standardization
    column_mapping = {
        'S NO': 'S_NO',
        'S No': 'S_NO',
        'ADM \nNO': 'ADM_NO',
        'ADM NO': 'ADM_NO',
        'ADMISSION NO': 'ADM_NO',
        'STUDENT': 'STUDENT',
        'FATHER\nHUSBAND': 'FATHER_HUSBAND',
        'Father/                                  Husband': 'FATHER_HUSBAND',
        'FATHER/HUSBAND': 'FATHER_HUSBAND',
        'COURSE': 'COURSE',
        'DURATION': 'DURATION',
        'ADDRESS': 'ADDRESS',
        'MOB': 'MOBILE',
        'MOBILE': 'MOBILE',
        'E - MAIL': 'EMAIL',
        'E-MAIL': 'EMAIL',
        'EDUCATION': 'EDUCATION',
        'QUALIFICATION': 'EDUCATION',
        'CURRENT STATUS': 'CURRENT_STATUS',
        'CURRENT\n STATUS': 'CURRENT_STATUS',
        'MONTHLY INCOME': 'MONTHLY_INCOME',
        'MONTHLY INCOME (RS.)': 'MONTHLY_INCOME',
        'MONTHLY INCOME (IN RS)': 'MONTHLY_INCOME',
        'MONTHLY \nINCOME (RS.)': 'MONTHLY_INCOME'
    }
    
    # Apply the mapping
    df.columns = df.columns.map(lambda x: column_mapping.get(x.strip(), x.strip()))
    
    # Replace spaces with underscores in remaining column names
    df.columns = df.columns.str.replace(' ', '_')
    
    # Add COURSE column for NIIT sheets if it doesn't exist
    if 'COURSE' not in df.columns and 'NIIT' in file_name:
        df['COURSE'] = 'NIIT'
        df['DURATION'] = '3 Months'
    
    return df

def extract_year_from_sheet(sheet_name):
    """
    Extract year from sheet name in format like 2022-2023
    """
    # Look for patterns like "2021-22", "2022-23", "2023-24", "2024-25"
    # or "Batch 2021-22", "Batch (2021-22)", etc.
    patterns = [
        r'(\d{4})\s*-\s*(\d{2})',  # 2021-22, 2022-23, 2024 - 25, etc.
        r'(\d{4})\s*-\s*(\d{4})',  # 2021-2022, 2022-2023, etc.
        r'Batch\s*\(?(\d{4})\s*-\s*(\d{2})\)?',  # Batch (2021-22), Batch 2024 - 25
        r'Batch\s*(\d{4})\s*-\s*(\d{2})',  # Batch 2021-22
    ]
    
    for pattern in patterns:
        match = re.search(pattern, sheet_name)
        if match:
            start_year = match.group(1)
            end_year = match.group(2)
            
            # If end_year is 2 digits, convert to 4 digits
            if len(end_year) == 2:
                end_year = start_year[:2] + end_year
            
            return f"{start_year}-{end_year}"
    
    return None

def create_unique_identifier(df):
    """
    Create a unique identifier from the combination of specified columns
    """
    # Define the columns to combine
    columns_to_combine = ['ADM_NO', 'STUDENT', 'FATHER_HUSBAND', 'COURSE', 'DURATION', 'YEAR']
    
    # Check if all required columns exist
    missing_columns = [col for col in columns_to_combine if col not in df.columns]
    if missing_columns:
        print(f"Warning: Missing columns for unique identifier: {missing_columns}")
        return df
    
    # Create unique identifier by combining all specified columns
    df['UNIQUE_ID'] = df[columns_to_combine].astype(str).agg('_'.join, axis=1)
    
    return df

def reorder_columns(df):
    """
    Reorder columns to ensure consistent order across all sheets
    """
    # Define the standard column order
    standard_order = [
        'S_NO',
        'ADM_NO', 
        'STUDENT',
        'FATHER_HUSBAND',
        'COURSE',
        'DURATION',
        'ADDRESS',
        'MOBILE',
        'EMAIL',
        'EDUCATION',
        'CURRENT_STATUS',
        'MONTHLY_INCOME',
        'YEAR',
        'UNIQUE_ID'
    ]
    
    # Get existing columns that are in the standard order
    existing_columns = [col for col in standard_order if col in df.columns]
    
    # Add any additional columns that might exist
    additional_columns = [col for col in df.columns if col not in standard_order]
    final_order = existing_columns + additional_columns
    
    # Reorder the dataframe
    return df[final_order]

def clean_course_names(df):
    """
    Clean and standardize course names to combine similar courses
    """
    # Create a mapping for course name standardization
    course_mapping = {
        'BASIC SKILLS': 'Basic Skills',
        'Basic Skills': 'Basic Skills',
        'Basic Skill': 'Basic Skills',
        '        Basic Skill': 'Basic Skills',
        'BASIC SKIILS': 'Basic Skills',
        'BASIC + TALLY': 'Basic + Tally',
        'BASIC +Tally': 'Basic + Tally',
        'Tally': 'Tally',
        'DIT': 'DIT',
        'DTP': 'DTP',
        'NIIT': 'NIIT'
    }
    
    # Apply the mapping
    df['COURSE'] = df['COURSE'].map(lambda x: course_mapping.get(str(x).strip(), str(x).strip()))
    
    return df

def clean_duration(df):
    """
    Standardize duration values.
    """
    duration_mapping = {
        '3 months': '3 Months',
        '3 month': '3 Months',
        '3 Months': '3 Months',
        '3 Month': '3 Months',
        '6 months': '6 Months',
        '6 month': '6 Months',
        '6 Months': '6 Months',
        '6 Month': '6 Months',
        '1 year': '1 Year',
        '1 Year': '1 Year',
        '12 months': '1 Year',
        '12 Months': '1 Year',
        '12 month': '1 Year',
        '12 Month': '1 Year',
    }
    df['DURATION'] = df['DURATION'].map(lambda x: duration_mapping.get(str(x).strip(), str(x).strip()))
    return df

def extract_gender(df):
    """
    Extract gender from the FATHER_HUSBAND column and create a new GENDER column.
    """
    def get_gender(value):
        value = str(value).lower()
        if 'w/o' in value or 'wife of' in value or 'wife' in value:
            return 'Female'
        elif 's/o' in value or 'son of' in value or 'son' in value:
            return 'Male'
        elif 'd/o' in value or 'daughter of' in value or 'daughter' in value:
            return 'Female'
        elif 'h/o' in value or 'husband of' in value or 'husband' in value:
            return 'Male'
        else:
            return 'Unknown'
    df['GENDER'] = df['FATHER_HUSBAND'].apply(get_gender)
    return df

def add_employment_status(df):
    """
    Add a new column EMPLOYMENT_STATUS: 'Employed' if MONTHLY_INCOME contains any digit other than zero, else 'Not Employed'.
    """
    def get_status(val):
        if pd.isnull(val) or str(val).strip() == '':
            return 'Not Employed'
        val_str = str(val)
        # If there is any digit other than zero, it's employed
        if re.search(r'[1-9]', val_str):
            return 'Employed'
        return 'Not Employed'
    df['EMPLOYMENT_STATUS'] = df['MONTHLY_INCOME'].apply(get_status)
    return df

# Read the Excel files
excel_files = ["Updated (07 -10 -2024) Batch 2021 to 2024.xlsx", "EDITED NIIT 10 sept 2024.xlsx"]

# Dictionary to store file names as keys and their sheet names as values
file_sheet_dict = {}

# Get all sheet names for each file
for file in excel_files:
    try:
        xls = pd.ExcelFile(file)
        file_sheet_dict[file] = xls.sheet_names
    except Exception as e:
        print(f"Error reading {file}: {str(e)}")

# Dictionary to store dataframes for each sheet
sheet_data = {}

# Read each sheet and store in dictionary
for file, sheets in file_sheet_dict.items():
    for sheet in sheets:
        print(f"\nReading sheet: {sheet} from {file}")
        df = pd.read_excel(file, sheet_name=sheet)
        
        # Standardize column names
        df = standardize_columns(df, sheet, file)
        
        # Extract year from sheet name
        year = extract_year_from_sheet(sheet)
        print(year)
        if year:
            df['YEAR'] = year

        # Clean course names
        df = clean_course_names(df)
        
        # Clean duration
        df = clean_duration(df)
        
        # Create unique identifier
        df = create_unique_identifier(df)
        
        # Reorder columns to ensure consistent order
        df = reorder_columns(df)
        
        # Extract gender
        df = extract_gender(df)
        
        # Add employment status
        df = add_employment_status(df)
        
        
        sheet_data[f"{file}_{sheet}"] = df
        
        # Print standardized column names for each sheet
        print(f"Standardized columns in {sheet}:")
        print(df.columns.tolist())

# print(sheet_data.values())
 