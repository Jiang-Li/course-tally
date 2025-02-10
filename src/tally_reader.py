import pandas as pd

def read_tally_file(file_path):
    """
    Read Excel file and find the table starting with 'Subj' cell,
    Also checks if rows are uniquely identified by key columns
    
    Args:
        file_path (str): Path to the Excel file
        
    Returns:
        tuple: (DataFrame, bool, list) - (table data, whether rows are uniquely identified, key columns used)
    """
    # First, read the sheet without headers to get all cells
    df_raw = pd.read_excel(file_path, header=None)
    
    # Find the cell containing 'Subj'
    subj_location = None
    for row_idx, row in df_raw.iterrows():
        for col_idx, cell in enumerate(row):
            if isinstance(cell, str) and cell.strip().lower() == 'subj':
                subj_location = (row_idx, col_idx)
                break
        if subj_location:
            break
    
    if not subj_location:
        raise ValueError("Could not find 'Subj' in the Excel file")
    
    # Read the Excel file again, but now starting from the 'Subj' row and column
    df = pd.read_excel(
        file_path,
        header=subj_location[0],
        usecols=range(subj_location[1], df_raw.shape[1])
    )
    
    # Keep rows that have at least one non-NaN value
    df = df.dropna(thresh=1)
    
    # Remove any columns that are completely empty
    df = df.dropna(axis=1, how='all')
    
    # Get first three columns plus Room for uniqueness check
    key_columns = list(df.iloc[:, :3].columns) + ['Room']
    print(f"\nChecking columns: {key_columns}")  # Debug print
    
    if 'Room' not in df.columns:
        print("\nWarning: 'Room' column not found in the data")
        return df, False, key_columns
        
    key_cols_df = df[key_columns]
    is_unique = len(df) == len(key_cols_df.drop_duplicates())
    print(f"\nTotal rows: {len(df)}")  # Debug print
    print(f"Unique rows: {len(key_cols_df.drop_duplicates())}")  # Debug print
    
    if not is_unique:
        print("\nDuplicate entries found:")
        print("-" * 50)
        duplicates = df[key_cols_df.duplicated(keep=False)].sort_values(by=key_columns)
        for _, row in duplicates.iterrows():
            print(f"{', '.join([f'{col}: {row[col]}' for col in key_columns])}")
        print("-" * 50)
    else:
        print("\nNo duplicates found - all rows are uniquely identified by key columns")
    
    return df, is_unique, key_columns 