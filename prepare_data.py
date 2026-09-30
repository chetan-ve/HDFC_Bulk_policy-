import pandas as pd
import numpy as np
import os

def prepare_filtered_lot1():
    input_file = r'C:\Users\chetan.p\Downloads\HLI_spp.xlsx'
    format_file = r'C:\Users\chetan.p\Downloads\remaining120.xlsx'
    output_file = r'c:\Users\chetan.p\Documents\script\assets\lot1_pt_gt_25.xlsx'
    
    print(f"Reading {input_file} (sheet: 'lot 1')...")
    df_raw = pd.read_excel(input_file, sheet_name='lot 1')
    print(f"Total rows in lot 1: {len(df_raw)}")
    
    # 1. Filter PT > 25
    df = df_raw[df_raw['PT'] > 25].copy().reset_index(drop=True)
    print(f"Rows after PT > 25 filter: {len(df)}")
    
    # 2. Rename Last Premium Paid Date -> Last Premium Date
    if 'Last Premium Paid Date' in df.columns:
        df.rename(columns={'Last Premium Paid Date': 'Last Premium Date'}, inplace=True)
        
    # 3. Standardize Gender to M/F
    gender_map = {'Male': 'M', 'Female': 'F', 'M': 'M', 'F': 'F'}
    df['Gender'] = df['Gender'].map(gender_map).fillna(df['Gender'])
    
    # 4. Handle pw (Partial Withdrawal)
    if 'pw' in df.columns:
        df['pw'] = df['pw'].fillna(0)
    else:
        df['pw'] = 0
        
    # 5. Handle level_cover_period
    # If missing, fallback to PPT (as seen in format file for LC policies)
    if 'level_cover_period' in df.columns:
        df['level_cover_period'] = df['level_cover_period'].fillna(df['PPT']).astype(int)
    else:
        df['level_cover_period'] = df['PPT'].astype(int)
        
    # 6. Fund names and allocations
    # Fill missing Fund allocations with 0.0
    for i in range(1, 8):
        alloc_col = f'Fund alloc {i}'
        name_col = f'Fund {i}'
        if alloc_col in df.columns:
            df[alloc_col] = df[alloc_col].fillna(0.0)
        else:
            df[alloc_col] = 0.0
            
        if name_col in df.columns:
            # If fund name is missing or whitespace, fill with 'Flexi Cap'
            df[name_col] = df[name_col].astype(str).str.strip().replace({'': 'Flexi Cap', 'nan': 'Flexi Cap', 'None': 'Flexi Cap'})
        else:
            df[name_col] = 'Flexi Cap'
            
    # 7. Reset Sr no from 1 to N
    df['Sr no'] = range(1, len(df) + 1)
    
    # 8. Match exact columns of remaining120.xlsx
    df_fmt = pd.read_excel(format_file, nrows=1)
    target_columns = list(df_fmt.columns)
    
    # Ensure all target columns exist in df
    for col in target_columns:
        if col not in df.columns:
            print(f"Adding missing column: {col}")
            df[col] = np.nan
            
    df_final = df[target_columns]
    
    os.makedirs(os.path.dirname(output_file), exist_ok=True)
    df_final.to_excel(output_file, index=False, sheet_name='Sheet1')
    print(f"Successfully saved {len(df_final)} policies to {output_file}")
    print(f"Columns: {list(df_final.columns)}")
    return output_file

if __name__ == '__main__':
    prepare_filtered_lot1()
