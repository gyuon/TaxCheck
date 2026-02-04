import pandas as pd
import numpy as np
import os

base_path = "c:\\Users\\gyuon\\Downloads\\인별납부내역자동화\\"
file1 = base_path + "※정리본_25.10.18_오류검출 (1).xlsx"
file2 = base_path + "※정리본_25.10.18_오류검출new.xlsx"

if not os.path.exists(file1) or not os.path.exists(file2):
    print("Files not found.")
    exit()

print(f"Comparing:\n1. {file1}\n2. {file2}\n")

xl1 = pd.ExcelFile(file1)
xl2 = pd.ExcelFile(file2)

print(f"Sheets in File 1: {xl1.sheet_names}")
print(f"Sheets in File 2: {xl2.sheet_names}")

common_sheets = set(xl1.sheet_names).intersection(xl2.sheet_names)

for sheet in sorted(common_sheets):
    print(f"\n{'='*30}\nComparing Sheet: '{sheet}'\n{'='*30}")
    df1 = pd.read_excel(file1, sheet_name=sheet)
    df2 = pd.read_excel(file2, sheet_name=sheet)
    
    # 1. Shape comparison
    if df1.shape != df2.shape:
        print(f"[FAIL] Shape mismatch: File1 {df1.shape} vs File2 {df2.shape}")
    else:
        print(f"[PASS] Shape matches: {df1.shape}")

    # 2. Column comparison
    if list(df1.columns) != list(df2.columns):
        print(f"[FAIL] Columns mismatch:")
        print(f"   File1: {list(df1.columns)}")
        print(f"   File2: {list(df2.columns)}")
        # Try to align columns for value comparison if possible
        common_cols = list(set(df1.columns).intersection(df2.columns))
        df1 = df1[common_cols]
        df2 = df2[common_cols]
    else:
        print(f"[PASS] Columns match")

    # 3. Value comparison
    # Align data by index/columns if needed, but here assuming row order matters.
    # Convert to same types for comparison (e.g. numeric)
    # Fill NaN with a placeholder to compare
    
    try:
        # Sort if "이름", "해당년", "해당월" columns exist to ensure alignment
        # Index column "번호" usually exists in these files, drop it for comparison if generated
        if "번호" in df1.columns:
            df1 = df1.drop(columns=["번호"])
        if "번호" in df2.columns:
            df2 = df2.drop(columns=["번호"])
            
        sort_cols = [c for c in ["이름", "해당년", "해당월", "코드1"] if c in df1.columns]
        if sort_cols:
            df1 = df1.sort_values(sort_cols).reset_index(drop=True)
            df2 = df2.sort_values(sort_cols).reset_index(drop=True)

        # Normalize types
        for col in df1.columns:
             if df1[col].dtype == object:
                 df1[col] = df1[col].astype(str).str.strip()
             if df2[col].dtype == object:
                 df2[col] = df2[col].astype(str).str.strip()

        # Numeric rounding
        numeric_cols = df1.select_dtypes(include=[np.number]).columns
        if not numeric_cols.empty:
            df1[numeric_cols] = df1[numeric_cols].fillna(0).round(0)
            df2[numeric_cols] = df2[numeric_cols].fillna(0).round(0)

        diff_mask = (df1 != df2) & ~(df1.isnull() & df2.isnull()) & ~(df1.isna() & df2.isna())
        
        # specific check for empty strings vs None
        
        if diff_mask.any().any():
            print("[FAIL] Value mismatches found!")
            diff_count = diff_mask.sum().sum()
            print(f"   Total cells differing: {diff_count}")
            
            # Show first few differences
            rows_with_diff = diff_mask.any(axis=1)
            print(f"\n   First 5 rows with differences:")
            print(df1[rows_with_diff].head())
            print("\n   Vs File 2:")
            print(df2[rows_with_diff].head())
        else:
            print("[PASS] All values match!")
            
    except Exception as e:
        print(f"[ERROR] Error during value comparison: {e}")
