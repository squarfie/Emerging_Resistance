#!D:\PROGRAMMING\whonet_column_checker\check\python.exe

import pandas as pd
import os
from tabulate import tabulate
import openpyxl

def check_columns(input_folder, input_file1, input_file2, output_folder):
    input1_path = os.path.join(input_folder, input_file1)
    input2_path = os.path.join(input_folder, input_file2)
    
    # Load input_file1 (Excel)
    try:
        df_input1 = pd.read_excel(input1_path)
        input1_columns = df_input1.columns.tolist()
    except Exception as e:
        print(f"\n❌ Error reading the first input Excel file: {e}")
        return
    
    # Load input_file2 (Excel)
    try:
        df_input2 = pd.read_excel(input2_path)
        input2_columns = df_input2.columns.tolist()
    except Exception as e:
        print(f"\n❌ Error reading the second input Excel file: {e}")
        return

    # Check and align columns
    missing_df1_columns = [col for col in input2_columns if col not in input1_columns]
    missing_df2_columns = [col for col in input1_columns if col not in input2_columns]

    for col in missing_df1_columns:
        df_input1[col] = pd.NA
    for col in missing_df2_columns:
        df_input2[col] = pd.NA

    df_input2 = df_input2[df_input1.columns]

    # Validate AccessionNo
    if 'AccessionNo' not in df_input1.columns or 'AccessionNo' not in df_input2.columns:
        print("❌ 'AccessionNo' column must exist in both files.")
        return

    accession1 = df_input1['AccessionNo'].dropna().astype(str)
    accession2 = df_input2['AccessionNo'].dropna().astype(str)

    # Match & Unmatch
    df1_matched = df_input1[df_input1['AccessionNo'].astype(str).isin(accession2)]
    df2_unmatched = df_input2[~df_input2['AccessionNo'].astype(str).isin(accession1)]
    df1_unmatched = df_input1[~df_input1['AccessionNo'].astype(str).isin(accession2)]

    combined_df = pd.concat([df1_matched, df2_unmatched, df1_unmatched], ignore_index=True)

    # Save output
    output_file_name = f"{os.path.splitext(input_file1)[0]}_combined_output.xlsx"
    output_file = os.path.join(output_folder, output_file_name)

    try:
        combined_df.to_excel(output_file, index=False)
        print(f"\n✅ Combined data saved to: {output_file}")
    except Exception as e:
        print(f"\n❌ Error saving the output file: {e}")
        return

    # Summary
    print("\n📊 Summary:")
    print(f"- Rows in Input 1: {len(df_input1)}")
    print(f"- Rows in Input 2: {len(df_input2)}")
    print(f"- Matched rows: {len(df1_matched)}")
    print(f"- Unmatched from Input 1: {len(df1_unmatched)}")
    print(f"- Unmatched from Input 2: {len(df2_unmatched)}")
    print(f"- Total Combined Rows: {len(combined_df)}")

if __name__ == "__main__":
    print("=== WHONET Dual File Merger (Based on AccessionNo) ===\n")
    
    # Hardcoded paths
    input_folder = r"D:\Emerging_Resistance\Emerging_Res_Tool\input"
    output_folder = r"D:\Emerging_Resistance\Emerging_Res_Tool\output"
    
    # Only ask for filenames
    input_file1 = input("📄 Enter first Excel filename (e.g., file1.xlsx): ").strip()
    input_file2 = input("📄 Enter second Excel filename (e.g., file2.xlsx): ").strip()

    check_columns(input_folder, input_file1, input_file2, output_folder)
