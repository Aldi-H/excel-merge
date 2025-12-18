import pandas as pd
import glob
import os

# folder_path = "/mnt/d/Work/Job/CPNS/Perbendaharaan/OMSPAN/2025/DAK_TPG_NOVEMBER"
folder_path = "/mnt/d/Work/Job/CPNS/Perbendaharaan/OMSPAN/2025/DAK_TAMBAHAN_PENGHASILAN_NOVEMBER"

excel_files = glob.glob(os.path.join(folder_path, "*.xlsx")) + \
            glob.glob(os.path.join(folder_path, "*.xls"))

if not excel_files:
    print(f"No Excel files found in {folder_path}")
else:
    print(f"Found {len(excel_files)} Excel file(s)")
    
    dfs = []
    
    for file in excel_files:
        print(f"Reading: {os.path.basename(file)}")
        df = pd.read_excel(file, dtype=str)
        dfs.append(df)
    
    merged_df = pd.concat(dfs, ignore_index=True)
    
    output_path = os.path.join(folder_path, "DAK_TAMBAHAN_PENGHASILAN_NOVEMBER.xlsx")
    merged_df.to_excel(output_path, index=False)
    
    print(f"\nMerged {len(dfs)} files successfully!")
    print(f"Output saved to: {output_path}")
    print(f"Total rows: {len(merged_df)}")