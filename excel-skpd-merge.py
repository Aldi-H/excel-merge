import pandas as pd
import glob
import os
from openpyxl import load_workbook

folder_path = "/mnt/d/Work/Job/CPNS/Perbendaharaan/SKPD/2025"

excel_files = glob.glob(os.path.join(folder_path, "*.xlsx")) + \
            glob.glob(os.path.join(folder_path, "*.xls"))

if not excel_files:
    print(f"No Excel files found in {folder_path}")
else:
    print(f"Found {len(excel_files)} Excel file(s)")
    
    dfs = []
    
    for file in excel_files:
        print(f"Reading: {os.path.basename(file)}")

        df = pd.read_excel(
            file,
            keep_default_na=False,
            na_filter=False
        )
        
        # Konversi kolom tertentu ke string jika diperlukan
        # if 'NIP' in df.columns:
        #     df['NIP'] = df['NIP'].astype(str)
        # if 'NO. REKENING' in df.columns:
        #     df['NO. REKENING'] = df['NO. REKENING'].astype(str)

        df.insert(0, "SOURCE FILE", os.path.basename(file))
        dfs.append(df)

    merged_df = pd.concat(dfs, ignore_index=True)  

    output_path = os.path.join(folder_path, "SKPD_2025.xlsx")
    merged_df.to_excel(output_path, index=False, engine='openpyxl')

    print(f"\nMerged {len(dfs)} files successfully!")
    print(f"Output saved to: {output_path}")
    print(f"Total rows: {len(merged_df)}")