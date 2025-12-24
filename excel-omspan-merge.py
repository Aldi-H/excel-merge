import pandas as pd
import glob
import os

# folder_path = "/mnt/d/Work/Job/CPNS/Perbendaharaan/OMSPAN/2025/DAK_TPG_NOVEMBER"
folder_path = "/mnt/d/Work/Job/CPNS/Perbendaharaan/OMSPAN/2025/DESEMBER/DAK_TPG_DESEMBER/KESELURUHAN"

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
            dtype={
                "NIP": str,
                "NO. REKENING": str,
            })
        
        df.insert(0, "SOURCE FILE", os.path.basename(file))

        dfs.append(df)
    
    merged_df = pd.concat(dfs, ignore_index=True)
    
    merged_df.insert(1, "ANGKATAN", merged_df["NIP"].str[8:12])
    
    output_path = os.path.join(folder_path, "DAK_TPG_DESEMBER.xlsx")
    merged_df.to_excel(output_path, index=False)
    
    print(f"\nMerged {len(dfs)} files successfully!")
    print(f"Output saved to: {output_path}")
    print(f"Total rows: {len(merged_df)}")