import pandas as pd
import glob
import os
from datetime import datetime

# folder_path = "/mnt/d/Work/Job/CPNS/Perbendaharaan/OMSPAN/2026/JUNI/TAMSIL/MEI_GEL 1"
folder_path = "/mnt/d/Work/Job/CPNS/Perbendaharaan/OMSPAN/2026/JUNI/TPG/APRL_GEL 3"

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
                "BANK": str,
            })
        
        df.insert(0, "SOURCE FILE", os.path.basename(file))

        dfs.append(df)

    merged_df = pd.concat(dfs, ignore_index=True)

    merged_df.insert(1, "ANGKATAN", merged_df["NIP"].str[8:12])

    datetime = datetime.now().strftime("%Y%m%d")

    # output_path = os.path.join(folder_path, f"DAK_TAMSIL_MEI_GEL_1_{datetime}.xlsx")
    output_path = os.path.join(folder_path, f"DAK_TPG_APRIL_GEL_3_{datetime}.xlsx")
    merged_df.to_excel(output_path, index=False)
    
    print(f"\nMerged {len(dfs)} files successfully!")
    print(f"Output saved to: {output_path}")
    print(f"Total rows: {len(merged_df)}")