import pandas as pd
import os

# === SETTINGS ===
source_folder = "/mnt/d/Work/Job/CPNS/Perbendaharaan/Rekon Pajak 2025/JUNI/test-file"
combined_file_path = "/mnt/d/Work/Job/CPNS/Perbendaharaan/Rekon Pajak 2025/JUNI/test-combined.xlsx"
sheet_name = "Combined Data"

# === Start merging ===
files_to_merge = [
    f for f in os.listdir(source_folder)
    if (f.endswith('.xlsx') or f.endswith('.xls')) and not f.startswith('~$')
]

print(f"📁 Found {len(files_to_merge)} Excel files.")
combined_df = pd.DataFrame()
reference_columns = None

for file in files_to_merge:
    file_path = os.path.join(source_folder, file)
    print(f"\n📄 Processing: {file}")

    try:
        df = pd.read_excel(file_path, skiprows=4, dtype=str)

        df = df.dropna(how='all')

        df = df[~df.iloc[:, 0].astype(str).str.upper().isin(["JUMLAH", "TOTAL"])]

        df.columns = df.columns.str.strip().str.upper().str.replace(r"\s+", " ", regex=True)

        if df.columns.duplicated().any():
            print(f"⚠️ Duplicate columns detected in {file}, dropping duplicates:")
            print("🔁 Dropped:", df.columns[df.columns.duplicated()].tolist())
            df = df.loc[:, ~df.columns.duplicated()]

        df["Source File"] = file

        if combined_df.empty:
            reference_columns = df.columns.tolist()
            combined_df = df
            print("✅ Used as reference column set.")
        else:
            expected_cols = [col for col in reference_columns if col != "Source File"]
            missing_cols = [col for col in expected_cols if col not in df.columns]

            if not missing_cols:
                reordered_cols = expected_cols + ["Source File"]
                df = df[reordered_cols]
                combined_df = pd.concat([combined_df, df], ignore_index=True)
                print(f"✅ Added {len(df)} rows.")
            else:
                print(f"⚠️ Skipped {file} due to missing columns: {missing_cols}")
                print("🧾 Available columns:", df.columns.tolist())

    except Exception as e:
        print(f"❌ Error reading {file}: {e}")

# === Export combined result ===
if not combined_df.empty:
    with pd.ExcelWriter(combined_file_path, engine='openpyxl', mode='w') as writer:
        combined_df.to_excel(writer, sheet_name=sheet_name, index=False)
    print(f"\n✅ Done! Combined file saved to:\n{combined_file_path}")
    print(f"📊 Total rows combined: {len(combined_df)}")
else:
    print("⚠️ No data combined. Please check the input files.")
