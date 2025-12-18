import pandas as pd
import os
import json
import sys
import re

folder_path = "/mnt/d/Work/Job/CPNS/Perbendaharaan/Rekon Pajak 2025/OKTOBER/OPD"
output_path = "/mnt/d/Work/Job/CPNS/Perbendaharaan/Rekon Pajak 2025/OKTOBER"

def list_excel_files(path):
    files = [
        os.path.join(path, f)
        for f in os.listdir(path)
        if f.lower().endswith(('.xlsx', '.xls'))
    ]

    if not files:
        raise ValueError("No Excel files found in folder.")

    # newest first
    files.sort(key=os.path.getctime, reverse=True)
    return files


def merge_excel_files(paths):
    dfs = []

    for p in paths:
        try:
            df = pd.read_excel(
                p,
                sheet_name=0,
                skiprows=4,
                dtype={
                    "KODE_AKUN_BELANJA": str,
                    "KODE_AKUN_POTONGAN_PAJAK": str,
                    "NPWP_BENDAHARA": str,
                    "ID_BILLING": str,
                    "NTPN": str
                }
            )

            dfs.append(df)

        except Exception as e:
            raise Exception(f"Failed reading file: {p}\nError: {e}")

    if not dfs:
        raise ValueError("No data read from any file.")

    merged = pd.concat(dfs, ignore_index=True)
    return format_dataframe(merged)


def format_dataframe(df):
    currency_cols = ["NILAI_BELANJA_SP2D", "JUMLAH_PAJAK"]
    for col in currency_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype('int64')

    if "KODE_AKUN_BELANJA" in df.columns:
        df["KODE_AKUN_BELANJA"] = (
            df["KODE_AKUN_BELANJA"]
            .astype(str).str.replace(".", "").str.strip()
        )

    if "NPWP_BENDAHARA" in df.columns:
        df["NPWP_BENDAHARA"] = (
            df["NPWP_BENDAHARA"]
            .astype(str).apply(lambda x: re.sub(r"[^\d]", "", x))
        )

    if "ID_BILLING" in df.columns:
        df["ID_BILLING"] = df["ID_BILLING"].astype(str).str.replace(".0", "")

    if "KODE_AKUN_POTONGAN_PAJAK" in df.columns:
        df["KODE_AKUN_POTONGAN_PAJAK"] = (
            df["KODE_AKUN_POTONGAN_PAJAK"].astype(str).str.replace("-100", "").str.strip()
        )

    return df


def main():
    try:
        excel_paths = list_excel_files(folder_path)

        merged_df = merge_excel_files(excel_paths)

        out_file = os.path.join(output_path, "REKON_PAJAK_OKTOBER_2025.xlsx")
        merged_df.to_excel(out_file, index=False)

        print(json.dumps({"success": True, "output_file": out_file}, indent=2))
        sys.exit(0)

    except Exception as e:
        print(json.dumps({"success": False, "error": str(e)}, indent=2), file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
