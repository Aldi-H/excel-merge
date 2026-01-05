import pandas as pd
import os
from glob import glob
from tkinter import Tk, filedialog

folder_path = "/mnt/d/Work/Job/CPNS/Perbendaharaan/Rekon Pajak 2025/SEPTEMBER/Rekon/KPPN"
output_path = "/mnt/d/Work/Job/CPNS/Perbendaharaan/Rekon Pajak 2025/SEPTEMBER/Rekon/KPPN"

def select_files():
    """Open file dialog to manually select Excel files"""
    root = Tk()
    root.withdraw()  # Hide the main window
    root.attributes('-topmost', True)  # Bring dialog to front
    
    print("Opening file selection dialog...")
    files = filedialog.askopenfilenames(
        title="Select Excel files to merge",
        initialdir=folder_path,
        filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
    )
    
    root.destroy()
    return list(files)

def get_sheet_names(file_path):
    """Get all sheet names from an Excel file"""
    xl_file = pd.ExcelFile(file_path)
    return xl_file.sheet_names

def select_sheets_from_files(file_paths):
    """Interactive sheet selection for each file"""
    selected_sheets = []
    
    for file_path in file_paths:
        print(f"\n{'='*60}")
        print(f"File: {os.path.basename(file_path)}")
        print('='*60)
        
        try:
            sheet_names = get_sheet_names(file_path)
            
            if not sheet_names:
                print("No sheets found in this file. Skipping...")
                continue
            
            print(f"Available sheets ({len(sheet_names)}):")
            for idx, sheet in enumerate(sheet_names, 1):
                print(f"  {idx}. {sheet}")
            
            print("\nOptions:")
            print("  - Enter sheet numbers separated by comma (e.g., 1,3,5)")
            print("  - Enter 'all' to select all sheets")
            print("  - Enter 'skip' to skip this file")
            
            selection = input("\nYour choice: ").strip().lower()
            
            if selection == 'skip':
                print(f"Skipping {os.path.basename(file_path)}")
                continue
            elif selection == 'all':
                for sheet in sheet_names:
                    selected_sheets.append({
                        'file': file_path,
                        'sheet': sheet
                    })
                print(f"Selected all {len(sheet_names)} sheets")
            else:
                try:
                    indices = [int(x.strip()) for x in selection.split(',')]
                    for idx in indices:
                        if 1 <= idx <= len(sheet_names):
                            selected_sheets.append({
                                'file': file_path,
                                'sheet': sheet_names[idx - 1]
                            })
                        else:
                            print(f"Warning: Index {idx} is out of range. Skipping...")
                    print(f"Selected {len([i for i in indices if 1 <= i <= len(sheet_names)])} sheet(s)")
                except ValueError:
                    print("Invalid input. Skipping this file...")
                    
        except Exception as e:
            print(f"Error reading file: {e}")
            continue
    
    return selected_sheets

def format_dataframe(df):
    """Optional: Add any formatting you need here"""
    # Example formatting - customize as needed
    # df = df.dropna(how='all')  # Remove completely empty rows
    # df = df.reset_index(drop=True)
    return df

def merge_selected_sheets(selected_sheets):
    """Merge selected sheets while preserving specified data types"""
    dfs = []

    for item in selected_sheets:
        file_path = item['file']
        sheet_name = item['sheet']
        
        try:
            df = pd.read_excel(
                file_path,
                sheet_name=sheet_name,
                # skiprows=4,
                dtype={
                    "NPWP REKANAN/BENDAHARA": str,
                    "ID BILLING": str,
                    "NTPN": str
                }
            )

            df.insert(0, "SOURCE FILE", os.path.basename(file_path))
            df.insert(1, "SOURCE SHEET", sheet_name)
            dfs.append(df)
            
            print(f"✓ Loaded: {os.path.basename(file_path)} - {sheet_name} ({len(df)} rows)")

        except Exception as e:
            print(f"✗ Failed: {os.path.basename(file_path)} - {sheet_name}")
            print(f"  Error: {e}")
            continue

    if not dfs:
        raise ValueError("No data read from any sheet.")

    merged = pd.concat(dfs, ignore_index=True)
    return format_dataframe(merged)

def main():
    """Main function to execute the merge process"""
    try:
        print("="*60)
        print("EXCEL SHEET MERGER")
        print("="*60)
        print("\nChoose your option:")
        print("1. Auto-merge all sheets from all Excel files in folder")
        print("2. Select specific files and sheets manually")
        
        choice = input("\nEnter choice (1 or 2): ").strip()
        
        if choice == "1":
            # Get all Excel files from the folder
            excel_files = glob(os.path.join(folder_path, "*.xlsx"))
            
            # Filter out temporary files (starting with ~$)
            excel_files = [f for f in excel_files if not os.path.basename(f).startswith("~$")]
            
            if not excel_files:
                print(f"No Excel files found in: {folder_path}")
                return
            
            # Select all sheets from all files
            selected_sheets = []
            for file in excel_files:
                sheet_names = get_sheet_names(file)
                for sheet in sheet_names:
                    selected_sheets.append({
                        'file': file,
                        'sheet': sheet
                    })
                    
        elif choice == "2":
            # Manually select files
            excel_files = select_files()
            
            if not excel_files:
                print("No files selected. Exiting...")
                return
            
            # Manually select sheets
            selected_sheets = select_sheets_from_files(excel_files)
            
            if not selected_sheets:
                print("\nNo sheets selected. Exiting...")
                return
        else:
            print("Invalid choice. Exiting...")
            return
        
        print(f"\n{'='*60}")
        print(f"Total sheets to merge: {len(selected_sheets)}")
        print('='*60)
        
        # Merge all selected sheets
        print("\nMerging sheets...")
        merged_df = merge_selected_sheets(selected_sheets)
        
        # Create output filename
        output_file = os.path.join(output_path, "Merged_Data.xlsx")
        
        # Save to Excel
        print(f"\nSaving merged data to: {output_file}")
        merged_df.to_excel(output_file, index=False, engine='openpyxl')
        
        print(f"\n{'='*60}")
        print("SUCCESS!")
        print('='*60)
        print(f"Merged {len(selected_sheets)} sheet(s)")
        print(f"Total rows: {len(merged_df)}")
        print(f"Total columns: {len(merged_df.columns)}")
        print(f"Output file: {output_file}")
        
    except Exception as e:
        print(f"\n{'='*60}")
        print("ERROR!")
        print('='*60)
        print(f"{e}")
        raise

if __name__ == "__main__":
    main()