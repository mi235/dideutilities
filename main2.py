# -*- coding: utf-8 -*-
import pandas as pd
import glob
import os


def merge_tables_1_to_5(folder_path, output_path):
    """
    Merge Table 1 through Table 5 from multiple Excel files into one consolidated file,
    shifting non-empty cells to the left in each row and saving to final.xlsx.

    Args:
        folder_path (str): Path to folder containing Excel files
        output_path (str): Path for output consolidated file (e.g., E:\final.xlsx)
    """
    if not os.path.isdir(folder_path):
        raise ValueError(f"Folder path does not exist: {folder_path}")

    # Support both .xls and .xlsx files
    patterns = [os.path.join(folder_path, "*.xls"), os.path.join(folder_path, "*.xlsx")]
    files = []
    for pattern in patterns:
        files.extend(glob.glob(pattern))

    if not files:
        print("No Excel files found in the specified folder.")
        return

    combined_data = []
    error_files = []
    missing_tables_report = {f"Table {i}": [] for i in range(1, 6)}

    print(f"\nProcessing {len(files)} Excel files for Tables 1-5...")

    for file_path in files:
        try:
            ext = os.path.splitext(file_path)[1].lower()
            engine = 'openpyxl' if ext == '.xlsx' else 'xlrd'

            with pd.ExcelFile(file_path, engine=engine) as xl:
                available_sheets = xl.sheet_names

                for table_num in range(1, 6):
                    table_name = f"Table {table_num}"
                    if table_name in available_sheets:
                        df = pd.read_excel(file_path, sheet_name=table_name, engine=engine)

                        if not df.empty:
                            df['source_file'] = os.path.basename(file_path)
                            df['source_table'] = table_name
                            combined_data.append(df)
                    else:
                        missing_tables_report[table_name].append(os.path.basename(file_path))

        except Exception as e:
            error_files.append((os.path.basename(file_path), str(e)))

    if combined_data:
        final_df = pd.concat(combined_data, ignore_index=True)

        # Μετακίνηση τιμών προς τα αριστερά σε κάθε γραμμή
        final_df = final_df.apply(
            lambda row: pd.Series(row.dropna().tolist() + [''] * (len(row) - len(row.dropna()))),
            axis=1
        )

        try:
            final_df.to_excel(output_path, index=False, header=False, engine='openpyxl')
            print(f"\n✅ Saved cleaned file with shifted values to: {output_path}")
        except Exception as e:
            print(f"\n❌ Failed to save output file: {str(e)}")
    else:
        print("\nNo data found to merge.")

    # Generate reports
    print("\n=== Processing Report ===")

    for table_name, missing_files in missing_tables_report.items():
        if missing_files:
            print(f"\n⚠️ {table_name} missing in {len(missing_files)} files. Examples:")
            for f in missing_files[:3]:
                print(f" - {f}")
            if len(missing_files) > 3:
                print(f" - (and {len(missing_files) - 3} more)")

    if error_files:
        print(f"\n❌ Errors in {len(error_files)} files. Examples:")
        for f, e in error_files[:3]:
            print(f" - {f}: {e}")
        if len(error_files) > 3:
            print(f" - (and {len(error_files) - 3} more)")


def merge_single_table(folder_path, target_sheet, output_path):
    """
    Merge a single table from multiple Excel files into one consolidated file.

    Args:
        folder_path (str): Path to folder containing Excel files
        target_sheet (str): Name of sheet to merge
        output_path (str): Path for output consolidated file
    """
    pattern = os.path.join(folder_path, "*.xlsx")
    combined_data = []

    for file_path in glob.glob(pattern):
        try:
            xl = pd.ExcelFile(file_path, engine='openpyxl')
            if target_sheet in xl.sheet_names:
                df = pd.read_excel(file_path, sheet_name=target_sheet, engine='openpyxl')
                df['source_file'] = os.path.basename(file_path)
                combined_data.append(df)
                print(f"Read: {os.path.basename(file_path)}")
            else:
                print(f"⚠️ Sheet '{target_sheet}' not found in: {os.path.basename(file_path)}")
        except Exception as e:
            print(f"❌ Error in {os.path.basename(file_path)}: {str(e)}")

    if combined_data:
        final_df = pd.concat(combined_data, ignore_index=True)
        final_df.to_excel(output_path, index=False, engine='openpyxl')
        print(f"\n🎉 Successfully merged {len(combined_data)} files into: {output_path}")
    else:
        print("No data found to merge.")


if __name__ == "__main__":
    folder_path = r"C:\Users\user\Downloads\KK"
    output_path = r"E:\final.xlsx"

    # Συγχώνευση Tables 1 έως 5 και αποθήκευση στο τελικό αρχείο
    merge_tables_1_to_5(folder_path, output_path)

    # Εναλλακτική χρήση για ένα μόνο φύλλο
    # merge_single_table(folder_path, "Table 2", r"E:\merged_table_2.xlsx")
