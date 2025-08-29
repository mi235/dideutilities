# -*- coding: utf-8 -*-
import pandas as pd
import os
from tkinter import Tk, filedialog

def merge_all_excels_all_fields(excel_paths, output_path):
    combined_data = []

    for file_path in excel_paths:
        try:
            ext = os.path.splitext(file_path)[1].lower()
            engine = 'openpyxl' if ext == '.xlsx' else 'xlrd'
            xl = pd.ExcelFile(file_path, engine=engine)
        except Exception as e:
            print(f"❌ Σφάλμα στο αρχείο '{file_path}': {e}")
            continue

        for sheet_name in xl.sheet_names:
            try:
                preview_df = xl.parse(sheet_name, header=None)
                if len(preview_df) <1:
                    print(f"⚠️ Το φύλλο '{sheet_name}' στο '{os.path.basename(file_path)}' έχει < 2 γραμμές. Αγνοείται.")
                    continue

                # Διαβάζουμε με header από1η γραμμή (index 0)
                df = xl.parse(sheet_name, header=0)
                df.columns = [str(col).strip() for col in df.columns]

                if not df.empty:
                    df.insert(0, 'source_sheet', sheet_name)
                    df.insert(0, 'source_file', os.path.basename(file_path))
                    combined_data.append(df)
                    print(f"✅ Προστέθηκε το φύλλο: {sheet_name} από {os.path.basename(file_path)}")
                else:
                    print(f"⚠️ Άδειο φύλλο: {sheet_name} στο {os.path.basename(file_path)}")
            except Exception as e:
                print(f"❌ Σφάλμα στο φύλλο '{sheet_name}': {e}")

    if combined_data:
        final_df = pd.concat(combined_data, ignore_index=True)
        try:
            final_df.to_excel(output_path, index=False, engine='openpyxl')
            print(f"\n🎉 Ολοκληρώθηκε η συγχώνευση στο: {output_path}")
        except Exception as e:
            print(f"\n❌ Σφάλμα κατά την αποθήκευση: {e}")
    else:
        print("\n⚠️ Δεν βρέθηκαν δεδομένα προς συγχώνευση.")

if __name__ == "__main__":
    print("🔄 Συγχώνευση όλων των φύλλων από πολλαπλά Excel (με όλα τα πεδία)")

    root = Tk()
    root.withdraw()

    # 📂 Επιλογή πολλών Excel αρχείων
    input_files = filedialog.askopenfilenames(
        title="📥 Επιλέξτε ένα ή περισσότερα αρχεία Excel",
        filetypes=[("Excel files", "*.xls *.xlsx")]
    )

    if not input_files:
        print("❌ Δεν επιλέχθηκαν αρχεία.")
        exit()

    # 📤 Αποθήκευση αρχείου εξόδου
    output_file = filedialog.asksaveasfilename(
        title="📤 Επιλέξτε πού θα αποθηκευτεί το αρχείο",
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")]
    )

    if not output_file:
        print("❌ Δεν επιλέχθηκε αρχείο αποθήκευσης.")
        exit()

    merge_all_excels_all_fields(input_files, output_file)
