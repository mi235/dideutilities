<<<<<<< HEAD
# -*- coding: utf-8 -*-
import pandas as pd
import os
from tkinter import Tk, filedialog


def merge_all_sheets(input_path, output_path):
    """
    Merges all sheets of a single Excel workbook into one sheet, adding sheet name as first column.
    """
    if not os.path.isfile(input_path):
        print(f"\n❌ Το αρχείο δεν βρέθηκε: {input_path}")
        return

    try:
        xl = pd.ExcelFile(input_path, engine='openpyxl')
    except Exception as e:
        print(f"\n❌ Δεν ήταν δυνατή η ανάγνωση του αρχείου: {e}")
        return

    combined_data = []

    for sheet_name in xl.sheet_names:
        try:
            df = xl.parse(sheet_name)
            if not df.empty:
                df.insert(0, 'source_sheet', sheet_name)
                combined_data.append(df)
                print(f"✅ Προστέθηκε το φύλλο: {sheet_name}")
            else:
                print(f"⚠️ Άδειο φύλλο: {sheet_name}")
        except Exception as e:
            print(f"❌ Σφάλμα στο φύλλο '{sheet_name}': {e}")

    if combined_data:
        final_df = pd.concat(combined_data, ignore_index=True)

        try:
            final_df.to_excel(output_path, index=False, engine='openpyxl')
            print(f"\n🎉 Η συγχώνευση ολοκληρώθηκε στο: {output_path}")
        except Exception as e:
            print(f"\n❌ Σφάλμα κατά την αποθήκευση: {e}")
    else:
        print("\n⚠️ Δεν βρέθηκαν δεδομένα προς συγχώνευση.")


if __name__ == "__main__":
    print("🔄 Συγχώνευση όλων των φύλλων ενός αρχείου Excel")

    # Απενεργοποιούμε το κενό παράθυρο Tk
    root = Tk()
    root.withdraw()

    # Επιλογή αρχείου εισόδου
    input_file = filedialog.askopenfilename(
        title="📥 Επιλέξτε το αρχείο Excel εισόδου",
        filetypes=[("Excel files", "*.xlsx")])

    if not input_file:
        print("❌ Δεν επιλέχθηκε αρχείο.")
        exit()

    # Επιλογή αρχείου εξόδου
    output_file = filedialog.asksaveasfilename(
        title="📤 Επιλέξτε που θα αποθηκευτεί το αρχείο",
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")])

    if not output_file:
        print("❌ Δεν επιλέχθηκε αρχείο αποθήκευσης.")
        exit()

    merge_all_sheets(input_file, output_file)
