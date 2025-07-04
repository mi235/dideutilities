# -*- coding: utf-8 -*-
import pandas as pd
import os
from tkinter import Tk, filedialog


def filter_dodekanisa(input_file):
    # Φορτώνουμε όλα τα φύλλα
    try:
        xls = pd.read_excel(input_file, sheet_name=None, engine="openpyxl")
    except Exception as e:
        print(f"❌ Σφάλμα κατά την ανάγνωση του αρχείου: {e}")
        return

    if "ΟΛΑ" not in xls:
        print("❌ Το αρχείο δεν περιέχει φύλλο με όνομα 'ολα'")
        return

    df = xls["ΟΛΑ"]

    # Φιλτράρισμα: εντοπισμός της λέξης "ΔΩΔΕΚΑΝΗΣ" σε οποιοδήποτε κελί κάθε γραμμής
    mask = df.apply(lambda row: row.astype(str).str.upper().str.contains("ΔΩΔΕΚΑΝΗΣ").any(), axis=1)
    filtered_df = df[mask]

    if filtered_df.empty:
        print("⚠️ Δεν βρέθηκαν γραμμές με 'ΔΩΔΕΚΑΝΗΣ*'")
        return

    try:
        # Γράφουμε όλα τα αρχικά φύλλα + το νέο φύλλο
        with pd.ExcelWriter(input_file, engine="openpyxl", mode='a', if_sheet_exists='replace') as writer:
            filtered_df.to_excel(writer, sheet_name="ΔΩΔΕΚΑΝΗΣΑ", index=False)
            print(f"✅ Προστέθηκε νέο φύλλο 'ΔΩΔΕΚΑΝΗΣΑ' στο αρχείο: {input_file}")
    except Exception as e:
        print(f"❌ Σφάλμα κατά την αποθήκευση: {e}")


if __name__ == "__main__":
    print("📥 Επιλέξτε αρχείο Excel για επεξεργασία")

    root = Tk()
    root.withdraw()

    file_path = filedialog.askopenfilename(
        title="Επέλεξε αρχείο Excel",
        filetypes=[("Excel files", "*.xlsx")])

    if file_path:
        filter_dodekanisa(file_path)
    else:
        print("❌ Δεν επιλέχθηκε αρχείο.")
