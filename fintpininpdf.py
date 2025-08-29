import os
import camelot
import pandas as pd
from tkinter import Tk, filedialog

def pdf_to_excel(pdf_file):
    # Διαβάζουμε όλους τους πίνακες από το PDF
    tables = camelot.read_pdf(pdf_file, pages="all")

    if len(tables) == 0:
        print(f"⚠ Δεν βρέθηκε πίνακας στο {pdf_file}")
        return

    # Δημιουργούμε όνομα Excel με ίδιο όνομα αρχείου
    excel_file = os.path.splitext(pdf_file)[0] + ".xlsx"

    # Αποθήκευση όλων των πινάκων σε Excel
    with pd.ExcelWriter(excel_file, engine="openpyxl") as writer:
        for i, table in enumerate(tables):
            df = table.df
            sheet_name = f"Πίνακας_{i+1}"
            df.to_excel(writer, sheet_name=sheet_name, index=False)

    print(f"✅ Αποθηκεύτηκε: {excel_file}")


if __name__ == "__main__":
    # Άνοιγμα παραθύρου για επιλογή φακέλου
    root = Tk()
    root.withdraw()
    folder_selected = filedialog.askdirectory(title="Διάλεξε φάκελο με PDF")

    if folder_selected:
        # Όλα τα PDF του φακέλου
        pdf_files = [f for f in os.listdir(folder_selected) if f.lower().endswith(".pdf")]

        if not pdf_files:
            print("⚠ Δεν βρέθηκαν PDF στον φάκελο.")
        else:
            for pdf_file in pdf_files:
                pdf_path = os.path.join(folder_selected, pdf_file)
                pdf_to_excel(pdf_path)

            print("\n🎉 Τέλος! Όλα τα Excel δημιουργήθηκαν στον ίδιο φάκελο.")
    else:
        print("Δεν επιλέχθηκε φάκελος.")
