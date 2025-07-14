# -*- coding: utf-8 -*-
import os
from tkinter import Tk, filedialog
import tabula
import pandas as pd

def pdf_to_excel():
    # Επιλογή αρχείου PDF
    root = Tk()
    root.withdraw()
    pdf_file = filedialog.askopenfilename(
        title="Επιλέξτε PDF αρχείο",
        filetypes=[("PDF Files", "*.pdf")]
    )

    if not pdf_file:
        print("❌ Δεν επιλέχθηκε αρχείο PDF.")
        return

    try:
        print(f"📄 Ανάγνωση αρχείου PDF: {pdf_file}")
        # Ανάγνωση όλων των πινάκων από όλες τις σελίδες
        dfs = tabula.read_pdf(pdf_file, pages='all', multiple_tables=True, lattice=True)

        if not dfs:
            print("⚠️ Δεν βρέθηκαν πίνακες στο PDF.")
            return

        # Δημιουργία Excel αρχείου
        output_excel = os.path.splitext(pdf_file)[0] + "_converted.xlsx"
        with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
            for i, df in enumerate(dfs):
                sheet_name = f"Table_{i+1}"
                df.to_excel(writer, sheet_name=sheet_name, index=False)

        print(f"✅ Αποθηκεύτηκε ως: {output_excel}")

    except Exception as e:
        print(f"❌ Σφάλμα κατά τη μετατροπή: {e}")


if __name__ == "__main__":
    pdf_to_excel()
