import pandas as pd
from tkinter import Tk, filedialog
import os

def katharise_ods(input_file, output_file):
    # Διαβάζουμε το .ods με engine "odf"
    df = pd.read_excel(input_file, engine="odf")

    # Κρατάμε μόνο τις γραμμές που:
    # - έχουν τιμή στη στήλη Α ή
    # - είναι οι τελευταίες από τα "μπλοκ" (η επόμενη γραμμή έχει NaN στη στήλη Α)
    mask = df.iloc[:, 0].notna() | df.iloc[:, 0].shift(-1).isna()
    df_katharo = df[mask]

    # Αποθήκευση σε .ods ή .xlsx
    ext = os.path.splitext(output_file)[1].lower()
    if ext == ".ods":
        df_katharo.to_excel(output_file, engine="odf", index=False)
    else:
        # default: .xlsx
        df_katharo.to_excel(output_file, engine="openpyxl", index=False)

    print(f"Το καθαρισμένο αρχείο αποθηκεύτηκε ως: {output_file}")

if __name__ == "__main__":
    root = Tk()
    root.withdraw()

    # Επιλογή αρχείου εισόδου (.ods)
    print("Επίλεξε το αρχείο ODS που θέλεις να καθαρίσεις...")
    input_file = filedialog.askopenfilename(
        title="Επίλεξε αρχείο ODS",
        filetypes=[("ODS files", "*.ods")]
    )

    if not input_file:
        print("Δεν επιλέχθηκε αρχείο. Έξοδος.")
        exit()

    # Επιλογή αρχείου εξόδου (.ods ή .xlsx)
    print("Δώσε όνομα για το νέο αρχείο...")
    output_file = filedialog.asksaveasfilename(
        title="Αποθήκευση καθαρισμένου αρχείου",
        defaultextension=".ods",
        filetypes=[("ODS files", "*.ods"), ("Excel files", "*.xlsx")]
    )

    if not output_file:
        print("Δεν επιλέχθηκε αρχείο εξόδου. Έξοδος.")
        exit()

    # Εκτέλεση καθαρισμού
    katharise_ods(input_file, output_file)
