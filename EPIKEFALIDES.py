import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

def epilogi_eisodou():
    root = tk.Tk()
    root.withdraw()
    filepath = filedialog.askopenfilename(
        title="Επιλογή αρχείου Excel",
        filetypes=[("Excel αρχεία", "*.xlsx *.xls")]
    )
    return filepath

def epilogi_exodou():
    root = tk.Tk()
    root.withdraw()
    filepath = filedialog.asksaveasfilename(
        title="Αποθήκευση νέου αρχείου Excel",
        defaultextension=".xlsx",
        filetypes=[("Excel αρχεία", "*.xlsx")]
    )
    return filepath


def eisagogi_grammon_klados(df):
    # Ταξινόμηση με 3 κριτήρια: KLADOS, ΕΙΔΙΚΗ ΚΑΤΗΓΟΡΙΑ, MORIA
    df = df.sort_values(["KLADOS", "ΕΙΔΙΚΗ ΚΑΤΗΓΟΡΙΑ", "MORIA"], ascending=[True, True, False])

    columns = df.columns.tolist()
    rows = []

    for klados, group in df.groupby("KLADOS"):
        empty_row = [""] * len(columns)
        empty_row[0] = f"KLADOS = {klados}"
        rows.append(empty_row)
        rows.extend(group.values.tolist())

    return pd.DataFrame(rows, columns=columns)


def main():
    # Επιλογή εισόδου
    excel_file = epilogi_eisodou()
    if not excel_file:
        messagebox.showwarning("Προσοχή", "Δεν επιλέχθηκε αρχείο εισόδου.")
        return

    # Ανάγνωση δεδομένων
    try:
        df = pd.read_excel(excel_file)
    except Exception as e:
        messagebox.showerror("Σφάλμα", f"Σφάλμα κατά την ανάγνωση του αρχείου:\n{e}")
        return

    # Επεξεργασία
    new_df = eisagogi_grammon_klados(df)

    # Επιλογή εξόδου
    output_file = epilogi_exodou()
    if not output_file:
        messagebox.showinfo("Ακύρωση", "Δεν επιλέχθηκε αρχείο εξόδου.")
        return

    # Αποθήκευση
    try:
        new_df.to_excel(output_file, index=False)
        messagebox.showinfo("Ολοκληρώθηκε", f"Το νέο αρχείο αποθηκεύτηκε:\n{output_file}")
    except Exception as e:
        messagebox.showerror("Σφάλμα", f"Σφάλμα κατά την αποθήκευση:\n{e}")

if __name__ == "__main__":
    main()
