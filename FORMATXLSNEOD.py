import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, messagebox

# ----- Συνάρτηση για καθάρισμα -----
def clean_file(input_file, output_file):
    try:
        df = pd.read_excel(input_file, engine="odf")
    except:
        df = pd.read_excel(input_file)

    # 1. Σβήνουμε πλήρως κενές σειρές
    df = df.dropna(how="all")

    # 2. Forward fill στην πρώτη στήλη
    df.iloc[:, 0] = df.iloc[:, 0].ffill()

    # 3. Διαγραφή σειρών με 'Επεξεργασία' ή 'Διαγραφή' στις στήλες O και W
    col_o = df.columns[14]  # Στήλη O
    col_w = df.columns[22]  # Στήλη W
    df = df[~df[col_o].astype(str).str.strip().isin(["Επεξεργασία", "Διαγραφή"])]
    df = df[~df[col_w].astype(str).str.strip().isin(["Επεξεργασία", "Διαγραφή"])]

    # 4. Συγχώνευση τιμών στη W που ξεκινούν με κόμμα
    col_w_index = df.columns.get_loc(col_w)
    rows_to_drop = []
    for i in range(1, len(df)):
        current_val = str(df.iloc[i, col_w_index]).strip()
        if current_val.startswith(','):
            prev_val = str(df.iloc[i - 1, col_w_index]).strip()
            df.iloc[i - 1, col_w_index] = prev_val + current_val
            rows_to_drop.append(i)
    df = df.drop(index=df.index[rows_to_drop]).reset_index(drop=True)

    # 5. Μορφοποίηση αριθμητικών
    def format_numeric(val, total_digits):
        try:
            return str(int(float(val))).zfill(total_digits)
        except:
            return str(val).strip() if pd.notnull(val) else ""

    df[df.columns[1]] = df[df.columns[1]].apply(lambda x: format_numeric(x, 9))   # Στήλη 2
    df[df.columns[11]] = df[df.columns[11]].apply(lambda x: format_numeric(x, 11)) # Στήλη L
    df[df.columns[12]] = df[df.columns[12]].apply(lambda x: format_numeric(x, 9))  # Στήλη M

    # 6. Μορφοποίηση ημερομηνιών (στήλη H)
    def format_date(val):
        if pd.notnull(val) and str(val).strip() != "":
            val_int = int(float(val))
            val_str = str(val_int).zfill(8)
            return f"{val_str[:2]}/{val_str[2:4]}/{val_str[4:]}"
        else:
            return ""
    df[df.columns[7]] = df[df.columns[7]].apply(format_date)

    # 7. Αποθήκευση
    df.to_excel(output_file, index=False)


# ----- GUI -----
def main():
    root = tk.Tk()
    root.withdraw()

    messagebox.showinfo("Υπενθύμιση",
                        "⚠️ Το αρχείο ODS δεν πρέπει να έχει συγχωνευμένα κελιά.\n"
                        "Παρακαλώ βεβαιωθείτε πριν συνεχίσετε.")

    input_file = filedialog.askopenfilename(
        title="Επιλέξτε αρχείο εισόδου (ODS/Excel)",
        filetypes=[("ODS/Excel files", "*.ods *.xlsx *.xls")]
    )
    if not input_file:
        return

    output_file = filedialog.asksaveasfilename(
        title="Αποθήκευση ως...",
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")]
    )
    if not output_file:
        return

    try:
        clean_file(input_file, output_file)
        messagebox.showinfo("Ολοκληρώθηκε", f"✅ Το καθαρισμένο αρχείο αποθηκεύτηκε:\n{output_file}")
    except Exception as e:
        messagebox.showerror("Σφάλμα", f"❌ Σφάλμα: {e}")


if __name__ == "__main__":
    main()
