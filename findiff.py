import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

def choose_file(title):
    root = tk.Tk()
    root.withdraw()
    return filedialog.askopenfilename(
        title=title,
        filetypes=[("Excel αρχεία", "*.xlsx *.xls")]
    )

def save_file(title):
    root = tk.Tk()
    root.withdraw()
    return filedialog.asksaveasfilename(
        title=title,
        defaultextension=".xlsx",
        filetypes=[("Excel αρχεία", "*.xlsx")]
    )

# Επιλογή αρχείων
file1 = choose_file("Επιλογή 1ου Excel")
file2 = choose_file("Επιλογή 2ου Excel")

if not file1 or not file2:
    messagebox.showwarning("Ακύρωση", "Δεν επιλέχθηκαν αρχεία.")
    exit()

# Φόρτωση πινάκων
df1 = pd.read_excel(file1)
df2 = pd.read_excel(file2)

# Ορισμός κλειδιών για ευθυγράμμιση
key_cols = ["ΑΦΜ", "ΑΔΑ πρόσληψης", "Πρόσμετρηση Υπηρεσίας Από","Σχολείο"]

# Πεδίο προς σύγκριση
compare_cols = [
     "Ώρες/Εβδομάδα", "ΑΠΟ",
    "ΕΩΣ", "Παρατηρήσεις"
]

# Φιλτράρω πεδία που υπάρχουν και στα δύο αρχεία (για ασφάλεια)
compare_cols = [c for c in compare_cols if c in df1.columns and c in df2.columns]

# Βάζω index με τα κλειδιά
df1_idx = df1.set_index(key_cols)
df2_idx = df2.set_index(key_cols)

# Κάνω join ώστε να έχω και τα δύο αρχεία μαζί για σύγκριση
combined = df1_idx.join(df2_idx, how="outer", lsuffix="_1", rsuffix="_2")

diff_records = []

# Ελέγχω για διαφορές σε κάθε πεδίο σύγκρισης
for col in compare_cols:
    col1 = f"{col}_1"
    col2 = f"{col}_2"
    # Μάσκα διαφορών με προσοχή στις τιμές NaN
    mask = ~combined[col1].eq(combined[col2])
    diff_rows = combined[mask][[col1, col2]]
    for idx, row in diff_rows.iterrows():
        diff_records.append({
            "ΑΦΜ": idx[0],
            "ΑΔΑ πρόσληψης": idx[1],
            "Στήλη": col,
            "Τιμή_Αρχείο_1": row[col1],
            "Τιμή_Αρχείο_2": row[col2]
        })

# Δημιουργία DataFrame με τις διαφορές
diff_df = pd.DataFrame(diff_records)

if diff_df.empty:
    messagebox.showinfo("Σύγκριση", "Δεν βρέθηκαν διαφορές στα επιλεγμένα πεδία.")
else:
    out_file = save_file("Αποθήκευση αποτελέσματος σύγκρισης")
    if out_file:
        diff_df.to_excel(out_file, index=False)
        messagebox.showinfo("Τέλος", f"Βρέθηκαν {len(diff_df)} διαφορές.\nΑποθηκεύτηκαν στο:\n{out_file}")
    else:
        messagebox.showinfo("Ακύρωση", "Δεν επιλέχθηκε αρχείο αποθήκευσης.\nΟι διαφορές είναι:")
        print(diff_df)
