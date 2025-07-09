import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

# -------------------------------------------------
#  Βοηθητικές συναρτήσεις GUI
# -------------------------------------------------
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

# -------------------------------------------------
#  Επιλογή αρχείων
# -------------------------------------------------
file1 = choose_file("Επιλογή 1ου Excel")
file2 = choose_file("Επιλογή 2ου Excel")
if not file1 or not file2:
    messagebox.showwarning("Ακύρωση", "Δεν επιλέχθηκαν αρχεία.")
    exit()

df1 = pd.read_excel(file1)
df2 = pd.read_excel(file2)

# -------------------------------------------------
#  Ορισμός κλειδιών και στηλών
# -------------------------------------------------
key_cols = ["ΑΦΜ", "ΑΔΑ πρόσληψης", "Σχολείο", "Ημ/νία Ανάληψης Υπηρεσίας"]

compare_cols = [
    'Επώνυμο', 'Όνομα', 'Πατρώνυμο', 'Κλάδος',
     'Πρόσμετρηση Υπηρεσίας Από',
    'Υπηρεσία Καταχώρισης', 'Ιδιότητα',
    'Αριθμός Απόφασης Πρόσληψης', 'Ημερομηνία Απόφασης Πρόσληψης',
     'Ώρες Μειωμένου', 'Ημ/νία Λήξης Υπηρεσίας',
    'Ώρες/Εβδομάδα', 'ΑΠΟ', 'ΕΩΣ', 'Πράξη', 'Παρατηρήσεις'
]

compare_cols = [c for c in compare_cols if c in df1.columns and c in df2.columns]

# -------------------------------------------------
#  Index στους πίνακες & έλεγχος διαφορών
# -------------------------------------------------
df1_idx = df1.set_index(key_cols)
df2_idx = df2.set_index(key_cols)

diff_records = []

# 1️⃣ Γραμμές που λείπουν εντελώς
only_in_1 = df1_idx.index.difference(df2_idx.index)
for idx in only_in_1:
    record = {key: idx[i] for i, key in enumerate(key_cols)}
    record.update({
        "Στήλη": "(ολόκληρη γραμμή)",
        "Τιμή_A": "υπάρχει", "Τιμή_B": "λείπει"
    })
    diff_records.append(record)

only_in_2 = df2_idx.index.difference(df1_idx.index)
for idx in only_in_2:
    record = {key: idx[i] for i, key in enumerate(key_cols)}
    record.update({
        "Στήλη": "(ολόκληρη γραμμή)",
        "Τιμή_A": "λείπει", "Τιμή_B": "υπάρχει"
    })
    diff_records.append(record)

# 2️⃣ Κελιά που διαφέρουν στις κοινές γραμμές
common_idx = df1_idx.index.intersection(df2_idx.index)

for col in compare_cols:
    s1 = df1_idx.loc[common_idx, col]
    s2 = df2_idx.loc[common_idx, col]
    mask = ~(s1.eq(s2) | (s1.isna() & s2.isna()))

    for idx in mask[mask].index:
        record = {key: idx[i] for i, key in enumerate(key_cols)}
        record.update({
            "Στήλη": col,
            "Τιμή_A": s1.loc[idx],
            "Τιμή_B": s2.loc[idx]
        })
        diff_records.append(record)

# -------------------------------------------------
#  Αποτέλεσμα
# -------------------------------------------------
diff_df = pd.DataFrame(diff_records)

if diff_df.empty:
    messagebox.showinfo("Σύγκριση", "Δεν βρέθηκαν διαφορές.")
else:
    out_file = save_file("Αποθήκευση αποτελέσματος")
    if out_file:
        diff_df.to_excel(out_file, index=False)
        messagebox.showinfo(
            "Τέλος",
            f"Βρέθηκαν {len(diff_df)} διαφορές.\nΑποθηκεύτηκαν στο:\n{out_file}"
        )
    else:
        print(diff_df)
