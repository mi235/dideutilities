import pandas as pd
from tkinter import filedialog, Tk

# Απόκρυψη του βασικού παραθύρου Tkinter
root = Tk()
root.withdraw()

# Επιλογή αρχείου εισόδου (.xlsx)
print("Επίλεξε το αρχείο Excel που θέλεις να καθαρίσεις...")
filename = filedialog.askopenfilename(
    title="Επίλεξε αρχείο Excel",
    filetypes=[("Excel files", "*.xlsx")]
)

if not filename:
    print("Δεν επιλέχθηκε αρχείο. Έξοδος.")
    exit()

# Διαβάζουμε το φύλλο
df = pd.read_excel(filename, sheet_name="clear")

# Επιλογή στηλών με βάση τον αριθμό τους (0 = πρώτη στήλη)
selected_columns = [2,3,4,5,6,9,15,18,19,22,27,28,36,43,44]
df_selected = df.iloc[:, selected_columns]

# Αποθήκευση σε νέο φύλλο
with pd.ExcelWriter(filename, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
    df_selected.to_excel(writer, sheet_name="Επιλεγμένες_Στήλες", index=False)

print("Το νέο φύλλο δημιουργήθηκε με επιτυχία ✅")
