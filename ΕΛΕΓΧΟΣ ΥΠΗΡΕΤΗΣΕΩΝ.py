import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from datetime import timedelta

# Επιλογή αρχείου με GUI
def choose_file(title):
    root = tk.Tk()
    root.withdraw()
    return filedialog.askopenfilename(
        title=title,
        filetypes=[("Excel αρχεία", "*.xlsx *.xls")]
    )

# Επιλογή αρχείου εισόδου
infile = choose_file("Επιλέξτε το αρχείο Excel για έλεγχο")
if not infile:
    messagebox.showerror("Σφάλμα", "Δεν επιλέχθηκε αρχείο.")
    exit()

# Διαβάζουμε τα δεδομένα
try:
    df = pd.read_excel(infile)
except Exception as e:
    messagebox.showerror("Σφάλμα", f"Αποτυχία ανάγνωσης Excel αρχείου:\n{str(e)}")
    exit()

# Βεβαιωνόμαστε ότι υπάρχουν οι απαιτούμενες στήλες
required_cols = {'ΑΦΜ', 'ΑΔΑ πρόσληψης', 'ΑΠΟ', 'ΕΩΣ', 'Σχολείο'}
if not required_cols.issubset(df.columns):
    messagebox.showerror("Σφάλμα", f"Το αρχείο πρέπει να περιέχει τις στήλες: {', '.join(required_cols)}")
    exit()

# Μετατροπή ημερομηνιών
df['ΑΠΟ'] = pd.to_datetime(df['ΑΠΟ'], errors='coerce')
df['ΕΩΣ'] = pd.to_datetime(df['ΕΩΣ'], errors='coerce')

# Φιλτράρουμε ΑΦΜ που εμφανίζονται πάνω από μία φορά
afm_counts = df['ΑΦΜ'].value_counts()
multi_afm = afm_counts[afm_counts > 1].index
df_multi = df[df['ΑΦΜ'].isin(multi_afm)]

results = []

for afm in df_multi['ΑΦΜ'].unique():
    df_afm = df_multi[df_multi['ΑΦΜ'] == afm]
    unique_adas = df_afm['ΑΔΑ πρόσληψης'].unique()

    for ada in unique_adas:
        df_ada = df_afm[df_afm['ΑΔΑ πρόσληψης'] == ada]

        unique_sxoleia = df_ada['Σχολείο'].nunique()

        # Έλεγχος για συνεχή διαστήματα
        df_ada_sorted = df_ada.sort_values(by='ΑΠΟ').reset_index(drop=True)
        is_continuous = True

        for i in range(1, len(df_ada_sorted)):
            prev_end = df_ada_sorted.loc[i - 1, 'ΕΩΣ']
            curr_start = df_ada_sorted.loc[i, 'ΑΠΟ']
            if pd.isna(prev_end) or pd.isna(curr_start):
                is_continuous = False
                break
            if curr_start != prev_end + timedelta(days=1):
                is_continuous = False
                break

        results.append({
            'ΑΦΜ': afm,
            'ΑΔΑ πρόσληψης': ada,
            'ΠΛΗΘΟΣ_ΣΧΟΛΕΙΩΝ': unique_sxoleia,
            'ΔΙΑΣΤΗΜΑ_ΣΥΝΕΧΟΜΕΝΟ': 'ΝΑΙ' if is_continuous else 'ΟΧΙ'
        })

# Εξαγωγή αποτελέσματος
res_df = pd.DataFrame(results)
outfile = "e:/elegxos_afm_ada_sxoleia.xlsx"
res_df.to_excel(outfile, index=False)

# Ενημέρωση χρήστη
messagebox.showinfo("Ολοκληρώθηκε", f"Ο έλεγχος ολοκληρώθηκε.\nΑποθηκεύτηκε στο:\n{outfile}")
