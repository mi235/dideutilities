import pandas as pd
import os

# ----- Ρυθμίσεις -----
input_file = r"E:\OneDrive\dded\2025-26\ΔΙΟΡΙΣΜΟΙ25\neodiristoi-2023-view.ods"
output_file = os.path.splitext(input_file)[0] + "-clean.xlsx"

# ----- 1. Ανάγνωση αρχείου -----
try:
    df = pd.read_excel(input_file, engine="odf")
    print("📂 Το αρχείο διαβάστηκε ως ODS (engine='odf').")
except:
    df = pd.read_excel(input_file)
    print("📂 Το αρχείο διαβάστηκε ως Excel.")

# ----- 2. Σβήνουμε πλήρως κενές σειρές -----
df = df.dropna(how="all")

# ----- 3. Forward fill στην πρώτη στήλη -----
df.iloc[:, 0] = df.iloc[:, 0].ffill()

# ----- 4. Διαγραφή σειρών με 'Επεξεργασία' ή 'Διαγραφή' στις στήλες O και W -----
col_o = df.columns[14]  # Στήλη O (15η)
col_w = df.columns[22]  # Στήλη W (23η)
df = df[~df[col_o].astype(str).str.strip().isin(["Επεξεργασία", "Διαγραφή"])]
df = df[~df[col_w].astype(str).str.strip().isin(["Επεξεργασία", "Διαγραφή"])]

# ----- 5. Concatenate στη στήλη W για γραμμές που ξεκινούν με κόμμα -----
col_w_index = df.columns.get_loc(col_w)  # αριθμητική θέση της W
rows_to_drop = []

for i in range(1, len(df)):
    current_val = str(df.iloc[i, col_w_index]).strip()

    if current_val.startswith(','):
        prev_val = str(df.iloc[i - 1, col_w_index]).strip()
        df.iloc[i - 1, col_w_index] = prev_val + current_val
        rows_to_drop.append(i)

df = df.drop(index=df.index[rows_to_drop]).reset_index(drop=True)


# ----- 6. Συνάρτηση μορφοποίησης αριθμητικών -----
def format_numeric(val, total_digits):
    try:
        return str(int(float(val))).zfill(total_digits)
    except:
        return str(val).strip() if pd.notnull(val) else ""


# Στήλη 2 → 9 ψηφία
col_2 = df.columns[1]
df[col_2] = df[col_2].apply(lambda x: format_numeric(x, 9))

# Στήλη L → 11 ψηφία
col_l = df.columns[11]
df[col_l] = df[col_l].apply(lambda x: format_numeric(x, 11))

# Στήλη M → 9 ψηφία
col_m = df.columns[12]
df[col_m] = df[col_m].apply(lambda x: format_numeric(x, 9))

# ----- 7. Μορφοποίηση στήλης H σε DD/MM/YYYY μόνο για μη κενά κελιά -----
col_h = df.columns[7]  # Στήλη H


def format_date(val):
    if pd.notnull(val) and str(val).strip() != "":
        val_int = int(float(val))  # αφαιρούμε το .0
        val_str = str(val_int).zfill(8)  # εξασφαλίζουμε 8 ψηφία
        return f"{val_str[:2]}/{val_str[2:4]}/{val_str[4:]}"
    else:
        return ""  # αφήνουμε κενό κελί ως έχει


df[col_h] = df[col_h].apply(format_date)

# ----- 8. Αποθήκευση -----
df.to_excel(output_file, index=False)
print(f"✅ Το καθαρισμένο και μορφοποιημένο αρχείο αποθηκεύτηκε ως:\n{output_file}")
