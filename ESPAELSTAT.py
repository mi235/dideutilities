# -*- coding: utf-8 -*-
import pandas as pd
import os
from tkinter import Tk, filedialog, simpledialog
import sys

# Αν έχεις Windows και θες σωστή απεικόνιση στην κονσόλα:
sys.stdout.reconfigure(encoding='utf-8')

def merge_all_excels_select_fields(excel_paths, output_path, selected_fields, header_row):
    combined_data = []

    for file_path in excel_paths:
        try:
            xl = pd.ExcelFile(file_path, engine='openpyxl')
        except Exception as e:
            print(f"❌ Σφάλμα στο αρχείο '{file_path}': {e}")
            continue

        for sheet_name in xl.sheet_names:
            try:
                # Έλεγχος αν υπάρχουν τουλάχιστον 4 γραμμές
                preview_df = xl.parse(sheet_name, header=None)
                if len(preview_df) < (header_row + 1):
                    print(f"⚠️ Το φύλλο '{sheet_name}' στο '{os.path.basename(file_path)}' έχει λιγότερες από {header_row+1} γραμμές. Αγνοείται.")
                    continue

                # Διαβάζουμε με header από κατάλληλη γραμμή
                df = xl.parse(sheet_name, header=header_row)
                if not df.empty:
                    df.insert(0, 'source_sheet', sheet_name)
                    df.insert(0, 'source_file', os.path.basename(file_path))

                    # Κράτα μόνο τις επιλεγμένες στήλες (όσες υπάρχουν)
                    available_fields = [col for col in selected_fields if col in df.columns]
                    meta_fields = ['source_file', 'source_sheet']
                    df = df[meta_fields + available_fields]

                    combined_data.append(df)
                    print(f"✅ Προστέθηκε το φύλλο: {sheet_name} από {os.path.basename(file_path)}")
                else:
                    print(f"⚠️ Άδειο φύλλο: {sheet_name} στο {os.path.basename(file_path)}")
            except Exception as e:
                print(f"❌ Σφάλμα στο φύλλο '{sheet_name}': {e}")

    if combined_data:
        final_df = pd.concat(combined_data, ignore_index=True)
        try:
            final_df.to_excel(output_path, index=False, engine='openpyxl')
            print(f"\n🎉 Ολοκληρώθηκε η συγχώνευση στο: {output_path}")
        except Exception as e:
            print(f"\n❌ Σφάλμα κατά την αποθήκευση: {e}")
    else:
        print("\n⚠️ Δεν βρέθηκαν δεδομένα προς συγχώνευση.")

if __name__ == "__main__":
    print("🔄 Συγχώνευση όλων των φύλλων από πολλαπλά Excel")

    root = Tk()
    root.withdraw()

    # 🧠 Ρώτα αν είναι ΕΣΠΑ ή ΔΙΑΣ
    source_type = simpledialog.askstring(
        "Τύπος Δεδομένων",
        "Είναι τα αρχεία ΕΣΠΑ ή ΔΙΑΣ;\nΓράψε 'εσπα' ή 'διας'.")

    if not source_type:
        print("❌ Δεν επιλέχθηκε τύπος δεδομένων.")
        exit()

    source_type = source_type.strip().lower()
    if source_type == 'διας':
        header_row = 3
    elif source_type == 'εσπα':
        header_row = 4
    else:
        print("❌ Μη αποδεκτός τύπος. Πρέπει να είναι 'εσπα' ή 'διας'.")
        exit()

    # 📂 Επιλογή Excel αρχείων
    input_files = filedialog.askopenfilenames(
        title="📥 Επιλέξτε ένα ή περισσότερα αρχεία Excel",
        filetypes=[("Excel files", "*.xlsx")])

    if not input_files:
        print("❌ Δεν επιλέχθηκαν αρχεία.")
        exit()

    # ➕ Πρώτο πέρασμα για εύρεση πεδίων
    all_columns = set()
    for file_path in input_files:
        try:
            xl = pd.ExcelFile(file_path, engine='openpyxl')
            for sheet in xl.sheet_names:
                try:
                    preview_df = xl.parse(sheet, header=None)
                    if len(preview_df) >= (header_row + 1):
                        df = xl.parse(sheet, header=header_row)
                        valid_columns = [col for col in df.columns if isinstance(col, str) and col.strip()]
                        all_columns.update(valid_columns)
                except:
                    continue
        except Exception as e:
            print(f"⚠️ Πρόβλημα με το {file_path}: {e}")

    all_columns = sorted([col for col in all_columns if isinstance(col, str)])

    # ✅ Εμφάνιση πεδίων
    print("\n📋 Διαθέσιμα πεδία:")
    for i, col in enumerate(all_columns, 1):
        print(f"{i}. {col}")

    selected_indices = simpledialog.askstring(
        "Επιλογή Πεδίων",
        "Δώσε τους αριθμούς των πεδίων που θέλεις (χωρισμένα με κόμμα):\nπ.χ. 1,3,5")

    if not selected_indices:
        print("❌ Δεν επιλέχθηκαν πεδία.")
        exit()

    try:
        selected_fields = [all_columns[int(i.strip()) - 1] for i in selected_indices.split(',')]
    except:
        print("❌ Μη έγκυρη επιλογή.")
        exit()

    print("\n🔍 Επιλεγμένα πεδία:")
    for field in selected_fields:
        print(f"• {field}")

    # 📤 Επιλογή αρχείου εξόδου
    output_file = filedialog.asksaveasfilename(
        title="📤 Επιλέξτε πού θα αποθηκευτεί το αρχείο",
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")])

    if not output_file:
        print("❌ Δεν επιλέχθηκε αρχείο αποθήκευσης.")
        exit()

    merge_all_excels_select_fields(input_files, output_file, selected_fields, header_row)
