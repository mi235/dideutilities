# -*- coding: utf-8 -*-
import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog

def merge_excel_sheets_select_fields(input_path, output_path):
    if not os.path.isfile(input_path):
        print(f"\n❌ Το αρχείο δεν βρέθηκε: {input_path}")
        return

    try:
        xl = pd.ExcelFile(input_path, engine='openpyxl')
    except Exception as e:
        print(f"\n❌ Δεν ήταν δυνατή η ανάγνωση του αρχείου: {e}")
        return

    all_headers = set()
    for sheet in xl.sheet_names:
        try:
            df = xl.parse(sheet, header=5)
            all_headers.update(df.columns.tolist())
        except Exception as e:
            print(f"⚠️ Παράβλεψη φύλλου '{sheet}': {e}")

    if not all_headers:
        print("\n⚠️ Δεν βρέθηκαν πεδία.")
        return

    selected_fields = show_field_selector(sorted(all_headers))
    if not selected_fields:
        print("❌ Δεν επιλέχθηκαν πεδία.")
        return

    combined = []
    for sheet in xl.sheet_names:
        try:
            df = xl.parse(sheet, header=5)
            df = df[selected_fields]
            df.insert(0, 'source_sheet', sheet)
            combined.append(df)
            print(f"✅ Προστέθηκε το φύλλο: {sheet}")
        except Exception as e:
            print(f"❌ Σφάλμα στο φύλλο '{sheet}': {e}")

    if combined:
        final_df = pd.concat(combined, ignore_index=True)
        try:
            final_df.to_excel(output_path, index=False, engine='openpyxl')
            print(f"\n🎉 Αποθηκεύτηκε στο: {output_path}")
        except Exception as e:
            print(f"\n❌ Σφάλμα κατά την αποθήκευση: {e}")
    else:
        print("\n⚠️ Δεν βρέθηκαν δεδομένα προς συγχώνευση.")

def show_field_selector(options):
    selected = []

    def submit():
        nonlocal selected
        selected = [name for name, var in vars if var.get()]
        root.destroy()  # **Κλείνει το παράθυρο και βγαίνει από mainloop**

    root = tk.Tk()
    root.title("📝 Επιλογή Πεδίων")
    root.geometry("400x500")

    container = tk.Frame(root)
    container.pack(fill="both", expand=True)

    canvas = tk.Canvas(container)
    scrollbar = tk.Scrollbar(container, orient="vertical", command=canvas.yview)
    scrollable_frame = tk.Frame(canvas)

    scrollable_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
    )

    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    vars = []
    for opt in options:
        var = tk.BooleanVar()
        chk = tk.Checkbutton(scrollable_frame, text=opt, variable=var, anchor='w', justify='left')
        chk.pack(fill='x', padx=10, anchor='w')
        vars.append((opt, var))

    submit_btn = tk.Button(root, text="✅ Συνέχεια", command=submit)
    submit_btn.pack(pady=10)

    root.mainloop()
    return selected


if __name__ == "__main__":
    print("🔄 Συγχώνευση όλων των φύλλων Excel με επιλογή πεδίων...")

    root = tk.Tk()
    root.withdraw()

    input_file = filedialog.askopenfilename(
        title="📥 Επιλέξτε το αρχείο Excel εισόδου",
        filetypes=[("Excel αρχεία", "*.xlsx")]
    )

    if not input_file:
        print("❌ Δεν επιλέχθηκε αρχείο εισόδου.")
        exit()

    output_file = filedialog.asksaveasfilename(
        title="📤 Επιλέξτε που θα αποθηκευτεί το αρχείο",
        defaultextension=".xlsx",
        filetypes=[("Excel αρχεία", "*.xlsx")]
    )

    if not output_file:
        print("❌ Δεν επιλέχθηκε αρχείο εξόδου.")
        exit()

    merge_excel_sheets_select_fields(input_file, output_file)
