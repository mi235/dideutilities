import os
import shutil
from tkinter import Tk, filedialog
from openpyxl import Workbook

# Διάλεξε φάκελο με τα έγγραφα
root = Tk()
root.withdraw()
source_dir = filedialog.askdirectory(title="Επίλεξε φάκελο με έγγραφα")

# Διάλεξε βασικό φάκελο με τους υποφακέλους
base_dir = filedialog.askdirectory(title="Επίλεξε βασικό φάκελο με υποφακέλους")
root.destroy()

# λίστα με φακέλους που πήραν αρχεία
folders_filled = set()

# Πέρασμα από όλα τα αρχεία του source_dir
for filename in os.listdir(source_dir):
    file_path = os.path.join(source_dir, filename)

    if os.path.isfile(file_path):
        # πάρε τα ΠΡΩΤΑ 9 ψηφία του ονόματος αρχείου (χωρίς επέκταση)
        name_without_ext, ext = os.path.splitext(filename)
        prefix = name_without_ext[:9]

        # ψάξε φάκελο που τελειώνει με αυτά τα 9 ψηφία
        for folder in os.listdir(base_dir):
            if folder.endswith(prefix):
                dest_path = os.path.join(base_dir, folder, filename)

                if not os.path.exists(dest_path):
                    shutil.copy2(file_path, dest_path)
                    print(f"Αντιγράφηκε: {filename} -> {folder}")
                else:
                    print(f"Υπάρχει ήδη: {filename} στον {folder}")

                # καταγράφουμε ότι ο φάκελος γέμισε
                folders_filled.add(folder)
                break

# Δημιουργία Excel report
if folders_filled:
    report_path = os.path.join(base_dir, "report.xlsx")
    wb = Workbook()
    ws = wb.active
    ws.title = "Φάκελοι"

    ws.append(["Φάκελοι που γέμισαν"])
    for folder in sorted(folders_filled):
        ws.append([folder])

    wb.save(report_path)
    print(f"\n✅ Δημιουργήθηκε report: {report_path}")
else:
    print("\n⚠ Δεν γέμισε κανένας φάκελος.")
