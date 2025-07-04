import tkinter as tk
from tkinter import filedialog, messagebox
import xml.etree.ElementTree as ET
import pandas as pd
import os

def select_xml():
    xml_path.set(filedialog.askopenfilename(filetypes=[("XML Files", "*.xml")]))

def select_excel():
    excel_path.set(filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")]))

def run_script():
    try:
        # Διαβάζουμε τα paths
        xml_file = xml_path.get()
        excel_file = excel_path.get()

        if not xml_file or not excel_file:
            messagebox.showerror("Σφάλμα", "Παρακαλώ επιλέξτε και τα δύο αρχεία.")
            return

        # Φορτώνουμε XML
        tree = ET.parse(xml_file)
        root = tree.getroot()

        # Διαβάζουμε Excel
        df = pd.read_excel(excel_file, dtype=str)
        df.fillna('', inplace=True)

        # Φτιάχνουμε λεξικό: afm -> apodoxes
        afm_to_apodoxes = dict(zip(df['AFMTEXT'], df['fapodoxes']))

        # Επεξεργασία κάθε <AnaggeliaE7>
        for anaggelia in root.findall(".//AnaggeliaE7"):
            afm_elem = anaggelia.find("f_afm")
            apodoxes_elem = anaggelia.find("f_apodoxes")


            if afm_elem is not None:
                afm = afm_elem.text.strip()

                if afm in afm_to_apodoxes and apodoxes_elem is not None:
                    apodoxes_elem.text = afm_to_apodoxes[afm]


        # Αποθήκευση νέου XML
        new_file = os.path.splitext(xml_file)[0] + "_new.xml"
        tree.write(new_file, encoding='utf-8', xml_declaration=True)
        messagebox.showinfo("Επιτυχία", f"Το νέο XML αποθηκεύτηκε:\n{new_file}")

    except Exception as e:
        messagebox.showerror("Σφάλμα", f"Παρουσιάστηκε πρόβλημα:\n{str(e)}")

# GUI
root = tk.Tk()
root.title("Επεξεργασία XML από Excel")
root.geometry("500x250")

xml_path = tk.StringVar()
excel_path = tk.StringVar()

tk.Label(root, text="Επιλογή XML").pack()
tk.Entry(root, textvariable=xml_path, width=60).pack()
tk.Button(root, text="Browse XML", command=select_xml).pack()

tk.Label(root, text="Επιλογή Excel").pack()
tk.Entry(root, textvariable=excel_path, width=60).pack()
tk.Button(root, text="Browse Excel", command=select_excel).pack()

tk.Button(root, text="Εκτέλεση", command=run_script, bg="lightgreen").pack(pady=20)

root.mainloop()


