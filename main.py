# -*- coding: utf-8 -*-
import pandas as pd
import glob
import os

folder_path = r"E:\OneDrive\dded\2024-25\anaplir\opsd\erc"
pattern = os.path.join(folder_path, "*.xlsx")
combined_data = []
target_sheet = "Βασικοί Τίτλοι Σπουδών"

for file_path in glob.glob(pattern):
    try:
        # 👇 Use engine='openpyxl' in both ExcelFile and read_excel
        xl = pd.ExcelFile(file_path, engine='openpyxl')
        if target_sheet in xl.sheet_names:
            df = pd.read_excel(file_path, sheet_name=target_sheet, engine='openpyxl')

            df['srcc'] = os.path.basename(file_path)
            combined_data.append(df)
            print(u"read:", os.path.basename(file_path))
        else:
            print(u"⚠️ Το φύλλο '{}' δεν υπάρχει στο: {}".format(target_sheet, os.path.basename(file_path)))
    except Exception as e:
        print(u"❌ Σφάλμα στο", os.path.basename(file_path), ":", str(e))

if combined_data:
    final_df = pd.concat(combined_data, ignore_index=True)
    final_df.to_excel("e:\\all.xlsx", index=False, engine='openpyxl')  # 👈 And here for writing
    print(u"\n🎉 ok  all.xlsx!")
else:
    print(u"not found.")