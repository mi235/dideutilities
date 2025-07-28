import os
import xml.etree.ElementTree as ET
from openpyxl import Workbook

NAMESPACE = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
ns = {'w': NAMESPACE}

def extract_first_two_ypopsin_from_xml(xml_path):
    try:
        with open(xml_path, 'r', encoding='utf-8') as f:
            content = f.read()

        # Αν δεν έχει δηλωμένο namespace, το προσθέτουμε
        if 'xmlns:w=' not in content:
            content = content.replace('<w:document', f'<w:document xmlns:w="{NAMESPACE}"', 1)

        root = ET.fromstring(content)

        paragraphs = []
        for para in root.findall('.//w:p', ns):
            texts = [node.text for node in para.findall('.//w:t', ns) if node.text]
            full_text = ''.join(texts).strip()
            if full_text:
                paragraphs.append(full_text)

        if "Έχοντας υπόψη:" in paragraphs:
            idx = paragraphs.index("Έχοντας υπόψη:") + 1
            return paragraphs[idx:idx + 2]
        else:
            return ["", ""]
    except Exception as e:
        return [f"Σφάλμα: {e}", ""]

# 📂 Τροποποίησε με το φάκελο που έχεις τα XML
folder_path = "E:\ΧΜΛΣ"

# 📄 Δημιουργία Excel
wb = Workbook()
ws = wb.active
ws.title = "Λαμβάνοντας Υπόψη"
ws.append(["Όνομα αρχείου", "Υπόψη 1", "Υπόψη 2"])

for filename in os.listdir(folder_path):
    if filename.endswith(".xml"):
        full_path = os.path.join(folder_path, filename)
        yp1, yp2 = extract_first_two_ypopsin_from_xml(full_path)
        ws.append([filename, yp1, yp2])

wb.save("E:\ypopsin_from_xml.xlsx")
print("✅ Έτοιμο το Excel: ypopsin_from_xml.xlsx")
