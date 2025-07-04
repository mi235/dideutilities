import xml.etree.ElementTree as ET
import pandas as pd

df = pd.read_excel('E:\OIKON.xlsx')
afm_to_new_apodoxes = dict(zip(df['afm'].astype(str), df['f_apodoxes'].astype(str)))

# Φορτώνουμε το XML που περιέχει πολλές <AnaggeliaE7>
tree = ET.parse('E:\input.xml')
root = tree.getroot()


# Για κάθε <AnaggeliaE7>, ελέγχεις αν το ΑΦΜ είναι στη λίστα
for anaggelia in root.findall('AnaggeliaE7'):
    afm_elem = anaggelia.find('f_afm')
    if afm_elem is not None:
        afm = afm_elem.text.strip()
        if afm in afm_to_new_apodoxes:
            apodoxes_elem = anaggelia.find('f_apodoxes')
            if apodoxes_elem is not None:
                apodoxes_elem.text = afm_to_new_apodoxes[afm]

# Αποθήκευση σε νέο XML
tree.write('e:\output.xml', encoding='utf-8', xml_declaration=True)
