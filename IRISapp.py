import pytesseract
from PIL import Image, ImageEnhance, ImageFilter
import re
import pandas as pd
from openpyxl.styles import Font

# Pfad zur Bilddatei
image_path = "20250311_T0B1_1_spec.bmp"

# Bild laden
img = Image.open(image_path)

# 1. Umwandeln in Graustufen
gray = img.convert('L')

# 2. Kontrast verstärken
enhancer = ImageEnhance.Contrast(gray)
gray = enhancer.enhance(2.0)  # Kontrastfaktor anpassbar

# 3. Bild schärfen
gray = gray.filter(ImageFilter.SHARPEN)

# 4. Rauschreduktion mittels Medianfilter
gray = gray.filter(ImageFilter.MedianFilter(size=3))

# 5. Binarisierung: Schwellwert anwenden, um den Text hervorzuheben
threshold = 128
bw = gray.point(lambda x: 255 if x > threshold else 0, mode='1')

# 6. Optional: Bild vergrößern, falls es relativ klein ist (kann zu besseren OCR-Ergebnissen führen)
width, height = bw.size
if width < 1000:
    bw = bw.resize((width * 2, height * 2), Image.ANTIALIAS)

# OCR: Text extrahieren (mit deutscher Sprachunterstützung)
extracted_text = pytesseract.image_to_string(bw, lang="deu")
print("Extrahierter Text:")
print(extracted_text)

# Regulärer Ausdruck zum Erfassen der LAB-Werte und der Prozentangabe
# Annahme: Es wird immer ein Muster wie "L* <Wert>  a* <Wert>  b* <Wert>  <Prozent>" verwendet
pattern = r"L\*[\s:]*([\d.,]+).*?a\*[\s:]*([\d.,-]+).*?b\*[\s:]*([\d.,-]+).*?(\d+[%])"
matches = re.findall(pattern, extracted_text, re.DOTALL)

data = []
for idx, (L_val, a_val, b_val, percent) in enumerate(matches, start=1):
    # Fehlerprüfung: Wenn der extrahierte Farbcode (hier als L_val interpretiert)
    # nur drei Ziffern enthält (ohne Punkt oder Komma), gilt dies als OCR-Fehler.
    numeric_str = L_val.replace(',', '').replace('.', '')
    error_flag = False
    if len(numeric_str) == 3:
        error_flag = True

    # Hinweis: Bei Fehlern wird "MANUEL" eingetragen.
    hinweis = "MANUEL" if error_flag else ""
    
    data.append({
        "Histogramm Balken": idx,
        "L*": L_val,
        "a*": a_val,
        "b*": b_val,
        "Prozent": percent,
        "Hinweis": hinweis
    })

# DataFrame erstellen
df = pd.DataFrame(data)
print(df)

# Ergebnisse in eine Excel-Datei schreiben und den Fehlerhinweis (MANUEL) in roter Schrift formatieren.
excel_path = "ausgelesene_farbcodes.xlsx"
with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
    df.to_excel(writer, index=False, sheet_name="Farbcodes")
    workbook = writer.book
    worksheet = writer.sheets["Farbcodes"]

    # Erstelle einen roten Font
    red_font = Font(color="FF0000")
    # Annahme: Spalte 6 ("Hinweis") enthält den Fehlerhinweis
    for row in range(2, len(df) + 2):  # Zeile 1 ist die Header-Zeile
        cell = worksheet.cell(row=row, column=6)
        if cell.value == "MANUEL":
            cell.font = red_font

print(f"Ergebnisse wurden in '{excel_path}' gespeichert.")
