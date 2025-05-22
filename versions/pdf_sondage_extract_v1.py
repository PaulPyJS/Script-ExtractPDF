import pdfplumber
from tkinter import Tk, filedialog
import logging

logging.getLogger("pdfminer").setLevel(logging.ERROR)

def choose_pdf():
    root = Tk()
    root.withdraw()
    return filedialog.askopenfilename(
        title="Choisir un fichier PDF",
        filetypes=[("Fichiers PDF", "*.pdf")]
    )


def extract_values_near_keywords(pdf_path, keywords,
                                  tolerance_left=5, tolerance_right=30,
                                  min_dy_from_label=5):
    results = {}

    with pdfplumber.open(pdf_path) as pdf:
        page = pdf.pages[0]
        words = page.extract_words()

        for keyword in keywords:
            ref_word = next((w for w in words if keyword.lower() in w['text'].lower()), None)
            if not ref_word:
                continue

            x_ref = (ref_word['x0'] + ref_word['x1']) / 2
            y_ref = ref_word['top']

            values = []
            for w in words:
                try:
                    val = float(w['text'].replace(",", "."))
                except ValueError:
                    continue

                x_c = (w['x0'] + w['x1']) / 2
                y_c = w['top']

                # Condition : proche horizontalement ET en dessous d'une certaine distance verticale
                if (x_ref - tolerance_left <= x_c <= x_ref + tolerance_right) and (y_c > y_ref + min_dy_from_label):
                    values.append((y_c, val))

            results[keyword] = [v for _, v in sorted(values)]

    return results


# === Lancement ===
pdf_path = choose_pdf()
if pdf_path:
    keywords = ["Pl", "Pf*", "Module", "EM"]
    extracted = extract_values_near_keywords(
        pdf_path,
        keywords,
        tolerance_left=10,
        tolerance_right=30,
        min_dy_from_label=50  # ← distance minimale verticale sous le mot-clé
    )
    for label, vals in extracted.items():
        print(f"\n--- Valeurs extraites pour {label} ---")
        for v in vals:
            print(v)
else:
    print("Aucun fichier sélectionné.")