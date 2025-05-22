import pdfplumber
import logging
import matplotlib.pyplot as plt
from matplotlib.patches import Rectangle
from tkinter import Tk, filedialog

logging.getLogger("pdfminer").setLevel(logging.ERROR)


# === Script : EXTRACT VALUE UNDER KEYWORD IN A TABLE-LIKE FORMAT - ESSAIS PRESSIOMETRIQUES ===
# = v2 : Debug by highlight of keyword in the page : still only one page at a time
#
class PDFKeywordExtractor:
    def __init__(self, pdf_path, keywords, dpi=150, tolerances=None):
        self.pdf_path = pdf_path
        self.keywords = keywords
        self.dpi = dpi
        self.tolerances = tolerances or {
            "tolerance_left": 10,
            "tolerance_right": 30,
            "min_dy_from_label": 50,
            "tolerance_same_column": 15
        }

    def pt_to_px(self, val):
        return val * self.dpi / 72

    def get_keyword_x_positions(self, words):
        positions = {}
        for kw in self.keywords:
            for w in words:
                if w['text'].strip().lower() == kw.lower():
                    x = (w['x0'] + w['x1']) / 2
                    positions[kw] = x
                    break
        return positions

    # =========================== DEBUG =======================================
    #
    def highlight_keywords_on_page(self, page):
        words = page.extract_words()
        im = page.to_image(resolution=self.dpi)
        pil_image = im.original

        fig, ax = plt.subplots()
        ax.imshow(pil_image, origin="upper")


        print("\n--- Coordonnées des mots-clés trouvés (en pts) ---")
        found_any = False
        for word in words:
            text = word.get('text', '')
            for kw in self.keywords:
                if text.strip().lower() == kw.lower():
                    x0 = self.pt_to_px(word['x0'])
                    x1 = self.pt_to_px(word['x1'])
                    width = x1 - x0

                    top = self.pt_to_px(word['top'])
                    height = self.pt_to_px(word['bottom'] - word['top'])

                    rect = Rectangle((x0, top), width, height, edgecolor='red', linewidth=2, facecolor='none')
                    ax.add_patch(rect)

                    center_x = x0 + width / 2
                    center_y = top + height / 2
                    ax.plot(center_x, center_y, 'ro', markersize=3)

                    print(f"{kw:<10} → x={word['x0']:.1f}, y={word['top']:.1f}")
                    found_any = True
                    break

        if not found_any:
            print("❗ Aucun mot-clé trouvé pour surlignage.")

        ax.set_title("Mots-clés surlignés")
        ax.axis("off")
        plt.tight_layout()
        plt.show()
    #
    # =========================== DEBUG =======================================

    def extract_values_near_keyword(self, words, keyword):
        ref_word = next((w for w in words if w['text'].strip().lower() == keyword.lower()), None)
        if not ref_word:
            return []

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

            if (x_ref - self.tolerances['tolerance_left'] <= x_c <= x_ref + self.tolerances['tolerance_right']) \
               and (y_c > y_ref + self.tolerances['min_dy_from_label']):
                values.append((y_c, val))

        return [v for _, v in sorted(values)]

    def process_first_page(self):
        with pdfplumber.open(self.pdf_path) as pdf:
            page = pdf.pages[0]
            words = page.extract_words()

            # =========================== DEBUG =======================================
            # self.highlight_keywords_on_page(page)
            # =========================== DEBUG =======================================

            x_positions = self.get_keyword_x_positions(words)
            if "Pf*" in x_positions and "Pl*" in x_positions:
                distance = abs(x_positions["Pf*"] - x_positions["Pl*"])
                print(f"Distance entre Pf et Pl : {distance:.1f} pts")
                if distance <= self.tolerances['tolerance_same_column']:
                    print("✅ Pf et Pl sont dans la même colonne → traitement spécial")
                else:
                    print("❌ Pf et Pl sont éloignés → traitement standard")
            else:
                print("❗ Pf ou Pl non trouvé(s)")

            for kw in self.keywords:
                values = self.extract_values_near_keyword(words, kw)
                if values:
                    print(f"\n--- Valeurs extraites pour {kw} ---")
                    for v in values:
                        print(v)
                else:
                    print(f"\nAucune valeur trouvée pour {kw}")


# === Lancement ===
def choose_pdf():
    root = Tk()
    root.withdraw()
    return filedialog.askopenfilename(
        title="Choisir un fichier PDF",
        filetypes=[("Fichiers PDF", "*.pdf")]
    )


if __name__ == "__main__":
    pdf_path = choose_pdf()
    if pdf_path:
        extractor = PDFKeywordExtractor(
            pdf_path,
            keywords=["Pl*", "Pf*", "Module", "EM"]
        )
        extractor.process_first_page()
    else:
        print("Aucun fichier sélectionné.")
