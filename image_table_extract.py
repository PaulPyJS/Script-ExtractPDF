import os
import io
import re
import fitz  # PyMuPDF
import cv2
import numpy as np

import pytesseract
from PIL import Image
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

import pandas as pd
from pandas import ExcelWriter


# === Script : EXTRACT TABLE FROM IMAGE USING OCR ===
# ================= TESTS
#
def traiter_tableaux_image(pdf_path, page_nums, output_dir):
    pdf_name = os.path.splitext(os.path.basename(pdf_path))[0]
    output_file = os.path.join(output_dir, f"{pdf_name}_tables_image.xlsx")

    pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

    doc = fitz.open(pdf_path)
    image_tables = []

    for page_num in page_nums:
        try:
            page = doc.load_page(page_num - 1)  # Pages sont 0-indexées dans fitz
            pix = page.get_pixmap(dpi=300)
            img_data = pix.tobytes("png")
            image = Image.open(io.BytesIO(img_data))

            ocr_text = pytesseract.image_to_string(image, lang='fra')

            lignes = ocr_text.strip().split("\n")
            lignes = [l for l in lignes if l.strip()]
            tableau = [re.split(r"\s{2,}", ligne.strip()) for ligne in lignes]

            max_len = max(len(row) for row in tableau)
            tableau = [row + [""] * (max_len - len(row)) for row in tableau]

            df = pd.DataFrame(tableau)
            sheet_name = f"Page{page_num}_image"
            image_tables.append((sheet_name, df))
        except Exception as e:
            print(f"❌ Erreur OCR page {page_num} : {e}")

    if image_tables:
        with ExcelWriter(output_file, engine="openpyxl") as writer:
            for name, df in image_tables:
                df.to_excel(writer, sheet_name=name, index=False)
        print(f"✅ Export tableaux OCR : {output_file}")
    else:
        print("❌ Aucun tableau image extrait")

# Logic from https://livefiredev.com/how-to-extract-table-from-image-in-python-opencv-ocr/
# Using openCV + pysseract
def detecter_tableaux_par_image(pdf_path, page_nums, output_dir=None):
    from pytesseract import image_to_string
    tableaux = []

    for page_num in page_nums:
        print(f"🔎 Traitement OCR structuré page {page_num}")
        doc = fitz.open(pdf_path)
        page = doc.load_page(page_num - 1)  # 0-indexé

        pix = page.get_pixmap(dpi=300)
        img_data = pix.tobytes("png")
        image = Image.open(io.BytesIO(img_data))
        img_cv = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2BGR)
        gray = cv2.cvtColor(img_cv, cv2.COLOR_BGR2GRAY)

        _, binary = cv2.threshold(gray, 180, 255, cv2.THRESH_BINARY_INV)

        horizontal_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (40, 1))
        vertical_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, 40))
        horizontal_lines = cv2.morphologyEx(binary, cv2.MORPH_OPEN, horizontal_kernel, iterations=2)
        vertical_lines = cv2.morphologyEx(binary, cv2.MORPH_OPEN, vertical_kernel, iterations=2)
        table_mask = cv2.add(horizontal_lines, vertical_lines)

        contours, _ = cv2.findContours(table_mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        boxes = sorted([cv2.boundingRect(c) for c in contours], key=lambda b: (b[1], b[0]))

        lignes = []
        for x, y, w, h in boxes:
            cell_img = img_cv[y:y + h, x:x + w]
            cell_gray = cv2.cvtColor(cell_img, cv2.COLOR_BGR2GRAY)
            text = image_to_string(cell_gray, lang="fra", config="--psm 6").strip()
            lignes.append((y, x, text))  # pour trier ensuite

            cv2.rectangle(img_cv, (x, y), (x + w, y + h), (0, 255, 0), 1)

            if output_dir:
                img_name = f"{os.path.splitext(os.path.basename(pdf_path))[0]}_page{page_num}_x{x}_y{y}.png"
                img_path = os.path.join(output_dir, img_name)
                cv2.imwrite(img_path, cell_img)

        lignes.sort()
        texte_organise = [t[2] for t in lignes]
        n = max(1, len(texte_organise) // 5)  # devine ~5 colonnes
        tableau_final = [texte_organise[i:i + n] for i in range(0, len(texte_organise), n)]

        df = pd.DataFrame(tableau_final)
        tableaux.append((f"Page{page_num}_image", df))

        # ==================== DEBUG ==========================
        # plt.figure(figsize=(20, 10))
        # plt.imshow(cv2.cvtColor(img_cv, cv2.COLOR_BGR2RGB))
        # plt.title(f"Détection cellules - Page {page_num}")
        # plt.axis("off")
        # plt.show()


    if output_dir and tableaux:
        output_file = os.path.join(output_dir, f"{os.path.basename(pdf_path).split('.')[0]}_tableaux_image_test.xlsx")
        with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
            for name, df in tableaux:
                df.to_excel(writer, sheet_name=name, index=False)
        print(f"✅ Export des tableaux OCR dans : {output_file}")

    ocr_output = os.path.join(output_dir, f"{os.path.basename(pdf_path).split('.')[0]}_ocr_tables.xlsx")
    tableaux_output_dir = output_dir
    ocr_sur_images_decoupees(tableaux_output_dir, ocr_output, zoom_factor=2)


def ocr_sur_images_decoupees(dossier_images, output_file=None, zoom_factor=2):
    import tkinter as tk
    from tkinter import ttk

    tableaux = []
    images = [f for f in os.listdir(dossier_images) if f.lower().endswith(".png")]

    if not images:
        print("❌ Aucun fichier image trouvé pour OCR.")
        return []

    # Test progress bar
    progress_win = tk.Toplevel()
    progress_win.title("Traitement OCR")
    tk.Label(progress_win, text="OCR des tableaux en cours...").pack(padx=10, pady=5)
    progress_label = tk.Label(progress_win, text="Initialisation...")
    progress_label.pack()
    progress_bar = ttk.Progressbar(progress_win, length=300, mode="determinate", maximum=len(images))
    progress_bar.pack(padx=10, pady=10)
    progress_win.update()

    for i, image_file in enumerate(sorted(images), start=1):
        image_path = os.path.join(dossier_images, image_file)
        image = Image.open(image_path)

        new_size = (image.width * zoom_factor, image.height * zoom_factor)
        image_zoomed = image.resize(new_size, Image.Resampling.LANCZOS)

        ocr_text = pytesseract.image_to_string(image_zoomed, lang='fra')

        lignes = ocr_text.strip().split("\n")
        lignes = [l for l in lignes if l.strip()]
        tableau = [re.split(r"\s{2,}", ligne.strip()) for ligne in lignes]

        max_len = max((len(row) for row in tableau), default=0)
        tableau = [row + [""] * (max_len - len(row)) for row in tableau]

        df = pd.DataFrame(tableau)
        feuille = os.path.splitext(image_file)[0]
        tableaux.append((feuille, df))

        progress_label.config(text=f"Traitement : {image_file} ({i}/{len(images)})")
        progress_bar["value"] = i
        progress_win.update()

    progress_win.destroy()



    if output_file:
        with ExcelWriter(output_file, engine="openpyxl") as writer:
            for feuille, df in tableaux:
                df.to_excel(writer, sheet_name=feuille[:31], index=False)
        print(f"✅ OCR terminé. Résultat : {output_file}")

    return tableaux