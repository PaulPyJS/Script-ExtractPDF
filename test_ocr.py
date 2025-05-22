import re

import cv2
import numpy as np
import pytesseract
import pandas as pd
import os
from PIL import Image

pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

def extraire_table_depuis_image(image_path, output_excel="ocr_grid_output.xlsx", zoom_factor=3):
    img = cv2.imread(image_path)
    if img is None:
        print("❌ Image introuvable.")
        return

    # ZOOM to test
    img = cv2.resize(img, (img.shape[1]*zoom_factor, img.shape[0]*zoom_factor), interpolation=cv2.INTER_CUBIC)
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    gray = cv2.bilateralFilter(gray, 9, 75, 75)

    # TEST
    _, binary = cv2.threshold(gray, 128, 255, cv2.THRESH_BINARY_INV)

    # Trying to catch lines (h,v)
    horizontal_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (40, 1))
    detect_horizontal = cv2.morphologyEx(binary, cv2.MORPH_OPEN, horizontal_kernel, iterations=2)

    vertical_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, 40))
    detect_vertical = cv2.morphologyEx(binary, cv2.MORPH_OPEN, vertical_kernel, iterations=2)


    table_mask = cv2.add(detect_horizontal, detect_vertical)

    # Trying to catch cells
    contours, _ = cv2.findContours(table_mask, cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE)
    boxes = [cv2.boundingRect(c) for c in contours]
    boxes = [b for b in boxes if b[2] > 30 and b[3] > 20]


    rows = {}
    for box in boxes:
        x, y, w, h = box
        row_key = y // 10
        rows.setdefault(row_key, []).append(box)


    sorted_rows = []
    for row_key in sorted(rows.keys()):
        row = sorted(rows[row_key], key=lambda b: b[0])  # tri x (colonne)
        sorted_rows.append(row)

    # OCR cells by cells TEST
    tableau = []
    for row in sorted_rows:
        ligne = []
        for x, y, w, h in row:
            roi = img[y:y + h, x:x + w]
            config = "--psm 6 -c preserve_interword_spaces=1" # NEW CONFIG TO TEST WITH PSM 6 7 8 9 10 etc
            text = pytesseract.image_to_string(roi, config = config, lang="fra").strip()
            text = re.sub(r"[^\x00-\x7F]+", " ", text)
            text = re.sub(r"\s+", " ", text)
            ligne.append(text)
        tableau.append(ligne)

    df = pd.DataFrame(tableau)
    sheet = os.path.splitext(os.path.basename(image_path))[0][:31]
    df.to_excel(output_excel, sheet_name=sheet, index=False)
    return os.path.splitext(os.path.basename(image_path))[0][:31], df

def extraire_plusieurs_images(image_paths, output_excel="ocr_grid_output.xlsx"):
    with pd.ExcelWriter(output_excel, engine="openpyxl") as writer:
        for image_path in image_paths:
            print(f"🔍 Traitement : {image_path}")
            sheet_name, df = extraire_table_depuis_image(image_path)
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            print(f"✅ Feuille ajoutée : {sheet_name}")
    print(f"\n📁 Fichier Excel généré : {output_excel}")



if __name__ == "__main__":
    images = ["TABLE_1.png", "TABLE_2.png"]
    extraire_plusieurs_images(images, output_excel="ocr_grid_output.xlsx")


