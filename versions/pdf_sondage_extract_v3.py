import pdfplumber
import logging
import matplotlib.pyplot as plt
from matplotlib.patches import Rectangle
import tkinter as tk
from tkinter import Tk, filedialog, messagebox

logging.getLogger("pdfminer").setLevel(logging.ERROR)


# === Script : EXTRACT VALUE UNDER KEYWORD IN A TABLE-LIKE FORMAT - ESSAIS PRESSIOMETRIQUES ===
# = v2 : Debug by highlight of keyword in the page : still only one page at a time
# = v3 : Adding UI to validate data from extract to manually adapt if necessary
#
class PDFKeywordExtractor:
    def __init__(self, pdf_path, keywords, dpi=150, tolerances=None, column_distance_threshold=15):
        self.pdf_path = pdf_path
        self.keywords = keywords
        self.dpi = dpi
        self.column_distance_threshold = column_distance_threshold
        self.drag_data = {}

        # Tolerance par mot-clef
        self.tolerances = tolerances or {
            self.keywords[0]: {"left": 10, "right": 30, "min_dy": 50},
            self.keywords[1]: {"left": 10, "right": 30, "min_dy": 50},
            self.keywords[2]: {"left": 10, "right": 54, "min_dy": 50}
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

    # = UI for validation of the list, and suppression of strange values
    #
    def show_validation_ui(self, pf_list, pl_list, em_list):
        def update_listbox(listbox, data):
            listbox.delete(0, tk.END)
            for item in data:
                 listbox.insert(tk.END, item)

        def add_value(list_ref, listbox):
            popup = tk.Toplevel()
            popup.title("Ajouter une valeur")

            tk.Label(popup, text="Nouvelle valeur :").pack(padx=10, pady=5)
            entry = tk.Entry(popup)
            entry.pack(padx=10, pady=5)

            def submit():
                try:
                    val = float(entry.get())
                    list_ref.append(val)
                    update_listbox(listbox, list_ref)
                    popup.destroy()
                except ValueError:
                    messagebox.showwarning("Erreur", "Veuillez entrer un nombre valide.")

            tk.Button(popup, text="Valider", command=submit).pack(pady=10)
            entry.focus()

        def remove_selected(list_ref, listbox):
            selection = listbox.curselection()
            if selection:
                index = selection[0]
                del list_ref[index]
                update_listbox(listbox, list_ref)

        def validate_and_process():
            if len(pf_list) != len(pl_list) or len(pl_list) != len(em_list):
                messagebox.showerror("Erreur", "Les listes n'ont pas la même longueur.")
                return
            final_data = list(zip(pf_list, pl_list, em_list))
            print("\n📋 Données finales validées :")
            for row in final_data:
                print(row)
            root.destroy()

        root = tk.Tk()
        root.title("Vérification des valeurs extraites")

        tk.Label(root, text="UI de verification",
                 font=("Helvetica", 10)).pack(pady=(10, 5))

        frame = tk.Frame(root)
        frame.pack(padx=10, pady=5)

        lists = [(self.keywords[0], pf_list), (self.keywords[1], pl_list), (self.keywords[2], em_list)]
        listboxes = []

        for i, (label, data) in enumerate(lists):
            col_frame = tk.Frame(frame)
            col_frame.grid(row=0, column=i, padx=10)

            listbox = tk.Listbox(col_frame, height=20, width=10)
            listbox.pack()

            listboxes.append((listbox, data))
            update_listbox(listbox, data)

            tk.Label(col_frame, text=label).pack(pady=5)

            btn_frame = tk.Frame(col_frame)
            btn_frame.pack(pady=5)
            tk.Button(btn_frame, text="+", width=2, command=lambda d=data, lb=listbox: add_value(d, lb)).pack(
                side=tk.LEFT, padx=2)
            tk.Button(btn_frame, text="-", width=2, command=lambda d=data, lb=listbox: remove_selected(d, lb)).pack(
                side=tk.LEFT, padx=2)

            def on_drag_start(event, idx=i):
                self.drag_data[idx] = listbox.nearest(event.y)

            def on_drag_motion(event):
                pass  # pour un effet visuel éventuel

            def on_drag_release(event, idx=i, data_ref=data):
                from_idx = self.drag_data.get(idx)
                if from_idx is None:
                    return
                to_idx = listbox.nearest(event.y)
                if from_idx != to_idx:
                    item = data_ref.pop(from_idx)
                    data_ref.insert(to_idx, item)
                    update_listbox(listbox, data_ref)

            listbox.bind("<ButtonPress-1>", on_drag_start)
            listbox.bind("<B1-Motion>", on_drag_motion)
            listbox.bind("<ButtonRelease-1>", on_drag_release)



        tk.Button(root, text="Valider", command=validate_and_process).pack(pady=10)
        root.mainloop()



    def extract_values_near_keyword(self, words, keyword):
        ref_word = next((w for w in words if w['text'].strip().lower() == keyword.lower()), None)
        if not ref_word:
            return []

        tol = self.tolerances.get(keyword, {
            "left": 10,
            "right": 30,
            "min_dy": 50
        })

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

            if (x_ref - tol['left'] <= x_c <= x_ref + tol['right']) and (y_c > y_ref + tol['min_dy']):
                values.append((y_c, val))

        return [v for _, v in sorted(values)]




    def process_special_pf_pl(self, pf_values, pl_values, values_by_keyword):
        print(f"\n🔧 Traitement spécial des valeurs {self.keywords[0]}/{self.keywords[1]}...")

        if pf_values != pl_values:
            print(f"❌ Les valeurs extraites pour {self.keywords[0]} et {self.keywords[1]} ne sont pas identiques...")
            return

        if len(pf_values) % 2 != 0:
            print(f"⚠️ Nombre impair de valeurs pour {self.keywords[0]}/{self.keywords[1]}. Une valeur sera ignorée.")
            pf_values = pf_values[:-1]

        pf_final = []
        pl_final = []

        # Regrouper les valeurs deux par deux en partant de la fin
        for i in range(len(pf_values) - 2, -1, -2):
            a, b = pf_values[i], pf_values[i + 1]
            pf_final.append(min(a, b))
            pl_final.append(max(a, b))

        pf_final.reverse()
        pl_final.reverse()

        # Récupérer le 3e mot-clé dynamique (self.keywords[2], EM, etc.)
        if len(self.keywords) < 3:
            print("❗ Pas assez de mots-clés pour récupérer la troisième colonne.")
            return

        third_kw = self.keywords[-1]
        em_values = values_by_keyword.get(third_kw, [])

        print(f"\n✅ Longueurs obtenues : Pf = {len(pf_final)}, Pl = {len(pl_final)}, {third_kw} = {len(em_values)}")

        if len(pf_final) != len(em_values):
            print("❌ Les longueurs des listes ne correspondent pas. Impossible d'associer les lignes.")

        self.show_validation_ui(pf_final, pl_final, em_values)



    def process_first_page(self):
        with pdfplumber.open(self.pdf_path) as pdf:
            page = pdf.pages[0]
            words = page.extract_words()

            # =========================== DEBUG =======================================
            # self.highlight_keywords_on_page(page)
            # =========================== DEBUG =======================================

            # Add every values found by keyword
            values_by_keyword = {}
            for kw in self.keywords:
                values = self.extract_values_near_keyword(words, kw)
                values_by_keyword[kw] = values
                if values:
                    print(f"\n--- Valeurs extraites pour {kw} ---")
                    for v in values:
                        print(v)
                else:
                    print(f"\nAucune valeur trouvée pour {kw}")

            # Verify position of Pf & Pl keyword
            x_positions = self.get_keyword_x_positions(words)
            if self.keywords[0] in x_positions and self.keywords[1] in x_positions:
                distance = abs(x_positions[self.keywords[0]] - x_positions[self.keywords[1]])
                print(f"Distance entre Pf et Pl : {distance:.1f} pts")

                # = Based on the distance between Pf & Pl
                # Pf & Pl too close mean that they share the same column
                if distance <= self.column_distance_threshold:
                    print("✅ Pf et Pl sont dans la même colonne → traitement spécial")
                    self.process_special_pf_pl(
                        values_by_keyword[self.keywords[0]],
                        values_by_keyword[self.keywords[1]],
                        values_by_keyword
                    )
                # Separate columns
                else:
                    print("❌ Pf et Pl sont éloignés → traitement standard")
            else:
                print("❗ Pf ou Pl non trouvé(s)")





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
            keywords=["Pf*", "Pl*", "Module"],
        )
        extractor.process_first_page()
    else:
        print("Aucun fichier sélectionné.")
