import pdfplumber
import logging
import matplotlib.pyplot as plt
from matplotlib.patches import Rectangle
import tkinter as tk
from tkinter import Tk, filedialog, messagebox

import statistics
from decimal import Decimal, getcontext
getcontext().prec = 10

logging.getLogger("pdfminer").setLevel(logging.ERROR)


# === Script : EXTRACT VALUE UNDER KEYWORD IN A TABLE-LIKE FORMAT - ESSAIS PRESSIOMETRIQUES ===
# = v2 : Debug by highlight of keyword in the page : still only one page at a time
# = v3 : Adding UI to validate data from extract to manually adapt if necessary
# = v4 : Adding Depth from user to create list based on a STEP, DEPTH START/END
# = v5 : Adapting code to detect variation between Y coordinate from the values = No values or issues
#        Allowing user to visualize these variations within UI

def detect_y_anomalies(y_val_list, keyword):
    if len(y_val_list) < 3:
        return [v for _, v in y_val_list], []

    y_val_list = sorted(y_val_list, key=lambda x: x[0])
    y_positions = [y for y, _ in y_val_list]
    dy_list = [y2 - y1 for y1, y2 in zip(y_positions, y_positions[1:])]

    median_dy = statistics.median(dy_list)
    min_dy = 0.7 * median_dy
    max_dy = 1.3 * median_dy

    print(f"\n📏 Médiane des écarts Y pour '{keyword}': {median_dy:.2f} pts")
    print(f"🔍 Seuils → trop petit: < {min_dy:.2f} pts | trop grand: > {max_dy:.2f} pts")
    print(f"↕️ Écarts Y pour '{keyword}':")

    output = []
    logs = []

    highlight_indices = []

    for i in range(len(y_val_list) - 1):
        y1, v1 = y_val_list[i]
        y2, v2 = y_val_list[i + 1]
        dy = abs(y2 - y1)

        print(f"  dy[{i}] = {dy:.2f} pts entre {v1} et {v2}")

        output.append(v1)

        dy = Decimal(str(y2 - y1))
        median_dy = Decimal(str(statistics.median(dy_list)))
        min_dy = median_dy * Decimal("0.7")
        max_dy = median_dy * Decimal("1.3")

        if dy > max_dy:
            print("    → TROU détecté")
            output.append(None)
            logs.append(f" NULL : Trou détecté pour '{keyword}' entre {v1} et {v2} (écart Y = {dy:.1f} pts)")
        elif dy < min_dy:
            print("    → TROP PETIT")
            highlight_indices.append(len(output) - 1)
            highlight_indices.append(len(output))
            logs.append(f" ⚠️  : Espacement trop petit pour '{keyword}' entre {v1} et {v2} (écart Y = {dy:.1f} pts)")

    output.append(y_val_list[-1][1])  # Dernière valeur

    return output, logs, highlight_indices






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
    def show_validation_ui(self, pf_list, pl_list, em_list, depth_list, pf_red=None, pl_red=None, em_red=None):
        pf_red = pf_red or []
        pl_red = pl_red or []
        em_red = em_red or []
        #
        def update_listbox(listbox, data, red_indices=None):
            red_indices = red_indices or []
            listbox.delete(0, tk.END)

            for i, item in enumerate(data):
                # Cas valeur tuple (y, val)
                if isinstance(item, tuple) and len(item) == 2:
                    display_val = item[1]
                    # INDEX 0 stand for COORDINATE Y
                    # INDEX 1 is the real values
                elif item is None:
                    display_val = "NULL"
                else:
                    display_val = item

                listbox.insert(tk.END, display_val)

                if str(display_val).upper() == "NULL":
                    listbox.itemconfig(i, {'fg': '#8B0000'})

                if i in red_indices:
                    listbox.itemconfig(i, {'fg': 'red'})

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

        def set_null(list_ref, listbox):
            selection = listbox.curselection()
            if selection:
                index = selection[0]
                list_ref[index] = "NULL"
                update_listbox(listbox, list_ref)

        def validate_and_process():
            if not (len(pf_list) == len(pl_list) == len(em_list) == len(depth_list)):
                messagebox.showerror("Erreur",
                                     "Les listes n'ont pas la même longueur.\nVeuillez corriger avant de valider.")
                return
            final_data = list(zip(pf_list, pl_list, em_list, depth_list))
            print("\n📋 Données finales validées :")
            for row in final_data:
                print(row)
            root.destroy()

        root = tk.Tk()
        root.title("Vérification des valeurs extraites")

        # --- Titre général ---
        tk.Label(root, text="UI de verification", font=("Helvetica", 11, "bold")).pack(pady=(10, 5))

        # --- Frame principal ---
        frame = tk.Frame(root)
        frame.pack(padx=10, pady=5)

        # --- Ligne des entêtes des colonnes ---
        for col_index, (label, _) in enumerate([
            ("Profondeur", depth_list),
            (self.keywords[0], pf_list),
            (self.keywords[1], pl_list),
            (self.keywords[2], em_list)
        ]):
            tk.Label(frame, text=label, font=("Helvetica", 9, "normal"), width=10).grid(row=0, column=col_index, pady=(0, 5))

        # --- Frame pour les colonnes ---
        lists = [
            ("Profondeur", depth_list),
            (self.keywords[0], pf_list),
            (self.keywords[1], pl_list),
            (self.keywords[2], em_list)
        ]
        listboxes = []

        for i, (label, data) in enumerate(lists):
            col_frame = tk.Frame(frame)
            col_frame.grid(row=1, column=i, padx=10)

            listbox = tk.Listbox(col_frame, height=22, width=10, font=("Helvetica", 10))
            listbox.pack()

            listboxes.append((listbox, data))
            if label == self.keywords[0]:
                update_listbox(listbox, data, red_indices=pf_red)
            elif label == self.keywords[1]:
                update_listbox(listbox, data, red_indices=pl_red)
            elif label == self.keywords[2]:
                update_listbox(listbox, data, red_indices=em_red)
            else:
                update_listbox(listbox, data)

            # --- Boutons + / - ---
            btn_frame = tk.Frame(col_frame)
            btn_frame.pack(pady=(6, 0))

            btn_style = {"font": ("Helvetica", 10, "bold"), "width": 2}
            tk.Button(btn_frame, text="+", **btn_style, command=lambda d=data, lb=listbox: add_value(d, lb)).pack(
                side=tk.LEFT, padx=2)
            tk.Button(btn_frame, text="-", **btn_style, command=lambda d=data, lb=listbox: remove_selected(d, lb)).pack(
                side=tk.LEFT, padx=2)

            # --- Bouton NULL ---
            tk.Button(col_frame, text="NULL", width=6, relief="groove", font=("Helvetica", 9),
                      command=lambda d=data, lb=listbox: set_null(d, lb)).pack(pady=(4, 2))

            # Drag and drop
            def on_drag_start(event, idx=i):
                self.drag_data[idx] = listbox.nearest(event.y)

            def on_drag_motion(event):
                pass

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

        # --- Bouton valider global ---
        tk.Button(root, text="Valider", command=validate_and_process).pack(pady=10)
        root.mainloop()

    def ask_user_for_depth_range(self):
        depth_values = []

        def on_submit():
            try:
                d_start = float(entry_start.get())
                d_end = float(entry_end.get())
                d_step = float(entry_step.get())

                if d_start >= d_end or d_step <= 0:
                    raise ValueError

                for i in range(int((d_end - d_start) / d_step) + 1):
                    depth_values.append(round(d_start + i * d_step, 3))

                win.destroy()
            except ValueError:
                messagebox.showerror("Erreur", "Veuillez entrer des valeurs valides.")

        win = tk.Toplevel()
        win.title("Paramètres du sondage")
        win.attributes("-topmost", True)
        win.lift()

        tk.Label(win, text="Profondeur de départ :").pack()
        entry_start = tk.Entry(win)
        entry_start.pack()

        tk.Label(win, text="Profondeur de fin :").pack()
        entry_end = tk.Entry(win)
        entry_end.pack()

        tk.Label(win, text="Pas (ex: 0.2) :").pack()
        entry_step = tk.Entry(win)
        entry_step.pack()

        tk.Button(win, text="Valider", command=on_submit).pack(pady=10)

        win.grab_set()
        win.wait_window()

        return depth_values



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

        return sorted(values, key=lambda x: x[0])




    # = LIST PROCESSING FOR STANDARD AND SPECIAL PF / PL
    #
    def process_standard_pf_pl(self, values_by_keyword):
        print("\n🔧 Traitement standard des valeurs...")

        pf_values = values_by_keyword.get(self.keywords[0], [])
        pl_values = values_by_keyword.get(self.keywords[1], [])
        em_values = values_by_keyword.get(self.keywords[2], [])

        pf_list, pf_logs, pf_red = detect_y_anomalies(pf_values, self.keywords[0])
        pl_list, pl_logs, pl_red = detect_y_anomalies(pl_values, self.keywords[1])
        em_list, em_logs, em_red = detect_y_anomalies(em_values, self.keywords[2])

        depths = self.ask_user_for_depth_range()
        print(depths)

        print(f"\n✅ Longueurs obtenues : {self.keywords[0]} = {len(pf_values)}, "
              f"{self.keywords[1]} = {len(pl_values)}, {self.keywords[2]} = {len(em_values)}")

        # STANDARD but user can still adjust extrem data
        self.show_validation_ui(pf_list, pl_list, em_list, depths, pf_red, pl_red, em_red)

        if not depths:
            print("❗ Aucune profondeur saisie.")
            return

    def process_special_pf_pl(self, pf_values, pl_values, values_by_keyword):
        print(f"\n🔧 Traitement spécial des valeurs {self.keywords[0]}/{self.keywords[1]}...")

        depths = self.ask_user_for_depth_range()
        print(depths)

        if pf_values != pl_values:
            print(f"❌ Les valeurs extraites pour {self.keywords[0]} et {self.keywords[1]} ne sont pas identiques...")
            return

        if len(pf_values) % 2 != 0:
            print(f"⚠️ Nombre impair de valeurs pour {self.keywords[0]}/{self.keywords[1]}. Une valeur sera ignorée.")
            pf_values = pf_values[:-1]

        pf_final = []
        pl_final = []

        for i in range(len(pf_values) - 2, -1, -2):
            y1, v1 = pf_values[i]
            y2, v2 = pf_values[i + 1]

            if v1 < v2:
                pf_final.append((y1, v1))
                pl_final.append((y2, v2))
            else:
                pf_final.append((y2, v2))
                pl_final.append((y1, v1))

        pf_final.reverse()
        pl_final.reverse()

        # Récupérer le 3e mot-clé dynamique (self.keywords[2], EM, etc.)
        if len(self.keywords) < 3:
            print("❗ Pas assez de mots-clés pour récupérer la troisième colonne.")
            return

        third_kw = self.keywords[-1]
        em_values_yval = values_by_keyword.get(third_kw, [])


        # Anomalie Y
        pf_final_vals, pf_logs, pf_red = detect_y_anomalies(pf_final, self.keywords[0])
        pl_final_vals, pl_logs, pl_red = detect_y_anomalies(pl_final, self.keywords[1])
        em_vals, em_logs, em_red = detect_y_anomalies(em_values_yval, self.keywords[2])


        print(f"\n✅ Longueurs obtenues : Pf = {len(pf_final)}, Pl = {len(pl_final)}, {third_kw} = {len(em_values_yval)}")

        if len(pf_final) != len(em_values_yval):
            print("❌ Les longueurs des listes ne correspondent pas. Impossible d'associer les lignes.")

        self.show_validation_ui(pf_final_vals, pl_final_vals, em_vals, depths, pf_red, pl_red, em_red)

        if not depths:
            print("❗ Aucune profondeur saisie.")
            return



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
                    print("✅ Pf et Pl sont éloignés → traitement standard")
                    self.process_standard_pf_pl(values_by_keyword)
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
