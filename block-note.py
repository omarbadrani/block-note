import tkinter as tk
from tkinter import filedialog, colorchooser, messagebox, simpledialog, font, ttk
from tkinter.scrolledtext import ScrolledText
from PIL import Image, ImageTk
import pandas as pd

class BlocNoteAvance(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Bloc-Notes Avancé")
        self.geometry("1000x750")
        self.configure(bg="#2b2b2b")

        # Styles ttk (dark theme)
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("Treeview",
                        background="#3c3f41",
                        foreground="white",
                        fieldbackground="#3c3f41",
                        font=('Consolas', 11))
        style.configure("Treeview.Heading",
                        background="#6c7b8b",
                        foreground="white",
                        font=('Arial', 12, 'bold'))
        style.map("Treeview", background=[('selected', '#4b6eaf')])
        style.configure("TButton", background="#4b6eaf", foreground="white")

        # Frames
        self.main_frame = tk.Frame(self, bg="#2b2b2b")
        self.main_frame.pack(expand=True, fill='both')

        # Texte avec scroll
        self.text_area = ScrolledText(self.main_frame, wrap=tk.WORD, font=("Consolas", 14), height=15,
                                     bg="#1e1e1e", fg="white", insertbackground="white", undo=True)
        self.text_area.pack(expand=False, fill='x')

        # Barre d'état
        self.status_bar = tk.Label(self.main_frame, text="Caractères: 0  Ligne: 1  Colonne: 0",
                                   anchor='w', bg="#2b2b2b", fg="lightgray")
        self.status_bar.pack(fill='x')

        # Treeview pour Excel
        self.tree_frame = tk.Frame(self.main_frame)
        self.tree_frame.pack(expand=True, fill='both')
        self.tree = ttk.Treeview(self.tree_frame)
        self.tree.pack(side='left', expand=True, fill='both')
        self.tree_scrollbar = ttk.Scrollbar(self.tree_frame, orient="vertical", command=self.tree.yview)
        self.tree_scrollbar.pack(side='right', fill='y')
        self.tree.configure(yscrollcommand=self.tree_scrollbar.set)
        self.tree_frame.pack_forget()

        # Menu
        self.menu_bar = tk.Menu(self)
        self.config(menu=self.menu_bar)

        # Menu Fichier
        file_menu = tk.Menu(self.menu_bar, tearoff=0)
        file_menu.add_command(label="Nouveau", command=self.new_file)
        file_menu.add_command(label="Ouvrir...", command=self.open_file)
        file_menu.add_command(label="Ouvrir Excel...", command=self.open_excel_file)
        file_menu.add_command(label="Sauvegarder", command=self.save_file)
        file_menu.add_command(label="Exporter vers Excel...", command=self.export_to_excel)
        file_menu.add_separator()
        file_menu.add_command(label="Quitter", command=self.quit)
        self.menu_bar.add_cascade(label="Fichier", menu=file_menu)

        # Menu Édition
        edit_menu = tk.Menu(self.menu_bar, tearoff=0)
        edit_menu.add_command(label="Couper", command=lambda: self.text_area.event_generate("<<Cut>>"))
        edit_menu.add_command(label="Copier", command=lambda: self.text_area.event_generate("<<Copy>>"))
        edit_menu.add_command(label="Coller", command=lambda: self.text_area.event_generate("<<Paste>>"))
        edit_menu.add_separator()
        edit_menu.add_command(label="Rechercher & Remplacer", command=self.search_replace)
        self.menu_bar.add_cascade(label="Édition", menu=edit_menu)

        # Menu Format
        format_menu = tk.Menu(self.menu_bar, tearoff=0)
        format_menu.add_command(label="Gras", command=self.toggle_bold)
        format_menu.add_command(label="Italique", command=self.toggle_italic)
        format_menu.add_command(label="Souligné", command=self.toggle_underline)
        format_menu.add_separator()
        format_menu.add_command(label="Couleur du texte", command=self.choose_color)
        format_menu.add_separator()
        format_menu.add_command(label="Changer police...", command=self.change_font)
        self.menu_bar.add_cascade(label="Format", menu=format_menu)

        # Menu Insertion
        insert_menu = tk.Menu(self.menu_bar, tearoff=0)
        insert_menu.add_command(label="Insérer image", command=self.insert_image)
        insert_menu.add_command(label="Insérer tableau", command=self.insert_table)
        self.menu_bar.add_cascade(label="Insertion", menu=insert_menu)

        # Fonts pour tags
        self.font_bold = font.Font(self.text_area, self.text_area.cget("font"))
        self.font_bold.configure(weight="bold")
        self.font_italic = font.Font(self.text_area, self.text_area.cget("font"))
        self.font_italic.configure(slant="italic")
        self.font_underline = font.Font(self.text_area, self.text_area.cget("font"))
        self.font_underline.configure(underline=1)

        # Images refs
        self.image_refs = []

        # Bindings
        self.text_area.bind("<<Modified>>", self.update_status_bar)
        self.text_area.bind("<KeyRelease>", self.update_status_bar)
        self.text_area.edit_modified(False)

    def update_status_bar(self, event=None):
        text_content = self.text_area.get(1.0, tk.END)
        chars = len(text_content) - 1
        line, col = self.text_area.index(tk.INSERT).split('.')
        self.status_bar.config(text=f"Caractères: {chars}  Ligne: {line}  Colonne: {col}")

    def new_file(self):
        if messagebox.askyesno("Nouveau fichier", "Voulez-vous sauvegarder avant de créer un nouveau fichier ?"):
            self.save_file()
        self.text_area.delete(1.0, tk.END)
        self.hide_excel_view()

    def open_file(self):
        path = filedialog.askopenfilename(filetypes=[("Fichiers texte", "*.txt"), ("Tous fichiers", "*.*")])
        if path:
            try:
                with open(path, "r", encoding="utf-8") as file:
                    content = file.read()
                self.text_area.delete(1.0, tk.END)
                self.text_area.insert(tk.END, content)
                self.hide_excel_view()
            except Exception as e:
                messagebox.showerror("Erreur", f"Impossible d'ouvrir le fichier: {e}")

    def open_excel_file(self):
        path = filedialog.askopenfilename(filetypes=[("Fichiers Excel", "*.xlsx;*.xls")])
        if not path:
            return
        try:
            df = pd.read_excel(path)
            self.show_excel_data(df)
        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible d'ouvrir le fichier Excel:\n{e}")

    def show_excel_data(self, df):
        df = df.fillna(" ")
        self.tree_frame.pack(expand=True, fill='both')
        self.text_area.pack(expand=False, fill='x')
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = list(df.columns)
        self.tree["show"] = "headings"
        for col in df.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=130, anchor="center")
        for i, (_, row) in enumerate(df.iterrows()):
            values = list(row)
            tag = "evenrow" if i % 2 == 0 else "oddrow"
            self.tree.insert("", "end", values=values, tags=(tag,))
        self.tree.tag_configure("evenrow", background="#313335")
        self.tree.tag_configure("oddrow", background="#3c3f41")

    def hide_excel_view(self):
        self.tree_frame.pack_forget()

    def save_file(self):
        path = filedialog.asksaveasfilename(defaultextension=".txt",
                                            filetypes=[("Fichiers texte", "*.txt")])
        if path:
            try:
                content = self.text_area.get(1.0, tk.END)
                with open(path, "w", encoding="utf-8") as file:
                    file.write(content)
                self.hide_excel_view()
            except Exception as e:
                messagebox.showerror("Erreur", f"Impossible de sauvegarder le fichier: {e}")

    def toggle_tag(self, tag_name, font_style):
        try:
            current_tags = self.text_area.tag_names("sel.first")
        except tk.TclError:
            messagebox.showinfo("Info", "Veuillez sélectionner du texte pour appliquer ce style.")
            return
        if tag_name in current_tags:
            self.text_area.tag_remove(tag_name, "sel.first", "sel.last")
        else:
            self.text_area.tag_add(tag_name, "sel.first", "sel.last")
            self.text_area.tag_config(tag_name, font=font_style)

    def toggle_bold(self):
        self.toggle_tag("bold", self.font_bold)

    def toggle_italic(self):
        self.toggle_tag("italic", self.font_italic)

    def toggle_underline(self):
        self.toggle_tag("underline", self.font_underline)

    def choose_color(self):
        color_code = colorchooser.askcolor(title="Choisissez la couleur du texte")
        if color_code[1]:
            try:
                self.text_area.tag_add("colored", "sel.first", "sel.last")
                self.text_area.tag_config("colored", foreground=color_code[1])
            except tk.TclError:
                messagebox.showinfo("Info", "Veuillez sélectionner du texte pour changer sa couleur.")

    def change_font(self):
        fonts = sorted(list(font.families()))
        family = simpledialog.askstring("Police", "Entrez la famille de police:", initialvalue="Consolas")
        if family not in fonts:
            messagebox.showwarning("Attention", "Police inconnue. Utilisation de 'Consolas'.")
            family = "Consolas"
        size = simpledialog.askinteger("Taille", "Entrez la taille de la police:", initialvalue=14, minvalue=8, maxvalue=72)
        if size:
            try:
                self.text_area.tag_add("fontchange", "sel.first", "sel.last")
                font_selected = font.Font(family=family, size=size)
                self.text_area.tag_config("fontchange", font=font_selected)
            except tk.TclError:
                messagebox.showinfo("Info", "Veuillez sélectionner du texte pour changer la police.")

    def insert_image(self):
        path = filedialog.askopenfilename(filetypes=[("Images", "*.png;*.jpg;*.jpeg;*.gif")])
        if not path:
            return
        try:
            img = Image.open(path)
            img.thumbnail((300, 300))
            img_tk = ImageTk.PhotoImage(img)
            self.image_refs.append(img_tk)
            self.text_area.image_create(tk.INSERT, image=img_tk)
        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible d'insérer l'image:\n{e}")

    def insert_table(self):
        # Exemple simple d'insertion tableau ASCII
        table_ascii = (
            "+-------+-------+-------+\n"
            "| Col 1 | Col 2 | Col 3 |\n"
            "+-------+-------+-------+\n"
            "| Val 1 | Val 2 | Val 3 |\n"
            "| Val 4 | Val 5 | Val 6 |\n"
            "+-------+-------+-------+\n"
        )
        self.text_area.insert(tk.INSERT, table_ascii)

    def search_replace(self):
        def do_search():
            target = entry_search.get()
            replacement = entry_replace.get()
            content = self.text_area.get(1.0, tk.END)
            if target:
                new_content = content.replace(target, replacement)
                self.text_area.delete(1.0, tk.END)
                self.text_area.insert(1.0, new_content)
                top.destroy()
            else:
                messagebox.showinfo("Info", "Entrez un texte à rechercher.")

        top = tk.Toplevel(self)
        top.title("Rechercher & Remplacer")
        top.geometry("350x150")
        tk.Label(top, text="Rechercher:").pack(pady=5)
        entry_search = tk.Entry(top, width=30)
        entry_search.pack(pady=5)
        tk.Label(top, text="Remplacer par:").pack(pady=5)
        entry_replace = tk.Entry(top, width=30)
        entry_replace.pack(pady=5)
        tk.Button(top, text="Remplacer", command=do_search).pack(pady=10)

    def export_to_excel(self):
        path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                            filetypes=[("Fichier Excel", "*.xlsx")])
        if not path:
            return
        content = self.text_area.get(1.0, tk.END).strip()
        lines = content.splitlines()
        data = [line.split() for line in lines if line.strip()]
        try:
            df = pd.DataFrame(data)
            df.to_excel(path, index=False, header=False)
            messagebox.showinfo("Succès", f"Exportation réussie vers {path}")
        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible d'exporter vers Excel:\n{e}")

if __name__ == "__main__":
    app = BlocNoteAvance()
    app.mainloop()
