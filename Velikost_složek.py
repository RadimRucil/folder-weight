import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import time
import itertools
import pandas as pd
from openpyxl import load_workbook
import os
import queue

class FolderSizeApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Folder Size Checker")

        # Inicializace proměnných pro progress bar a status
        self.progress_var = tk.DoubleVar()
        self.progress_var.set(0)
        self.status_var = tk.StringVar()
        self.status_var.set("Prohledávám adresář")

        self.create_widgets()
        self.done = False
        self.total_folders = 0
        self.current_index = 0
        self.queue = queue.Queue()
        self.folder_sizes = []
        self.base_path = None

    def create_widgets(self):
        # Hlavní rámec pro widgety
        self.frame = tk.Frame(self.root)
        self.frame.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

        # Label pro zobrazení stavu a pokroku
        self.status_label = tk.Label(self.frame, textvariable=self.status_var, font=("Arial", 12))
        self.status_label.pack(pady=(0, 10))

        # Progress bar
        self.progress_bar = ttk.Progressbar(self.frame, orient="horizontal", length=400, mode="determinate", variable=self.progress_var)
        self.progress_bar.pack(pady=(0, 10))

        # Rámec pro textový widget a scroll bar
        self.text_frame = tk.Frame(self.frame)
        self.text_frame.pack(padx=5, pady=5, fill=tk.BOTH, expand=True)

        # Textový widget pro zobrazení výstupu
        self.text_output = tk.Text(self.text_frame, height=20, width=80, wrap='word')
        self.text_output.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Scrollbar pro textový widget
        self.scrollbar = tk.Scrollbar(self.text_frame, command=self.text_output.yview)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.text_output.config(yscrollcommand=self.scrollbar.set)

        # Rámec pro tlačítka
        self.button_frame = tk.Frame(self.frame)
        self.button_frame.pack(pady=(5, 0))

        # Tlačítko pro výběr adresáře
        self.select_button = tk.Button(self.button_frame, text="Vybrat adresář", command=self.select_directory)
        self.select_button.pack(side=tk.LEFT, padx=5)

        # Tlačítko pro spuštění hledání
        self.start_button = tk.Button(self.button_frame, text="Spustit", command=self.start_search)
        self.start_button.pack(side=tk.LEFT, padx=5)
        self.start_button.config(state=tk.DISABLED)

        # Tlačítko pro uložení výsledků do Excelu
        self.save_button = tk.Button(self.button_frame, text="Uložit výsledek", command=self.save_to_excel)
        self.save_button.pack(side=tk.LEFT, padx=5)
        self.save_button.config(state=tk.DISABLED)

        # Tlačítko pro ukončení aplikace
        self.quit_button = tk.Button(self.button_frame, text="Ukončit", command=self.root.quit)
        self.quit_button.pack(side=tk.RIGHT, padx=5)

    def select_directory(self):
        # Otevření dialogu pro výběr adresáře
        self.base_path = filedialog.askdirectory(title="Vyberte základní adresář")
        if self.base_path:
            self.status_var.set(f"Vybraný adresář: {self.base_path}")
            self.start_button.config(state=tk.NORMAL)

    def start_search(self):
        # Spuštění hledání ve vybraném adresáři
        if not self.base_path:
            messagebox.showwarning("Varování", "Nejprve vyberte adresář.")
            return

        self.start_button.config(state=tk.DISABLED)
        self.done = False
        self.folder_sizes = []

        self.queue.queue.clear()  # Vyčištění fronty před spuštěním
        self.thread = threading.Thread(target=self.search_folders)
        self.thread.start()

        self.animate()

    def search_folders(self):
        # Hledání složek a měření jejich velikosti
        self.folder_sizes = self.list_folders_by_size(self.base_path)
        self.done = True
        self.queue.put(('progress', 100))  # Zajistí nastavení pokroku na 100%
        self.queue.put(('done', None))  # Signál o dokončení hledání
        self.root.after(0, self.update_text_output)
        self.root.after(0, lambda: self.save_button.config(state=tk.NORMAL))

    def list_folders_by_size(self, base_path, max_depth=2):
        # Seznam složek podle velikosti
        folder_sizes = []
        all_folders = []

        for root, dirs, _ in os.walk(base_path):
            depth = root.count(os.sep) - base_path.count(os.sep)
            if depth < max_depth:
                for d in dirs:
                    folder_path = os.path.join(root, d)
                    all_folders.append(folder_path)
            else:
                dirs[:] = []

        self.total_folders = len(all_folders)

        for index, folder_path in enumerate(all_folders):
            size_mb = self.get_folder_size(folder_path) / (1024 * 1024)
            folder_sizes.append((folder_path, size_mb))
            self.current_index = index + 1
            self.queue.put(('progress', self.current_index / self.total_folders * 100))
            time.sleep(0.1)  # Simulace zpoždění pro udržení GUI responzivního

        folder_sizes.sort(key=lambda x: x[1], reverse=True)
        return folder_sizes

    def get_folder_size(self, folder_path):
        # Vypočítání velikosti složky
        total_size = 0
        try:
            for dirpath, dirnames, filenames in os.walk(folder_path):
                for filename in filenames:
                    filepath = os.path.join(dirpath, filename)
                    if os.path.exists(filepath):
                        total_size += os.path.getsize(filepath)
        except PermissionError:
            pass
        return total_size

    def update_progress(self):
        # Aktualizace pokroku podle fronty
        while not self.queue.empty():
            command, value = self.queue.get()
            if command == 'progress':
                self.progress_var.set(value)
                self.status_var.set(f"Prohledávám adresář [{value:.1f}%]")
                self.root.update_idletasks()
            elif command == 'done':
                self.progress_var.set(100)  # Nastavení pokroku na 100% explicitně
                self.status_var.set("Hotovo!")
                self.start_button.config(state=tk.NORMAL)
                self.save_button.config(state=tk.NORMAL)
                self.root.update_idletasks()  # Zajištění, že progress bar je úplně aktualizován

    def update_text_output(self):
        # Aktualizace textového widgetu s výsledky
        self.text_output.delete(1.0, tk.END)
        if self.folder_sizes:
            for folder_path, size_mb in self.folder_sizes:
                self.text_output.insert(tk.END, f"{folder_path}\nVelikost: {size_mb:.1f} MB\n\n")

    def save_to_excel(self):
        # Uložení výsledků do Excelu
        file_name = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                               filetypes=[("Excel soubory", "*.xlsx")],
                                               initialfile="vysledky.xlsx")
        if not file_name:
            return

        df = pd.DataFrame(self.folder_sizes, columns=['Složka', 'Velikost (MB)'])
        df['Velikost (MB)'] = df['Velikost (MB)'].round(1)
        df['Velikost (GB)'] = (df['Velikost (MB)'] / 1024).round(1)
        df.to_excel(file_name, index=False)

        wb = load_workbook(file_name)
        ws = wb.active

        # Nastavení šířky sloupců
        column_widths = {
            'A': 71,
            'B': 12,
            'C': 12
        }

        for col, width in column_widths.items():
            ws.column_dimensions[col].width = width

        wb.save(file_name)
        messagebox.showinfo("Hotovo", f"Výsledky byly uloženy do {file_name}")

    def animate(self):
        # Animace pro zobrazení stavu hledání
        if self.done:
            return

        def update_status():
            if self.done:
                return

            for c in itertools.cycle(['       ', '*      ', '**     ', '***    ', '****   ', '*****  ']):
                if self.done:
                    self.status_var.set("Hotovo!")
                    return
                self.status_var.set(f"Prohledávám adresář [{self.progress_var.get():.1f}%] {c}")
                self.update_progress()  # Zkontrolování fronty na aktualizace
                self.root.update_idletasks()
                time.sleep(0.5)

        # Spuštění animace ve vlákně, aby GUI nezamrzlo
        threading.Thread(target=update_status, daemon=True).start()

def main():
    root = tk.Tk()
    app = FolderSizeApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
