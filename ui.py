import os
import tkinter as tk
from tkinter import ttk, filedialog, Menu
import ttkbootstrap as tb
from filelist import add_file, remove_file, load_file_list, save_file_list
from printer import print_file, get_printers

class PrintFlowApp(tb.Window):
    def __init__(self):
        super().__init__(title="PrintFlow - Stampante PDF e Word in serie", themename="cosmo")
        self.geometry("700x500")
        self.file_list = load_file_list()
        self.build_ui()
        self.refresh_list()

    def build_ui(self):
        top = ttk.Frame(self)
        top.pack(fill=tk.X, padx=10, pady=5)

        ttk.Label(top, text="Stampante:").pack(side=tk.LEFT)
        self.printer_var = tk.StringVar()
        printers = get_printers()
        self.printer_combo = ttk.Combobox(top, textvariable=self.printer_var, values=printers, width=40)
        self.printer_combo.set(printers[0] if printers else "")
        self.printer_combo.pack(side=tk.LEFT, padx=5)

        list_frame = ttk.Frame(self)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        self.listbox = tk.Listbox(list_frame, selectmode=tk.SINGLE, font=("Arial", 11))
        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.listbox.config(yscrollcommand=scrollbar.set)

        self.listbox.bind("<Button-3>", self.show_context_menu)

        self.context_menu = Menu(self, tearoff=0)
        self.context_menu.add_command(label="Rimuovi", command=self.remove_selected)

        bottom = ttk.Frame(self)
        bottom.pack(fill=tk.X, padx=10, pady=5)

        ttk.Button(bottom, text="Aggiungi file", command=self.add_files).pack(side=tk.LEFT)
        ttk.Button(bottom, text="▲", width=3, command=self.move_up).pack(side=tk.LEFT, padx=2)
        ttk.Button(bottom, text="▼", width=3, command=self.move_down).pack(side=tk.LEFT, padx=2)
        ttk.Button(bottom, text="Stampa tutti", command=self.print_all).pack(side=tk.LEFT, padx=5)

        self.status_label = ttk.Label(self, text="Pronto")
        self.status_label.pack(fill=tk.X, padx=10)

        self.progress = ttk.Progressbar(self, mode="determinate")
        self.progress.pack(fill=tk.X, padx=10, pady=5)

    def refresh_list(self):
        self.listbox.delete(0, tk.END)
        for i, filepath in enumerate(self.file_list):
            basename = os.path.basename(filepath)
            self.listbox.insert(tk.END, f"{i+1}. {basename}")

    def add_files(self):
        files = filedialog.askopenfilenames(filetypes=[("PDF e Word", "*.pdf *.doc *.docx")])
        for f in files:
            add_file(self.file_list, f)
        save_file_list(self.file_list)
        self.refresh_list()

    def remove_selected(self):
        selection = self.listbox.curselection()
        if selection:
            remove_file(self.file_list, selection[0])
            save_file_list(self.file_list)
            self.refresh_list()

    def show_context_menu(self, event):
        try:
            self.listbox.selection_clear(0, tk.END)
            self.listbox.selection_set(self.listbox.nearest(event.y))
            self.context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.context_menu.grab_release()

    def move_up(self):
        idx = self.listbox.curselection()
        if idx and idx[0] > 0:
            i = idx[0]
            self.file_list[i-1], self.file_list[i] = self.file_list[i], self.file_list[i-1]
            save_file_list(self.file_list)
            self.refresh_list()
            self.listbox.selection_set(i-1)

    def move_down(self):
        idx = self.listbox.curselection()
        if idx and idx[0] < len(self.file_list) - 1:
            i = idx[0]
            self.file_list[i], self.file_list[i+1] = self.file_list[i+1], self.file_list[i]
            save_file_list(self.file_list)
            self.refresh_list()
            self.listbox.selection_set(i+1)

    def print_all(self):
        self.progress["maximum"] = len(self.file_list)
        self.progress["value"] = 0
        for i, filepath in enumerate(self.file_list):
            self.status_label.config(text=f"Stampando: {os.path.basename(filepath)}")
            self.update_idletasks()
            print_file(filepath, self.printer_var.get())
            self.progress["value"] += 1
        self.status_label.config(text="Stampa completata.")
        self.update_idletasks()