import csv
import os
import tkinter as tk
from datetime import datetime
from tkinter import filedialog, messagebox, ttk

CSV_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "metafores.csv")
COLUMNS = ["ΗΜΕΡΟΜΗΝΙΑ", "ΟΝΟΜΑΤΕΠΩΝΥΜΟ", "ΑΠΟ", "ΣΕ", "ΟΝΟΜΑ ΑΡΧΕΙΟΥ", "ΠΕΡΙΓΡΑΦΗ"]


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Καταχώρηση Μεταφοράς Αρχείων")
        self.geometry("1000x620")
        self.minsize(850, 520)

        self.records = self.load_records()
        self.vars = {
            "ΟΝΟΜΑΤΕΠΩΝΥΜΟ": tk.StringVar(),
            "ΑΠΟ": tk.StringVar(),
            "ΣΕ": tk.StringVar(),
            "ΟΝΟΜΑ ΑΡΧΕΙΟΥ": tk.StringVar(),
        }
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *_: self.refresh_table())

        self.build_form()
        self.build_table()
        self.refresh_table()

    # ---------- Αποθήκευση / φόρτωση ----------
    def load_records(self):
        if not os.path.exists(CSV_FILE):
            return []
        with open(CSV_FILE, newline="", encoding="utf-8-sig") as f:
            return [[row.get(c, "") for c in COLUMNS] for row in csv.DictReader(f)]

    def save_records(self):
        # utf-8-sig ώστε το Excel να διαβάζει σωστά τα ελληνικά
        with open(CSV_FILE, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f)
            writer.writerow(COLUMNS)
            writer.writerows(self.records)

    # ---------- Φόρμα ----------
    def build_form(self):
        form = ttk.LabelFrame(self, text=" Νέα καταχώρηση ", padding=12)
        form.pack(fill="x", padx=12, pady=(12, 6))
        form.columnconfigure(1, weight=1)
        form.columnconfigure(3, weight=1)

        ttk.Label(form, text="ΟΝΟΜΑΤΕΠΩΝΥΜΟ").grid(row=0, column=0, sticky="w", pady=4)
        self.name_entry = ttk.Entry(form, textvariable=self.vars["ΟΝΟΜΑΤΕΠΩΝΥΜΟ"])
        self.name_entry.grid(row=0, column=1, columnspan=3, sticky="ew", padx=8)

        ttk.Label(form, text="ΑΠΟ").grid(row=1, column=0, sticky="w", pady=4)
        ttk.Entry(form, textvariable=self.vars["ΑΠΟ"]).grid(row=1, column=1, sticky="ew", padx=8)

        ttk.Label(form, text="ΣΕ").grid(row=1, column=2, sticky="w", padx=(12, 0))
        ttk.Entry(form, textvariable=self.vars["ΣΕ"]).grid(row=1, column=3, sticky="ew", padx=8)

        ttk.Label(form, text="ΟΝΟΜΑ ΑΡΧΕΙΟΥ").grid(row=2, column=0, sticky="w", pady=4)
        ttk.Entry(form, textvariable=self.vars["ΟΝΟΜΑ ΑΡΧΕΙΟΥ"]).grid(
            row=2, column=1, columnspan=2, sticky="ew", padx=8
        )
        ttk.Button(form, text="Επιλογή αρχείου…", command=self.pick_file).grid(
            row=2, column=3, sticky="w", padx=8
        )

        ttk.Label(form, text="ΠΕΡΙΓΡΑΦΗ").grid(row=3, column=0, sticky="nw", pady=4)
        self.desc_text = tk.Text(form, height=3, wrap="word", font=("Segoe UI", 10))
        self.desc_text.grid(row=3, column=1, columnspan=3, sticky="ew", padx=8, pady=4)

        buttons = ttk.Frame(form)
        buttons.grid(row=4, column=0, columnspan=4, sticky="e", pady=(8, 0))
        ttk.Button(buttons, text="Καθαρισμός", command=self.clear_form).pack(side="left", padx=4)
        ttk.Button(buttons, text="Καταχώρηση", command=self.add_record).pack(side="left", padx=4)

        self.name_entry.focus()

    def pick_file(self):
        path = filedialog.askopenfilename(title="Επιλογή αρχείου")
        if path:
            self.vars["ΟΝΟΜΑ ΑΡΧΕΙΟΥ"].set(os.path.basename(path))

    def clear_form(self):
        for var in self.vars.values():
            var.set("")
        self.desc_text.delete("1.0", "end")
        self.name_entry.focus()

    def add_record(self):
        values = {k: v.get().strip() for k, v in self.vars.items()}
        description = self.desc_text.get("1.0", "end-1c").strip().replace("\n", " ")

        missing = [k for k, v in values.items() if not v]
        if missing:
            messagebox.showwarning("Ελλιπή στοιχεία", "Συμπλήρωσε: " + ", ".join(missing))
            return

        self.records.append(
            [
                datetime.now().strftime("%d/%m/%Y %H:%M"),
                values["ΟΝΟΜΑΤΕΠΩΝΥΜΟ"],
                values["ΑΠΟ"],
                values["ΣΕ"],
                values["ΟΝΟΜΑ ΑΡΧΕΙΟΥ"],
                description,
            ]
        )
        self.save_records()
        self.refresh_table()
        self.clear_form()

    # ---------- Πίνακας ----------
    def build_table(self):
        top = ttk.Frame(self)
        top.pack(fill="x", padx=12, pady=(6, 0))
        ttk.Label(top, text="Αναζήτηση:").pack(side="left")
        ttk.Entry(top, textvariable=self.search_var, width=35).pack(side="left", padx=8)
        ttk.Button(top, text="Διαγραφή επιλεγμένης", command=self.delete_selected).pack(side="right")
        self.count_label = ttk.Label(top, text="")
        self.count_label.pack(side="right", padx=12)

        frame = ttk.Frame(self)
        frame.pack(fill="both", expand=True, padx=12, pady=12)

        self.tree = ttk.Treeview(frame, columns=COLUMNS, show="headings", selectmode="browse")
        widths = {"ΗΜΕΡΟΜΗΝΙΑ": 120, "ΟΝΟΜΑΤΕΠΩΝΥΜΟ": 160, "ΑΠΟ": 110, "ΣΕ": 110,
                  "ΟΝΟΜΑ ΑΡΧΕΙΟΥ": 170, "ΠΕΡΙΓΡΑΦΗ": 280}
        for col in COLUMNS:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=widths[col], anchor="w")

        vsb = ttk.Scrollbar(frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        self.tree.bind("<Delete>", lambda e: self.delete_selected())

    def refresh_table(self):
        self.tree.delete(*self.tree.get_children())
        query = self.search_var.get().strip().lower()
        shown = 0
        # τελευταίες καταχωρήσεις πρώτα · iid = θέση της εγγραφής στη λίστα
        for i in reversed(range(len(self.records))):
            row = self.records[i]
            if query and query not in " ".join(row).lower():
                continue
            self.tree.insert("", "end", iid=str(i), values=row)
            shown += 1
        self.count_label.config(text=f"{shown} / {len(self.records)} εγγραφές")

    def delete_selected(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showinfo("Διαγραφή", "Διάλεξε πρώτα μια εγγραφή από τον πίνακα.")
            return
        if messagebox.askyesno("Διαγραφή", "Να διαγραφεί η επιλεγμένη εγγραφή;"):
            del self.records[int(selected[0])]
            self.save_records()
            self.refresh_table()


if __name__ == "__main__":
    App().mainloop()
