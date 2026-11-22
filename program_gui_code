import csv
import os
import re
import json
import tkinter as tk
import webbrowser
from tkinter import messagebox, ttk, filedialog, simpledialog


class DataInExcelStudio:
    def __init__(self, root):
        self.root = root
        self.root.title("Data In Excel")
        self.root.geometry("1100x750")
        self.root.configure(bg="#f1f5f9")

        self.current_file = None
        self.headers = []
        self.rows = []
        self.config_file = "config.json"
        self.channel_url = "https://www.youtube.com/@cc_cat_codes"

        self.colors = {
            "primary": "#3b82f6",
            "primary_hover": "#2563eb",
            "sidebar": "#0f172a",
            "bg": "#f8fafc",
            "card": "#ffffff",
            "text_main": "#1e293b",
            "text_muted": "#64748b",
            "accent": "#10b981",
            "danger": "#ef4444",
            "link": "#60a5fa"
        }

        self.setup_styles()
        self.create_layout()
        self.load_last_file()
        if not self.load_config().get("subscribed"):
            self.root.withdraw()
            self.root.after(100, self.require_subscription)

    def setup_styles(self):
        style = ttk.Style()
        style.theme_use('clam')
        style.configure("Treeview", background="white", foreground=self.colors["text_main"],
                        rowheight=35, fieldbackground="white", font=("Segoe UI", 10))
        style.map("Treeview", background=[('selected', self.colors["primary"])])
        style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"), background="#f1f5f9",
                        foreground=self.colors["text_muted"], borderwidth=0)

    def create_layout(self):
        self.sidebar = tk.Frame(self.root, bg=self.colors["sidebar"], width=260)
        self.sidebar.pack(side="left", fill="y")
        self.sidebar.pack_propagate(False)

        tk.Label(self.sidebar, text="Data In Excel", bg=self.colors["sidebar"], fg="white",
                 font=("Segoe UI", 22, "bold")).pack(pady=(40, 5))
        tk.Label(self.sidebar, text="By Anant Prakash", bg=self.colors["sidebar"], fg=self.colors["primary"],
                 font=("Segoe UI", 10, "bold")).pack(pady=(0, 40))

        action_frame = tk.Frame(self.sidebar, bg=self.colors["sidebar"])
        action_frame.pack(fill="x", padx=15, pady=10)

        self.create_side_btn(action_frame, "New File", self.create_file).pack(side="left", fill="x", expand=True, padx=2)
        self.create_side_btn(action_frame, "Open File", self.load_file).pack(side="right", fill="x", expand=True, padx=2)

        menu_items = [
            ("Add New / Update", self.focus_entry),
            ("Insert at Position", self.insert_at_pos),
            ("Delete Selection", self.delete_data),
            ("Search / Filter", self.search_data),
            ("Open in Excel", self.open_excel),
            ("Refresh Table", self.refresh_table),
            ("Exit Studio", self.root.quit)
        ]

        for text, func in menu_items:
            btn = self.create_nav_btn(self.sidebar, text, func)
            btn.pack(fill="x", padx=15, pady=4)

        footer = tk.Frame(self.sidebar, bg=self.colors["sidebar"])
        footer.pack(side="bottom", fill="x", pady=20)

        github_url = "https://github.com/yesap4"
        link_label = tk.Label(footer, text=github_url, bg=self.colors["sidebar"],
                              fg=self.colors["text_muted"], font=("Segoe UI", 8, "italic", "underline"),
                              cursor="hand2")
        link_label.pack()

        link_label.bind("<Button-1>", lambda e: webbrowser.open(github_url))
        link_label.bind("<Enter>", lambda e: link_label.config(fg=self.colors["link"]))
        link_label.bind("<Leave>", lambda e: link_label.config(fg=self.colors["text_muted"]))

        self.main_container = tk.Frame(self.root, bg=self.colors["bg"])
        self.main_container.pack(side="right", fill="both", expand=True)

        self.header = tk.Frame(self.main_container, bg="white", height=70)
        self.header.pack(side="top", fill="x")
        self.header.pack_propagate(False)

        self.file_info = tk.Label(self.header, text="No database loaded", bg="white",
                                  fg=self.colors["text_muted"], font=("Segoe UI", 12))
        self.file_info.pack(side="left", padx=25)

        self.content_body = tk.Frame(self.main_container, bg=self.colors["bg"], padx=25, pady=20)
        self.content_body.pack(fill="both", expand=True)

        self.table_frame = tk.Frame(self.content_body, bg="white")
        self.table_frame.pack(fill="both", expand=True)

        self.table = ttk.Treeview(self.table_frame, show="headings")
        self.table.bind("<<TreeviewSelect>>", self.on_row_select)

        scroll = ttk.Scrollbar(self.table_frame, orient="vertical", command=self.table.yview)
        self.table.configure(yscrollcommand=scroll.set)
        self.table.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")

        self.form_panel = tk.LabelFrame(self.content_body, text=" Record Management ", bg="white",
                                        fg=self.colors["text_muted"], font=("Segoe UI", 9, "bold"), pady=15, padx=15)
        self.form_panel.pack(fill="x", pady=(20, 0))
        self.form_container = tk.Frame(self.form_panel, bg="white")
        self.form_container.pack(fill="x")
        self.entry_widgets = {}

    def create_side_btn(self, parent, text, func):
        btn = tk.Button(parent, text=text, command=func, bg=self.colors["primary"], fg="white",
                        font=("Segoe UI", 8, "bold"), relief="flat", pady=8, cursor="hand2")
        btn.bind("<Enter>", lambda e: btn.config(bg=self.colors["primary_hover"]))
        btn.bind("<Leave>", lambda e: btn.config(bg=self.colors["primary"]))
        return btn

    def create_nav_btn(self, parent, text, func):
        btn = tk.Button(parent, text=text, command=func, bg="#1e293b", fg="#cbd5e1",
                        font=("Segoe UI", 10), relief="flat", pady=10, anchor="w", padx=20, cursor="hand2")
        btn.bind("<Enter>", lambda e: btn.config(bg="#334155", fg="white"))
        btn.bind("<Leave>", lambda e: btn.config(bg="#1e293b", fg="#cbd5e1"))
        return btn

    def validate_inputs(self, row_data):
        if not any(str(field).strip() for field in row_data):
            messagebox.showwarning("Validation Error", "All fields are empty. Please enter data.")
            return False

        for i, header in enumerate(self.headers):
            val = str(row_data[i]).strip()
            h_lower = header.lower()
            tokens = set(re.findall(r"[a-z0-9]+", h_lower))

            needs_value = (
                any(key in h_lower for key in ["name", "student", "roll"])
                or "id" in tokens
            )
            if needs_value and not val:
                messagebox.showwarning("Validation Error", f"The field '{header}' is required.")
                return False

            needs_number = (
                any(key in h_lower for key in ["number", "phone", "salary"])
                or bool(tokens & {"age", "no"})
            )
            if needs_number and val:
                clean_val = val.replace('.', '', 1)
                if not clean_val.isdigit():
                    messagebox.showwarning("Validation Error", f"Field '{header}' must contain only numbers.")
                    return False
        return True

    def read_all_rows(self):
        with open(self.current_file, 'r', newline='', encoding='utf-8') as f:
            return list(csv.reader(f))

    def write_all_rows(self, rows):
        with open(self.current_file, 'w', newline='', encoding='utf-8') as f:
            csv.writer(f).writerows(rows)

    def load_config(self):
        try:
            with open(self.config_file, 'r') as f:
                return json.load(f)
        except:
            return {}

    def save_config(self, **updates):
        config = self.load_config()
        config.update(updates)
        with open(self.config_file, 'w') as f:
            json.dump(config, f)

    def load_last_file(self):
        path = self.load_config().get("last_file")
        if path and os.path.exists(path):
            self.current_file = path
            self.refresh_table()

    def refresh_table(self):
        if not self.current_file:
            return
        if not os.path.exists(self.current_file):
            messagebox.showerror("File Not Found", f"{self.current_file} does not exist.")
            return
        for i in self.table.get_children():
            self.table.delete(i)
        try:
            with open(self.current_file, 'r', newline='', encoding='utf-8') as f:
                reader = csv.reader(f)
                self.headers = next(reader, [])
                if not self.headers:
                    messagebox.showwarning("Empty File", "This file has no header row.")
                    return
                self.table["columns"] = self.headers
                for h in self.headers:
                    self.table.heading(h, text=h.upper())
                    self.table.column(h, anchor="w", width=120)
                self.rows = list(reader)
                for idx, row in enumerate(self.rows):
                    self.table.insert("", "end", iid=str(idx), values=row)
            self.file_info.config(
                text=f"Active: {os.path.basename(self.current_file)}  |  {len(self.rows)} rows")
            self.rebuild_form()
            self.save_config(last_file=self.current_file)
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def rebuild_form(self):
        for widget in self.form_container.winfo_children():
            widget.destroy()
        self.entry_widgets = {}
        for i, h in enumerate(self.headers):
            frame = tk.Frame(self.form_container, bg="white")
            frame.grid(row=0, column=i, padx=5, sticky="ew")
            tk.Label(frame, text=h, bg="white", font=("Segoe UI", 8)).pack(anchor="w")
            ent = tk.Entry(frame, relief="solid", borderwidth=1)
            ent.pack(fill="x")
            self.entry_widgets[h] = ent
        tk.Button(self.form_container, text="Save / Add Last", command=self.save_record,
                  bg=self.colors["accent"], fg="white", relief="flat", padx=10).grid(row=0, column=len(self.headers), padx=10)
        tk.Button(self.form_container, text="Clear", command=self.clear_entries,
                  bg=self.colors["text_muted"], fg="white", relief="flat", padx=10).grid(row=0, column=len(self.headers) + 1, padx=4)

    def save_record(self):
        if not self.current_file:
            return
        new_row = [self.entry_widgets[h].get().strip() for h in self.headers]

        if not self.validate_inputs(new_row):
            return

        all_rows = self.read_all_rows()

        selected = self.table.selection()
        if selected:
            all_rows[int(selected[0]) + 1] = new_row
        else:
            all_rows.append(new_row)

        self.write_all_rows(all_rows)
        self.refresh_table()
        self.clear_entries()

    def clear_entries(self):
        for ent in self.entry_widgets.values():
            ent.delete(0, tk.END)
        self.table.selection_remove(self.table.selection())

    def insert_at_pos(self):
        if not self.current_file:
            return
        new_row = [self.entry_widgets[h].get().strip() for h in self.headers]

        if not self.validate_inputs(new_row):
            return

        reader = self.read_all_rows()
        header, data = reader[0], reader[1:]

        pos = simpledialog.askinteger("Insert", f"Enter row number (1 to {len(data) + 1}):")
        if pos is None:
            return
        if pos < 1 or pos > len(data) + 1:
            messagebox.showwarning("Invalid Row", f"Row number must be between 1 and {len(data) + 1}.")
            return

        data.insert(pos - 1, new_row)
        self.write_all_rows([header] + data)
        self.refresh_table()
        self.clear_entries()

    def on_row_select(self, e):
        sel = self.table.selection()
        if not sel:
            return
        row = self.rows[int(sel[0])]
        for i, h in enumerate(self.headers):
            self.entry_widgets[h].delete(0, tk.END)
            self.entry_widgets[h].insert(0, row[i] if i < len(row) else "")

    def create_file(self):
        p = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
        if not p:
            return
        h = simpledialog.askstring("Headers", "Comma separated columns:")
        if not h:
            return
        cols = [x.strip() for x in h.split(",") if x.strip()]
        if not cols:
            messagebox.showwarning("Invalid Headers", "Please enter at least one column name.")
            return
        if len({c.lower() for c in cols}) != len(cols):
            messagebox.showwarning("Invalid Headers", "Column names must be unique.")
            return
        with open(p, 'w', newline='', encoding='utf-8') as f:
            csv.writer(f).writerow(cols)
        self.current_file = p
        self.refresh_table()

    def load_file(self):
        p = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv"), ("All files", "*.*")])
        if p:
            self.current_file = p
            self.refresh_table()

    def delete_data(self):
        if not self.current_file:
            return
        sel = self.table.selection()
        if not sel or not messagebox.askyesno("Confirm", "Delete?"):
            return
        idx = int(sel[0]) + 1
        rows = self.read_all_rows()
        del rows[idx]
        self.write_all_rows(rows)
        self.refresh_table()
        self.clear_entries()

    def search_data(self):
        if not self.current_file:
            return
        q = simpledialog.askstring("Filter", "Search text:")
        self.refresh_table()
        if not q:
            return
        q = q.strip().lower()
        shown = 0
        for i in self.table.get_children():
            if any(q in str(v).lower() for v in self.rows[int(i)]):
                shown += 1
            else:
                self.table.delete(i)
        self.file_info.config(
            text=f"Active: {os.path.basename(self.current_file)}  |  {shown} of {len(self.rows)} rows match")

    def open_excel(self):
        if self.current_file:
            os.startfile(self.current_file)

    def focus_entry(self):
        if self.entry_widgets:
            list(self.entry_widgets.values())[0].focus_set()

    def require_subscription(self):
        dlg = tk.Toplevel(self.root)
        dlg.title("Subscription Required")
        dlg.configure(bg="white")
        dlg.resizable(False, False)

        w, h = 460, 330
        x = (dlg.winfo_screenwidth() - w) // 2
        y = (dlg.winfo_screenheight() - h) // 2
        dlg.geometry(f"{w}x{h}+{x}+{y}")

        state = {"confirmed": False, "started": False}

        tk.Label(dlg, text="Subscribe To Continue", bg="white", fg=self.colors["text_main"],
                 font=("Segoe UI", 16, "bold")).pack(pady=(28, 8))
        tk.Label(dlg, text="Subscribing to my YouTube channel @cc_cat_codes\nis required to use Data In Excel.",
                 bg="white", fg=self.colors["text_muted"], font=("Segoe UI", 10),
                 justify="center").pack(pady=(0, 20))

        confirm_btn = tk.Button(dlg, text="I've Subscribed", bg=self.colors["text_muted"], fg="white",
                                disabledforeground="white", font=("Segoe UI", 10, "bold"),
                                relief="flat", pady=8, padx=34, state="disabled")

        def countdown(n):
            if n > 0:
                confirm_btn.config(text=f"I've Subscribed ({n})")
                dlg.after(1000, lambda: countdown(n - 1))
            else:
                confirm_btn.config(text="I've Subscribed", state="normal",
                                   bg=self.colors["accent"], cursor="hand2")

        def open_channel():
            webbrowser.open(self.channel_url)
            if not state["started"]:
                state["started"] = True
                countdown(5)

        def confirm_subscribed():
            self.save_config(subscribed=True)
            state["confirmed"] = True
            dlg.destroy()

        confirm_btn.config(command=confirm_subscribed)

        subscribe_btn = self.create_side_btn(dlg, "Subscribe on YouTube", open_channel)
        subscribe_btn.config(font=("Segoe UI", 10, "bold"), padx=20)
        subscribe_btn.pack(pady=4)
        confirm_btn.pack(pady=4)

        tk.Button(dlg, text="Exit", command=dlg.destroy, bg="white", fg=self.colors["text_muted"],
                  font=("Segoe UI", 9), relief="flat", cursor="hand2", bd=0).pack(pady=(10, 0))

        dlg.protocol("WM_DELETE_WINDOW", dlg.destroy)
        dlg.grab_set()
        dlg.focus_force()
        self.root.wait_window(dlg)

        if state["confirmed"]:
            self.root.deiconify()
        else:
            self.root.destroy()


if __name__ == "__main__":
    root = tk.Tk()
    app = DataInExcelStudio(root)
    root.mainloop()
