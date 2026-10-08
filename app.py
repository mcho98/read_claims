#!/usr/bin/env python3
"""Desktop app for read_claims: pick the Excel file, pick a client, fill the claim form.

No terminal needed. Build a double-clickable Windows program with build.bat.
The form address is read from config.json next to the program.
"""
import json
import os
import queue
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import openpyxl

import read_claims as rc
from version import VERSION

DEFAULT_URL = "https://mock-website-theta.vercel.app"  # used when config.json is missing


def app_dir():
    """Folder holding config.json: next to the .exe (Windows), next to the .app (Mac),
    or next to this file when run from source."""
    if not getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(__file__))
    folder = os.path.dirname(sys.executable)
    if sys.platform == "darwin" and ".app/Contents/MacOS" in folder:
        folder = os.path.dirname(os.path.dirname(os.path.dirname(folder)))  # MacOS -> Contents -> .app -> its folder
    return folder


def load_config():
    """Read config.json next to the program; optional keys override the defaults."""
    try:
        with open(os.path.join(app_dir(), "config.json")) as f:
            config = json.load(f)
    except (OSError, ValueError):
        config = {}
    rc.apply_config(config)
    return config


class Worker:
    """Runs browser work on one background thread (Playwright must stay on the
    thread that created it) so the window never freezes."""

    def __init__(self):
        self.jobs = queue.Queue()
        threading.Thread(target=self._run, daemon=True).start()

    def _run(self):
        while True:
            job = self.jobs.get()
            if job is None:
                return
            job()

    def submit(self, job):
        self.jobs.put(job)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"Claim Filler {VERSION}")
        self.geometry("820x600")
        self.config = load_config()
        self.url_var = tk.StringVar(value=self.config.get("url") or DEFAULT_URL)
        self.ws = self.clients = self.date_columns = self.result = None
        self.filler = None
        self.worker = Worker()
        self.messages = queue.Queue()
        self.build_ui()
        self.after(100, self.drain_messages)
        self.protocol("WM_DELETE_WINDOW", self.on_close)

    def build_ui(self):
        pad = {"padx": 8, "pady": 4}
        top = ttk.Frame(self)
        top.pack(fill="x", **pad)
        ttk.Button(top, text="1. Choose Excel file...", command=self.choose_file).pack(side="left")
        self.file_label = ttk.Label(top, text="No file chosen")
        self.file_label.pack(side="left", padx=8)

        mid = ttk.Frame(self)
        mid.pack(fill="both", expand=True, **pad)
        left = ttk.LabelFrame(mid, text="2. Client")
        left.pack(side="left", fill="y")
        self.client_list = tk.Listbox(left, width=24, exportselection=False)
        self.client_list.pack(fill="y", expand=True, padx=4, pady=4)
        self.client_list.bind("<<ListboxSelect>>", self.on_client_selected)

        dates = ttk.LabelFrame(left, text="Dates (YYYYMMDD, optional)")
        dates.pack(fill="x", padx=4, pady=4)
        self.start_var, self.end_var = tk.StringVar(), tk.StringVar()
        for row, (text, var) in enumerate((("From", self.start_var), ("To", self.end_var))):
            ttk.Label(dates, text=text).grid(row=row, column=0, padx=4, pady=2, sticky="w")
            ttk.Entry(dates, textvariable=var, width=10).grid(row=row, column=1, padx=4, pady=2)
            var.trace_add("write", lambda *_: self.on_client_selected())

        right = ttk.LabelFrame(mid, text="Claims to enter")
        right.pack(side="left", fill="both", expand=True, padx=(8, 0))
        self.client_title = ttk.Label(right, text="No client selected", font=("TkDefaultFont", 20, "bold"))
        self.client_title.pack(anchor="w", padx=4, pady=(4, 0))
        cols = ("group", "date", "service", "hours", "cost")
        self.table = ttk.Treeview(right, columns=cols, show="headings", height=10)
        for col, text, width in (("group", "Group", 55), ("date", "Service date", 100), ("service", "Service", 110),
                                 ("hours", "Hours", 60), ("cost", "Total cost ($)", 100)):
            self.table.heading(col, text=text)
            self.table.column(col, width=width, anchor="w" if col in ("date", "service") else "e")
        self.table.pack(fill="both", expand=True, padx=4, pady=4)
        grp = ttk.Frame(right)
        grp.pack(fill="x", padx=4)
        ttk.Label(grp, text="Group to fill (website takes 10 at a time):").pack(side="left")
        self.group_var = tk.StringVar()
        self.group_box = ttk.Combobox(grp, textvariable=self.group_var, state="disabled", width=30)
        self.group_box.pack(side="left", padx=6)
        self.group_box.bind("<<ComboboxSelected>>", lambda _e: self.on_group_selected())
        self.groups = []
        self.summary = ttk.Label(right, text="")
        self.summary.pack(anchor="w", padx=4)
        self.problems = tk.Label(right, text="", fg="#b00020", justify="left", wraplength=480, anchor="w")
        self.problems.pack(fill="x", padx=4)

        site = ttk.Frame(self)
        site.pack(fill="x", **pad)
        ttk.Label(site, text="Claim website address:").pack(side="left")
        ttk.Entry(site, textvariable=self.url_var).pack(side="left", fill="x", expand=True, padx=8)

        bottom = ttk.Frame(self)
        bottom.pack(fill="x", **pad)
        self.open_btn = ttk.Button(bottom, text="3. Open claim website", command=self.open_site)
        self.open_btn.pack(side="left")
        self.fill_btn = ttk.Button(bottom, text="4. Fill form", command=self.fill_form, state="disabled")
        self.fill_btn.pack(side="left", padx=8)
        self.first_only = tk.BooleanVar(value=False)
        ttk.Checkbutton(bottom, text="Test: only the first claim", variable=self.first_only).pack(side="left", padx=8)

        self.log_box = tk.Text(self, height=8, state="disabled")
        self.log_box.pack(fill="x", **pad)
        self.log("Start by choosing the Excel file.")

    # --- thread-safe logging -------------------------------------------------
    def log(self, text):
        self.messages.put(text)

    def ui(self, func):
        """Run func on the window's thread (safe to call from the worker)."""
        self.messages.put(func)

    def drain_messages(self):
        while not self.messages.empty():
            text = self.messages.get()
            if callable(text):
                text()
                continue
            self.log_box.configure(state="normal")
            self.log_box.insert("end", text + "\n")
            self.log_box.see("end")
            self.log_box.configure(state="disabled")
        self.after(100, self.drain_messages)

    # --- steps ---------------------------------------------------------------
    def choose_file(self):
        path = filedialog.askopenfilename(title="Choose the timesheet workbook",
                                          filetypes=[("Excel files", "*.xlsx *.xlsm"), ("All files", "*.*")])
        if not path:
            return
        try:
            wb = openpyxl.load_workbook(path, data_only=True)
            ws = wb.worksheets[0]
            clients = rc.find_clients(ws)
            if not clients:
                raise ValueError(f"No '{rc.CLIENT_LABEL}' labels found on sheet '{ws.title}'.")
            _, date_columns = rc.find_date_columns(ws)
        except (Exception, SystemExit) as e:  # find_date_columns exits with a message
            messagebox.showerror("Could not read the file", str(e))
            return
        self.ws, self.clients, self.date_columns = ws, clients, date_columns
        self.file_label.configure(text=os.path.basename(path))
        self.client_list.delete(0, "end")
        self.client_title.configure(text="No client selected")
        for c in clients:
            self.client_list.insert("end", c["name"])
        self.clear_preview()
        self.log(f"Found {len(clients)} client(s). Pick one from the list.")

    def clear_preview(self):
        self.result = None
        self.groups = []
        self.group_var.set("")
        self.group_box.configure(state="disabled", values=())
        self.table.delete(*self.table.get_children())
        self.summary.configure(text="")
        self.problems.configure(text="")
        self.update_fill_button()

    def on_client_selected(self, _event=None):
        sel = self.client_list.curselection()
        if not sel:
            return
        self.clear_preview()
        self.client_title.configure(text=f"{sel[0] + 1}. {self.clients[sel[0]]['name']}")
        try:
            start, end = rc.parse_day(self.start_var.get()), rc.parse_day(self.end_var.get())
        except ValueError as e:
            self.problems.configure(text=str(e))
            return
        if start and end and start > end:
            self.problems.configure(text="'From' is after 'To'.")
            return
        self.result = rc.read_client(self.ws, self.clients[sel[0]], self.date_columns, start, end)
        self.groups = rc.group_claims(self.result["claims"])
        for n, group in enumerate(self.groups, 1):
            for c in group:
                cost = "" if c["total_cost"] is None else f"{c['total_cost']:.2f}"
                self.table.insert("", "end", values=(n, c["service_date"], c["service"], f"{c['hours']:g}", cost))
        if self.groups:
            labels = [rc.group_label(self.groups, n) for n in range(1, len(self.groups) + 1)]
            self.group_box.configure(values=labels, state="readonly" if len(labels) > 1 else "disabled")
            self.group_var.set(labels[0])
        hours = sum(c["hours"] for c in self.result["claims"])
        cost = sum(c["total_cost"] or 0 for c in self.result["claims"])
        self.summary.configure(text=f"{len(self.result['claims'])} claims, {hours:g} hours, ${cost:,.2f}   "
                                    f"(rate ${self.result['rate']}/hr)")
        self.problems.configure(text="\n".join(self.result["problems"] + self.result["warnings"]))
        self.update_fill_button()

    def open_site(self):
        url = self.url_var.get().strip()
        if not url:
            messagebox.showerror("Claim website", "Paste the claim website address first.")
            return
        if "://" not in url:
            url = "https://" + url
            self.url_var.set(url)
        self.remember_url(url)
        self.open_btn.configure(state="disabled")
        self.log(f"Opening {url} ...")

        def job():
            if self.filler:  # address may have changed: start over with a fresh window
                self.filler.close()
                self.filler = None
                self.ui(self.update_fill_button)
            filler = rc.FormFiller(url, log=self.log)
            filler.screenshot_dir = app_dir()
            try:
                filler.open()
            except Exception as e:
                self.log(f"Could not open the website: {e}")
                self.ui(lambda: self.open_btn.configure(state="normal"))
                return
            self.filler = filler
            self.ui(lambda: self.open_btn.configure(state="normal"))
            self.log("Chrome opened. Pass any security check, log in, and go to the Claim details form, then click 'Fill form'.")
            self.ui(self.update_fill_button)

        self.worker.submit(job)

    def on_group_selected(self):
        self.update_fill_button()

    def selected_group(self):
        """The claims of the group chosen in the drop-down (the only group if there is one)."""
        labels = list(self.group_box.cget("values"))
        if not self.groups or self.group_var.get() not in labels:
            return []
        return self.groups[labels.index(self.group_var.get())]

    def remember_url(self, url):
        """Save the address in config.json so it is pre-filled next time (best effort)."""
        self.config["url"] = url
        try:
            with open(os.path.join(app_dir(), "config.json"), "w") as f:
                json.dump(self.config, f, indent=2)
        except OSError:
            pass

    def update_fill_button(self):
        ready = self.filler and self.result and self.result["claims"] and not self.result["problems"]
        self.fill_btn.configure(state="normal" if ready else "disabled")

    def fill_form(self):
        group = self.selected_group()
        test_only = self.first_only.get()
        claims = group[:1] if test_only else group
        which = f" ({self.group_var.get().split(' (')[0]})" if len(self.groups) > 1 else ""
        if not messagebox.askyesno("Fill form",
                                   f"Enter {len(claims)} claim(s) for {self.result['client']}{which}?\n\n"
                                   + ("This is a test, so 'Predetermine Claim' will NOT be clicked.\n\n" if test_only else
                                      "When every claim is added and the totals match, 'Predetermine Claim' will be clicked.\n\n")
                                   + "Check that the browser is showing the Claim details form."):
            return
        self.fill_btn.configure(state="disabled")

        def job():
            try:
                outcome = self.filler.fill(claims, predetermine=not test_only)
                if outcome["ok"]:
                    self.log(f"Done for {self.result['client']}. Pick another group or client, or close this window.")
                else:
                    self.log("STOPPED. " + outcome["message"])
            except Exception as e:
                self.log(f"Stopped: {e}")
            self.ui(self.update_fill_button)

        self.worker.submit(job)

    def on_close(self):
        def job():
            if self.filler:
                self.filler.close()
        self.worker.submit(job)
        self.worker.submit(None)
        self.after(1000, self.destroy)


if __name__ == "__main__":
    App().mainloop()
