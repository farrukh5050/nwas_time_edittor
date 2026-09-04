"""NWAS Time Editor - the application's only entry point.

Run with:  pythonw gui.py

Tab 1 builds Modified_NWAS_File.xlsx from the two raw exports. Tab 2 prepares
the portal upload file and drives the NWAS portal.

Long jobs run on a worker thread; anything the underlying modules print is
captured and shown in the log at the bottom.
"""

import contextlib
import io
import os
import queue
import threading
import traceback
from typing import Any

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import combine_nwas_ghost
import nwas_pipeline
import portal

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD

    DND_AVAILABLE = True
except ImportError:  # browse-only fallback
    DND_FILES = ""
    TkinterDnD = None
    DND_AVAILABLE = False

EXPORT_TYPES = [
    ("Exports", "*.csv *.xlsx *.xls"),
    ("CSV files", "*.csv"),
    ("Excel files", "*.xlsx *.xls"),
    ("All files", "*.*"),
]
WORKBOOK_TYPES = [("Excel workbook", "*.xlsx *.xls"), ("All files", "*.*")]

SLOTS = [
    ("nwas", "NWAS journeys", "ID, Journey Time, Category Text..."),
    ("ghost", "Ghost bookings", "Your Reference 1, Completed at Time..."),
]

BLANK = os.linesep * 2


class _LogStream(io.TextIOBase):
    """Sends everything written to it to the UI's log queue."""

    def __init__(self, emit):
        self.emit = emit

    def write(self, text):
        for line in text.splitlines():
            if line.strip():
                self.emit(line)
        return len(text)


class App(ttk.Frame):
    def __init__(self, master):
        super().__init__(master, padding=12)
        self.grid(sticky="nsew")
        master.columnconfigure(0, weight=1)
        master.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

        self.paths = {"nwas": None, "ghost": None}
        self.slot_labels = {}
        self.output_path = tk.StringVar(
            value=os.path.abspath(nwas_pipeline.MODIFIED_FILE)
        )
        self.status = tk.StringVar(value="Ready.")
        self.upload_status = tk.StringVar()
        self.month = tk.StringVar()
        self.year = tk.StringVar()
        self.queue = queue.Queue()
        self.worker = None
        self.buttons = []

        self._build()
        self._refresh()

    # ---------------------------------------------------------------- layout
    def _build(self):
        ttk.Label(self, text="NWAS Time Editor", font=("Segoe UI", 15, "bold")).grid(
            row=0, column=0, sticky="w", pady=(0, 8)
        )

        book = ttk.Notebook(self)
        book.grid(row=1, column=0, sticky="nsew")
        build_tab = ttk.Frame(book, padding=12)
        portal_tab = ttk.Frame(book, padding=12)
        book.add(build_tab, text="  Build workbook  ")
        book.add(portal_tab, text="  Portal upload  ")
        self._build_tab(build_tab)
        self._portal_tab(portal_tab)

        log_box = ttk.LabelFrame(self, text="Log", padding=6)
        log_box.grid(row=2, column=0, sticky="nsew", pady=(10, 6))
        log_box.columnconfigure(0, weight=1)
        log_box.rowconfigure(0, weight=1)
        self.log = tk.Text(
            log_box, height=9, wrap="none", state="disabled",
            background="#fbfbfb", relief="flat",
        )
        self.log.grid(row=0, column=0, sticky="nsew")
        bar = ttk.Scrollbar(log_box, orient="vertical", command=self.log.yview)
        bar.grid(row=0, column=1, sticky="ns")
        self.log.configure(yscrollcommand=bar.set)

        self.progress = ttk.Progressbar(self, mode="indeterminate")
        self.progress.grid(row=3, column=0, sticky="ew")
        ttk.Label(self, textvariable=self.status, foreground="#333").grid(
            row=4, column=0, sticky="w", pady=(4, 0)
        )

    def _button(self, parent, text, command, **grid):
        button = ttk.Button(parent, text=text, command=command)
        button.grid(**grid)
        self.buttons.append(button)
        return button

    def _build_tab(self, tab):
        tab.columnconfigure(0, weight=1)
        row = 0

        if DND_AVAILABLE:
            hint = "Drop both exports below, in any order."
        else:
            hint = "Drag-and-drop unavailable (tkinterdnd2 missing) - use Browse."
        ttk.Label(tab, text=hint, foreground="#555").grid(row=row, column=0, sticky="w")
        row += 1

        self.drop_zone = tk.Label(
            tab, text="Drop the two exports here", relief="solid", borderwidth=1,
            padx=20, pady=20, background="#f4f4f4", foreground="#444",
        )
        self.drop_zone.grid(row=row, column=0, sticky="ew", pady=(6, 8))
        row += 1

        self._button(tab, "Browse for both files...", self.browse_both,
                     row=row, column=0, sticky="w", pady=(0, 10))
        row += 1

        for key, title, hint in SLOTS:
            box = ttk.LabelFrame(tab, text=title, padding=8)
            box.grid(row=row, column=0, sticky="ew", pady=(0, 6))
            box.columnconfigure(0, weight=1)
            label = ttk.Label(
                box, text="not loaded - expects " + hint, foreground="#888"
            )
            label.grid(row=0, column=0, sticky="w")
            self._button(box, "Browse...", lambda k=key: self.browse_slot(k),
                         row=0, column=1)
            self.slot_labels[key] = label
            row += 1

        out = ttk.LabelFrame(tab, text="Save workbook as", padding=8)
        out.grid(row=row, column=0, sticky="ew", pady=(6, 10))
        out.columnconfigure(0, weight=1)
        ttk.Entry(out, textvariable=self.output_path).grid(row=0, column=0, sticky="ew")
        self._button(out, "Change...", self.choose_output, row=0, column=1, padx=(6, 0))
        row += 1

        self._button(tab, "Create workbook from the two exports",
                     self.create_from_exports, row=row, column=0, sticky="w")

    def _portal_tab(self, tab):
        tab.columnconfigure(0, weight=1)

        step1 = ttk.LabelFrame(tab, text="1. Build the portal upload file", padding=8)
        step1.grid(row=0, column=0, sticky="ew")
        step1.columnconfigure(0, weight=1)
        self._button(step1, "From a ghost export (CSV or Excel)...",
                     self.upload_from_ghost_export, row=0, column=0, sticky="w")
        self._button(step1, "From a built Modified_NWAS workbook...",
                     self.upload_from_modified,
                     row=1, column=0, sticky="w", pady=(4, 6))
        ttk.Label(step1, textvariable=self.upload_status, foreground="#555").grid(
            row=2, column=0, sticky="w"
        )

        step2 = ttk.LabelFrame(tab, text="2. Open the portal", padding=8)
        step2.grid(row=1, column=0, sticky="ew", pady=(10, 0))
        self._button(step2, "Open Chrome and log in", self.open_portal,
                     row=0, column=0, sticky="w")
        ttk.Label(
            step2,
            text="Then load the dates you want in Chrome before using step 3.",
            foreground="#555",
        ).grid(row=1, column=0, sticky="w", pady=(4, 0))

        step3 = ttk.LabelFrame(
            tab, text="3. Update the page showing in Chrome", padding=8
        )
        step3.grid(row=2, column=0, sticky="ew", pady=(10, 0))
        self._button(step3, "Update times for the loaded jobs", self.update_times,
                     row=0, column=0, sticky="w")

        step4 = ttk.LabelFrame(
            tab, text="4. Or run a whole month unattended", padding=8
        )
        step4.grid(row=3, column=0, sticky="ew", pady=(10, 0))
        entries = ttk.Frame(step4)
        entries.grid(row=0, column=0, sticky="w")
        ttk.Label(entries, text="Month").grid(row=0, column=0)
        ttk.Entry(entries, textvariable=self.month, width=5).grid(
            row=0, column=1, padx=(4, 12)
        )
        ttk.Label(entries, text="Year").grid(row=0, column=2)
        ttk.Entry(entries, textvariable=self.year, width=7).grid(
            row=0, column=3, padx=(4, 0)
        )
        self._button(step4, "Auto-run all %d call signs" % len(portal.CALL_SIGNS),
                     self.run_month, row=1, column=0, sticky="w", pady=(8, 0))

        if DND_AVAILABLE:
            for widget in (self.drop_zone, self):
                # tkinterdnd2 bolts these onto tkinter.BaseWidget at import time,
                # so they exist at runtime but not on the widget classes.
                target: Any = widget
                target.drop_target_register(DND_FILES)
                target.dnd_bind("<<Drop>>", self.on_drop)

    # ------------------------------------------------------------ file input
    def on_drop(self, event):
        # splitlist handles the {braced paths} tkdnd uses for names with spaces.
        self.add_files(self.tk.splitlist(event.data))
        return event.action

    def browse_both(self):
        chosen = filedialog.askopenfilenames(
            title="Select the NWAS and ghost exports", filetypes=EXPORT_TYPES
        )
        if chosen:
            self.add_files(chosen)

    def add_files(self, paths):
        paths = [p for p in paths if p]
        if len(paths) > 2:
            messagebox.showwarning(
                "Too many files",
                "Drop at most two files: one NWAS export and one ghost export.",
            )
            return

        for path in paths:
            try:
                report = nwas_pipeline.inspect(path)
            except nwas_pipeline.PipelineError as exc:
                messagebox.showerror("Cannot read file", str(exc))
                continue

            kind = report["kind"]
            if kind is None:
                closest = min(report["missing"].values(), key=len)
                messagebox.showerror(
                    "Unrecognised file",
                    "'%s' is not one of the two exports.%sClosest match is missing: %s"
                    % (os.path.basename(path), BLANK, ", ".join(closest)),
                )
                continue

            self._accept(kind, path)

        self._refresh()

    def browse_slot(self, kind):
        """Pick a file for one slot, and check it has that slot's columns."""
        titles = dict((k, t) for k, t, _ in SLOTS)
        path = filedialog.askopenfilename(
            title="Select the %s export" % titles[kind], filetypes=EXPORT_TYPES
        )
        if not path:
            return

        try:
            report = nwas_pipeline.inspect(path)
        except nwas_pipeline.PipelineError as exc:
            messagebox.showerror("Cannot read file", str(exc))
            return

        missing = report["missing"][kind]
        actual = report["kind"]

        if missing:
            hint = ""
            if actual is not None and actual != kind:
                hint = "%sThis looks like the %s export - use the other slot." % (
                    BLANK,
                    titles[actual],
                )
            messagebox.showerror(
                "Missing columns",
                "'%s' cannot be used as the %s export.%sMissing %d of %d columns: %s%s"
                % (
                    os.path.basename(path),
                    titles[kind],
                    BLANK,
                    len(missing),
                    len(nwas_pipeline.EXPECTED[kind]),
                    ", ".join(missing),
                    hint,
                ),
            )
            return

        self._accept(kind, path)
        self._refresh()

    def _accept(self, kind, path):
        """Store a validated file, defaulting the output beside the NWAS export."""
        self.paths[kind] = path
        default = os.path.abspath(nwas_pipeline.MODIFIED_FILE)
        if kind == "nwas" and self.output_path.get() == default:
            self.output_path.set(
                os.path.join(os.path.dirname(path), nwas_pipeline.MODIFIED_FILE)
            )

    def choose_output(self):
        current = self.output_path.get()
        chosen = filedialog.asksaveasfilename(
            title="Save workbook as",
            defaultextension=".xlsx",
            initialfile=os.path.basename(current) or nwas_pipeline.MODIFIED_FILE,
            initialdir=os.path.dirname(current) or os.getcwd(),
            filetypes=[("Excel workbook", "*.xlsx")],
        )
        if chosen:
            self.output_path.set(chosen)

    # -------------------------------------------------------- tab 1 actions
    def _confirm_output(self):
        output = self.output_path.get().strip()
        if not output:
            messagebox.showerror("No output path", "Choose where to save the workbook.")
            return None
        if os.path.exists(output) and not messagebox.askyesno(
            "Overwrite?", "'%s' already exists. Replace it?" % os.path.basename(output)
        ):
            return None
        return output

    def create_from_exports(self):
        if not all(self.paths.values()):
            messagebox.showinfo("Two files needed", "Load both exports first.")
            return
        output = self._confirm_output()
        if not output:
            return
        nwas_path, ghost_path = self.paths["nwas"], self.paths["ghost"]

        def job(progress):
            frames = nwas_pipeline.build_frames(
                nwas_path, ghost_path, progress=progress
            )
            return combine_nwas_ghost.write_workbook(
                frames[0], frames[1], output, progress=progress
            )

        self._start(job, self._workbook_done(output))

    def _workbook_done(self, output):
        def done(rows):
            nwas_rows, ghost_rows = rows
            self.status.set("Done. %d journeys, %d bookings." % (nwas_rows, ghost_rows))
            if messagebox.askyesno(
                "Workbook created",
                "Saved %d journeys to:%s%s%sOpen the containing folder?"
                % (nwas_rows, os.linesep, output, BLANK),
            ):
                os.startfile(os.path.dirname(os.path.abspath(output)))

        return done

    # -------------------------------------------------------- tab 2 actions
    def _write_upload_file(self, frame):
        target = os.path.abspath(nwas_pipeline.CLEANED_FILE)
        frame.to_excel(target, index=False)
        print("Upload file saved as '%s' (%d rows)." % (target, len(frame)))
        return target

    def upload_from_ghost_export(self):
        path = filedialog.askopenfilename(
            title="Select the ghost export", filetypes=EXPORT_TYPES
        )
        if not path:
            return

        def job(progress):
            frame = nwas_pipeline.clean_ghost_export(path, progress=progress)
            return self._write_upload_file(frame)

        self._start(job, self._upload_done)

    def upload_from_modified(self):
        path = filedialog.askopenfilename(
            title="Select Modified_NWAS_File.xlsx", filetypes=WORKBOOK_TYPES
        )
        if not path:
            return

        def job(progress):
            frame = nwas_pipeline.cleaned_from_modified(path, progress=progress)
            return self._write_upload_file(frame)

        self._start(job, self._upload_done)

    def _upload_done(self, _target):
        self.status.set("Upload file ready.")

    def open_portal(self):
        self._start(lambda progress: portal.open_chrome_and_login(), self._portal_done)

    def update_times(self):
        self._start(lambda progress: portal.update_times(), self._portal_done)

    def run_month(self):
        try:
            month = int(self.month.get().strip())
            year = int(self.year.get().strip())
        except ValueError:
            messagebox.showerror("Invalid date", "Enter a month (1-12) and a year.")
            return
        if not 1 <= month <= 12:
            messagebox.showerror("Invalid month", "Month must be between 1 and 12.")
            return
        if not messagebox.askyesno(
            "Run the whole month?",
            "This works through %d call signs for every day of %02d/%d.%sIt takes a long "
            "while and drives Chrome unattended. Continue?"
            % (len(portal.CALL_SIGNS), month, year, BLANK),
        ):
            return
        self._start(
            lambda progress: portal.run_full_automation(year=year, month=month),
            self._portal_done,
        )

    def _portal_done(self, _result):
        self.status.set("Finished - see the log.")

    # --------------------------------------------------------------- running
    def _start(self, job, on_done):
        """Run job(progress) on a worker thread, capturing whatever it prints."""
        if self.worker is not None:
            messagebox.showinfo("Busy", "Wait for the current job to finish.")
            return

        def work():
            def emit(line):
                self.queue.put(("log", line))

            try:
                with contextlib.redirect_stdout(_LogStream(emit)):
                    result = job(lambda message: self.queue.put(("progress", message)))
                self.queue.put(("done", (on_done, result)))
            except Exception as exc:
                self.queue.put(("error", (exc, traceback.format_exc())))

        self.worker = threading.Thread(target=work, daemon=True)
        self._set_busy(True)
        self.worker.start()
        self.after(100, self._poll)

    def _poll(self):
        try:
            while True:
                kind, payload = self.queue.get_nowait()
                if kind == "progress":
                    self.status.set(payload)
                    self._append_log(payload)
                elif kind == "log":
                    self._append_log(payload)
                elif kind == "done":
                    on_done, result = payload
                    self._set_busy(False)
                    on_done(result)
                    return
                elif kind == "error":
                    exc, detail = payload
                    self._set_busy(False)
                    self._append_log(detail)
                    self.status.set("Failed - see the log.")
                    if isinstance(exc, (nwas_pipeline.PipelineError, RuntimeError)):
                        messagebox.showerror("Cannot continue", str(exc))
                    else:
                        messagebox.showerror("Something went wrong", detail)
                    return
        except queue.Empty:
            pass
        self.after(100, self._poll)

    def _append_log(self, text):
        self.log.configure(state="normal")
        self.log.insert("end", text.rstrip() + os.linesep)
        self.log.see("end")
        self.log.configure(state="disabled")

    def _set_busy(self, busy):
        if busy:
            self.progress.start(12)
        else:
            self.progress.stop()
            self.worker = None
        for button in self.buttons:
            button.configure(state="disabled" if busy else "normal")
        if not busy:
            self._refresh()

    # ---------------------------------------------------------------- status
    def _refresh(self):
        for key, _, hint in SLOTS:
            path = self.paths[key]
            label = self.slot_labels[key]
            if path:
                label.configure(text=os.path.basename(path), foreground="#0a6")
            else:
                label.configure(text="not loaded - expects " + hint, foreground="#888")

        target = os.path.abspath(nwas_pipeline.CLEANED_FILE)
        if os.path.exists(target):
            self.upload_status.set("Current upload file: %s" % target)
        else:
            self.upload_status.set("No upload file yet - build one above.")


def main():
    root = TkinterDnD.Tk() if TkinterDnD is not None else tk.Tk()
    root.title("NWAS Time Editor")
    root.minsize(660, 800)
    app = App(root)

    def on_close():
        if app.worker is not None and not messagebox.askyesno(
            "Job running", "A job is still running. Close anyway?"
        ):
            return
        portal.close_browser()
        root.destroy()

    root.protocol("WM_DELETE_WINDOW", on_close)
    root.mainloop()


if __name__ == "__main__":
    main()
