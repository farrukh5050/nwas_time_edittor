"""Drag-and-drop front end for building Modified_NWAS_File.xlsx.

Drop the NWAS journeys export and the ghost bookings export onto the window (or
browse for them). Filenames vary, so each file is identified by its columns
rather than its name.

Run with:  python gui.py
"""

import os
import queue
import threading
import traceback
from typing import Any

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinterdnd2 import DND_FILES, TkinterDnD

import nwas_pipeline

try:
    DND_AVAILABLE = True
except ImportError:  # browse-only fallback
    DND_FILES = ""
    TkinterDnD = None
    DND_AVAILABLE = False

DEFAULT_OUTPUT_NAME = "Modified_NWAS_File.xlsx"
FILE_TYPES = [
    ("Exports", "*.csv *.xlsx *.xls"),
    ("CSV files", "*.csv"),
    ("Excel files", "*.xlsx *.xls"),
    ("All files", "*.*"),
]

SLOTS = [
    ("nwas", "NWAS journeys", "ID, Journey Time, Category Text..."),
    ("ghost", "Ghost bookings", "Your Reference 1, Completed at Time..."),
]


class App(ttk.Frame):
    def __init__(self, master):
        super().__init__(master, padding=16)
        self.grid(sticky="nsew")
        master.columnconfigure(0, weight=1)
        master.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)

        self.paths = {"nwas": None, "ghost": None}
        self.slot_labels = {}
        self.output_path = tk.StringVar(value=os.path.abspath(DEFAULT_OUTPUT_NAME))
        self.status = tk.StringVar(value="Waiting for both files.")
        self.queue = queue.Queue()
        self.worker = None

        self._build()
        self._refresh()

    # ---------------------------------------------------------------- layout
    def _build(self):
        row = 0
        ttk.Label(self, text="NWAS Time Editor", font=("Segoe UI", 15, "bold")).grid(
            row=row, column=0, sticky="w"
        )
        row += 1

        if DND_AVAILABLE:
            subtitle = "Drop both exports below, in any order."
        else:
            subtitle = "Drag-and-drop unavailable (tkinterdnd2 missing) - use Browse."
        ttk.Label(self, text=subtitle, foreground="#555").grid(
            row=row, column=0, sticky="w", pady=(2, 12)
        )
        row += 1

        self.drop_zone = tk.Label(
            self,
            text="Drop the two files here",
            relief="solid",
            borderwidth=1,
            padx=20,
            pady=26,
            background="#f4f4f4",
            foreground="#444",
        )
        self.drop_zone.grid(row=row, column=0, sticky="ew")
        row += 1

        ttk.Button(self, text="Browse for both files...", command=self.browse).grid(
            row=row, column=0, sticky="w", pady=(8, 14)
        )
        row += 1

        for key, title, hint in SLOTS:
            box = ttk.LabelFrame(self, text=title, padding=8)
            box.grid(row=row, column=0, sticky="ew", pady=(0, 8))
            box.columnconfigure(0, weight=1)
            label = ttk.Label(box, text="not loaded - expects " + hint, foreground="#888")
            label.grid(row=0, column=0, sticky="w")
            ttk.Button(
                box, text="Browse...", width=10,
                command=lambda k=key: self.browse_slot(k)
            ).grid(row=0, column=1)
            self.slot_labels[key] = label
            row += 1

        out = ttk.LabelFrame(self, text="Save output as", padding=8)
        out.grid(row=row, column=0, sticky="ew", pady=(4, 12))
        out.columnconfigure(0, weight=1)
        ttk.Entry(out, textvariable=self.output_path).grid(row=0, column=0, sticky="ew")
        ttk.Button(out, text="Change...", command=self.choose_output).grid(
            row=0, column=1, padx=(6, 0)
        )
        row += 1

        self.run_button = ttk.Button(self, text="Create workbook", command=self.run)
        self.run_button.grid(row=row, column=0, sticky="w")
        row += 1

        self.progress = ttk.Progressbar(self, mode="indeterminate")
        self.progress.grid(row=row, column=0, sticky="ew", pady=(12, 4))
        row += 1

        ttk.Label(self, textvariable=self.status, foreground="#333").grid(
            row=row, column=0, sticky="w"
        )

        if DND_AVAILABLE:
            # tkinterdnd2 bolts drop_target_register/dnd_bind onto
            # tkinter.BaseWidget when it is imported, so they exist at runtime
            # but not on the widget classes themselves.
            for widget in (self.drop_zone, self):
                target: Any = widget
                target.drop_target_register(DND_FILES)
                target.dnd_bind("<<Drop>>", self.on_drop)

    # ------------------------------------------------------------ file input
    def on_drop(self, event):
        # splitlist handles the {braced paths} tkdnd uses for names with spaces.
        self.add_files(self.tk.splitlist(event.data))
        return event.action

    def browse(self):
        chosen = filedialog.askopenfilenames(
            title="Select the NWAS and ghost exports", filetypes=FILE_TYPES
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
                kind = nwas_pipeline.identify(path)
            except nwas_pipeline.PipelineError as exc:
                messagebox.showerror("Cannot read file", str(exc))
                continue

            if kind is None:
                missing = nwas_pipeline.missing_columns(path)
                messagebox.showerror(
                    "Unrecognised file",
                    "'%s' is not one of the two exports.\n\n"
                    "Closest match is missing these columns:\n  %s"
                    % (os.path.basename(path), "\n  ".join(missing)),
                )
                continue

            self._accept(kind, path)

        self._refresh()

    def browse_slot(self, kind):
        """Pick a file for one slot, and check it has that slot's columns."""
        titles = dict((k, t) for k, t, _ in SLOTS)
        path = filedialog.askopenfilename(
            title="Select the %s export" % titles[kind], filetypes=FILE_TYPES
        )
        if not path:
            return

        try:
            missing = nwas_pipeline.missing_for(path, kind)
            actual = nwas_pipeline.identify(path)
        except nwas_pipeline.PipelineError as exc:
            messagebox.showerror("Cannot read file", str(exc))
            return

        if missing:
            hint = ""
            if actual is not None and actual != kind:
                hint = "\n\nThis looks like the %s export - load it in the other slot." % titles[actual]
            messagebox.showerror(
                "Missing columns",
                "'%s' cannot be used as the %s export.\n\n"
                "Missing %d of %d required columns:\n  %s%s"
                % (
                    os.path.basename(path),
                    titles[kind],
                    len(missing),
                    len(nwas_pipeline.EXPECTED[kind]),
                    "\n  ".join(missing),
                    hint,
                ),
            )
            return

        self._accept(kind, path)
        self._refresh()

    def _accept(self, kind, path):
        """Store a validated file, defaulting the output beside the NWAS export."""
        self.paths[kind] = path
        default = os.path.abspath(DEFAULT_OUTPUT_NAME)
        if kind == "nwas" and self.output_path.get() == default:
            self.output_path.set(
                os.path.join(os.path.dirname(path), DEFAULT_OUTPUT_NAME)
            )

    def choose_output(self):
        current = self.output_path.get()
        chosen = filedialog.asksaveasfilename(
            title="Save workbook as",
            defaultextension=".xlsx",
            initialfile=os.path.basename(current) or DEFAULT_OUTPUT_NAME,
            initialdir=os.path.dirname(current) or os.getcwd(),
            filetypes=[("Excel workbook", "*.xlsx")],
        )
        if chosen:
            self.output_path.set(chosen)

    def _refresh(self):
        for key, title, hint in SLOTS:
            path = self.paths[key]
            label = self.slot_labels[key]
            if path:
                label.configure(text=os.path.basename(path), foreground="#0a6")
            else:
                label.configure(
                    text="not loaded - expects " + hint, foreground="#888"
                )

        ready = all(self.paths.values()) and self.worker is None
        self.run_button.configure(state="normal" if ready else "disabled")
        if self.worker is None:
            waiting = [title for key, title, _ in SLOTS if not self.paths[key]]
            if waiting:
                self.status.set("Waiting for: " + ", ".join(waiting))
            else:
                self.status.set("Ready.")

    # --------------------------------------------------------------- running
    def run(self):
        nwas_path, ghost_path = self.paths["nwas"], self.paths["ghost"]
        output = self.output_path.get().strip()
        if not output:
            messagebox.showerror("No output path", "Choose where to save the workbook.")
            return
        if os.path.exists(output) and not messagebox.askyesno(
            "Overwrite?",
            "'%s' already exists. Replace it?" % os.path.basename(output),
        ):
            return

        self.worker = threading.Thread(
            target=self._work, args=(nwas_path, ghost_path, output), daemon=True
        )
        self.run_button.configure(state="disabled")
        self.progress.start(12)
        self.worker.start()
        self.after(100, self._poll)

    def _work(self, nwas_path, ghost_path, output):
        try:
            rows = nwas_pipeline.build_workbook(
                nwas_path,
                ghost_path,
                output,
                progress=lambda msg: self.queue.put(("progress", msg)),
            )
            self.queue.put(("done", (output,) + rows))
        except Exception:
            self.queue.put(("error", traceback.format_exc()))

    def _poll(self):
        try:
            while True:
                kind, payload = self.queue.get_nowait()
                if kind == "progress":
                    self.status.set(payload)
                elif kind == "done":
                    self._finish(*payload)
                    return
                elif kind == "error":
                    self._fail(payload)
                    return
        except queue.Empty:
            pass
        self.after(100, self._poll)

    def _finish(self, output, nwas_rows, ghost_rows):
        self.progress.stop()
        self.worker = None
        self._refresh()
        self.status.set(
            "Done. %d journeys, %d bookings." % (nwas_rows, ghost_rows)
        )
        if messagebox.askyesno(
            "Workbook created",
            "Saved %d journeys to:\n%s\n\nOpen the containing folder?"
            % (nwas_rows, output),
        ):
            os.startfile(os.path.dirname(os.path.abspath(output)))

    def _fail(self, detail):
        self.progress.stop()
        self.worker = None
        self._refresh()
        self.status.set("Failed - see the message for details.")
        messagebox.showerror("Could not build the workbook", detail)


def main():
    root = TkinterDnD.Tk() if TkinterDnD is not None else tk.Tk()
    root.title("NWAS Time Editor")
    root.minsize(560, 640)
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
