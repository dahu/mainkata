#!/usr/bin/env python3
from __future__ import annotations

import os
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from generator import generate_from_inputs, resolve_output_path


class VocabPptxGui(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Term-definition PPTX generator")
        self.geometry("760x560")
        self.minsize(720, 500)

        self.csv_path_var = tk.StringVar()
        self.output_var = tk.StringVar()
        self.sets_var = tk.StringVar(value="6")
        self.set_size_var = tk.StringVar(value="10")
        self.seed_var = tk.StringVar(value="42")
        self.primary_side_var = tk.StringVar(value="term")
        self.show_alternate_var = tk.BooleanVar(value=True)
        self.export_selected_terms_var = tk.BooleanVar(value=False)
        self.status_var = tk.StringVar(value="Choose a CSV file to begin.")

        self._build_ui()

    def _build_ui(self) -> None:
        root = ttk.Frame(self, padding=16)
        root.pack(fill="both", expand=True)
        root.columnconfigure(1, weight=1)

        ttk.Label(root, text="CSV file").grid(
            row=0, column=0, sticky="w", padx=(0, 10), pady=8
        )
        ttk.Entry(root, textvariable=self.csv_path_var).grid(
            row=0, column=1, sticky="ew", pady=8
        )
        ttk.Button(root, text="Browse…", command=self.choose_csv).grid(
            row=0, column=2, sticky="ew", pady=8
        )

        ttk.Label(root, text="Output PPTX").grid(
            row=1, column=0, sticky="w", padx=(0, 10), pady=8
        )
        ttk.Entry(root, textvariable=self.output_var).grid(
            row=1, column=1, sticky="ew", pady=8
        )
        ttk.Button(root, text="Save as…", command=self.choose_output).grid(
            row=1, column=2, sticky="ew", pady=8
        )

        ttk.Label(root, text="Sets").grid(
            row=2, column=0, sticky="w", padx=(0, 10), pady=8
        )
        ttk.Spinbox(root, from_=1, to=999, textvariable=self.sets_var, width=10).grid(
            row=2, column=1, sticky="w", pady=8
        )

        ttk.Label(root, text="Set size").grid(
            row=3, column=0, sticky="w", padx=(0, 10), pady=8
        )
        ttk.Spinbox(
            root, from_=1, to=999, textvariable=self.set_size_var, width=10
        ).grid(row=3, column=1, sticky="w", pady=8)

        ttk.Label(root, text="Seed").grid(
            row=4, column=0, sticky="w", padx=(0, 10), pady=8
        )
        ttk.Entry(root, textvariable=self.seed_var, width=12).grid(
            row=4, column=1, sticky="w", pady=8
        )

        ttk.Label(root, text="Large text shows").grid(
            row=5, column=0, sticky="w", padx=(0, 10), pady=8
        )
        primary_frame = ttk.Frame(root)
        primary_frame.grid(row=5, column=1, columnspan=2, sticky="w", pady=8)
        ttk.Radiobutton(
            primary_frame, text="Term", value="term", variable=self.primary_side_var
        ).pack(side="left", padx=(0, 12))
        ttk.Radiobutton(
            primary_frame,
            text="Definition",
            value="definition",
            variable=self.primary_side_var,
        ).pack(side="left")

        ttk.Checkbutton(
            root,
            text="Show alternate side in smaller text",
            variable=self.show_alternate_var,
        ).grid(row=6, column=0, columnspan=3, sticky="w", pady=6)

        ttk.Checkbutton(
            root,
            text="Generate selected_terms CSV",
            variable=self.export_selected_terms_var,
        ).grid(row=7, column=0, columnspan=3, sticky="w", pady=6)

        help_text = (
            "CSV must contain Term and Definition headers.\n"
            "Header matching is case-insensitive, and one CSV file is processed at a time.\n"
            + "The file must also contain enough unique rows for the selected set size."
        )
        ttk.Label(root, text=help_text, justify="left").grid(
            row=8, column=0, columnspan=3, sticky="w", pady=(12, 8)
        )

        ttk.Button(root, text="Generate PPTX", command=self.generate).grid(
            row=9, column=0, columnspan=3, sticky="ew", pady=(8, 12)
        )

        ttk.Label(root, text="Status").grid(row=10, column=0, sticky="nw", pady=(4, 6))
        status_frame = ttk.Frame(root)
        status_frame.grid(row=10, column=1, columnspan=2, sticky="nsew", pady=(4, 6))
        root.rowconfigure(10, weight=1)
        status_frame.columnconfigure(0, weight=1)
        status_frame.rowconfigure(1, weight=1)

        ttk.Label(status_frame, textvariable=self.status_var).grid(
            row=0, column=0, sticky="w", pady=(0, 6)
        )
        self.log = tk.Text(status_frame, height=10, wrap="word", state="disabled")
        self.log.grid(row=1, column=0, sticky="nsew")

    def _require_int(self, value: str, field_name: str) -> int:
        text = value.strip()
        if not text:
            raise ValueError(f"{field_name} is required.")
        try:
            return int(text)
        except ValueError as exc:
            raise ValueError(f"{field_name} must be a whole number.") from exc

    def _validate_inputs(
        self,
    ) -> tuple[str, str | None, int, int, int, str, bool, bool]:
        csv_file = self.csv_path_var.get().strip()
        if not csv_file:
            raise ValueError("Please choose a CSV file.")

        output = self.output_var.get().strip() or None
        set_count = self._require_int(self.sets_var.get(), "Sets")
        set_size = self._require_int(self.set_size_var.get(), "Set size")
        seed = self._require_int(self.seed_var.get(), "Seed")
        primary_side = self.primary_side_var.get().strip()
        show_alternate = self.show_alternate_var.get()
        export_selected_terms = self.export_selected_terms_var.get()

        if set_count < 1:
            raise ValueError("Sets must be at least 1.")
        if set_size < 1:
            raise ValueError("Set size must be at least 1.")
        if primary_side not in {"term", "definition"}:
            raise ValueError("Large text shows must be either Term or Definition.")

        return (
            csv_file,
            output,
            set_count,
            set_size,
            seed,
            primary_side,
            show_alternate,
            export_selected_terms,
        )

    def _validate_output_path(self, csv_file: str, output: str | None) -> Path:
        csv_path = Path(csv_file).expanduser().resolve()
        output_path = resolve_output_path(csv_path, output)

        if output_path.exists() and output_path.is_dir():
            raise ValueError(
                f"Output path is a folder, not a .pptx file:\n{output_path}"
            )

        parent = output_path.parent
        if not parent.exists():
            raise ValueError(f"Output folder does not exist:\n{parent}")
        if not os.access(parent, os.W_OK):
            raise PermissionError(f"Output folder is not writable:\n{parent}")

        return output_path

    def choose_csv(self) -> None:
        filename = filedialog.askopenfilename(
            title="Choose CSV file",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
        )
        if not filename:
            return
        self.csv_path_var.set(filename)
        if not self.output_var.get().strip():
            output_path = resolve_output_path(Path(filename))
            self.output_var.set(str(output_path))
        self.status_var.set("CSV file selected.")
        self._append_log(f"CSV: {filename}")

    def choose_output(self) -> None:
        initial = self.output_var.get().strip()
        filename = filedialog.asksaveasfilename(
            title="Save PPTX as",
            defaultextension=".pptx",
            initialfile=Path(initial).name if initial else "",
            initialdir=str(Path(initial).parent) if initial else "",
            filetypes=[("PowerPoint files", "*.pptx")],
        )
        if not filename:
            return
        self.output_var.set(filename)
        self.status_var.set("Output file selected.")
        self._append_log(f"Output: {filename}")

    def generate(self) -> None:
        self.status_var.set("Validating inputs...")
        try:
            (
                csv_file,
                output,
                set_count,
                set_size,
                seed,
                primary_side,
                show_alternate,
                export_selected_terms,
            ) = self._validate_inputs()

            output_path = self._validate_output_path(csv_file, output)

            files_to_overwrite = []
            if output_path.exists():
                files_to_overwrite.append(output_path)

            if export_selected_terms:
                csv_out_path = output_path.with_name(
                    output_path.stem + "_selected_terms.csv"
                )
                if csv_out_path.exists():
                    files_to_overwrite.append(csv_out_path)

            if files_to_overwrite:
                message = "The following file(s) already exist:\n\n"
                message += "\n".join(str(p) for p in files_to_overwrite)
                message += "\n\nDo you want to overwrite them?"
                ok = messagebox.askyesno(
                    "Overwrite existing files?", message, parent=self
                )
                if not ok:
                    self.status_var.set("Generation cancelled.")
                    self._append_log("Cancelled: output file already exists.")
                    return

            self.status_var.set("Generating PPTX...")
            self._append_log(f"Generating: {output_path}")

            pptx_path, csv_out = generate_from_inputs(
                csv_file=csv_file,
                output=str(output_path),
                set_count=set_count,
                set_size=set_size,
                seed=seed,
                primary_side=primary_side,
                show_alternate=show_alternate,
                export_selected_terms=export_selected_terms,
            )

            self.status_var.set("Generation complete.")
            self._append_log(f"Created: {pptx_path}")
            if csv_out:
                self._append_log(f"Created: {csv_out}")
                done_message = f"Created:\n{pptx_path}\n\nCreated:\n{csv_out}"
            else:
                done_message = f"Created:\n{pptx_path}"
            messagebox.showinfo("Done", done_message)

        except ValueError as exc:
            self.status_var.set("Please fix the input values.")
            self._append_log(f"Validation error: {exc}")
            messagebox.showerror("Check your inputs", str(exc), parent=self)

        except PermissionError as exc:
            self.status_var.set("Cannot write to the selected output location.")
            self._append_log(f"Permission error: {exc}")
            messagebox.showerror(
                "Output location not writable",
                str(exc),
                parent=self,
            )

        except FileNotFoundError as exc:
            self.status_var.set("Input file not found.")
            self._append_log(f"File error: {exc}")
            messagebox.showerror(
                "File not found",
                str(exc),
                parent=self,
            )

        except Exception as exc:
            self.status_var.set("Generation failed.")
            self._append_log(f"Unexpected error: {exc}")
            messagebox.showerror(
                "Generation failed",
                "Something went wrong while creating the PowerPoint.\n"
                "Please check the CSV file and output location, then try again.",
                parent=self,
            )

    def _append_log(self, message: str) -> None:
        self.log.configure(state="normal")
        self.log.insert("end", message + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")


def main() -> None:
    app = VocabPptxGui()
    app.mainloop()


if __name__ == "__main__":
    main()
