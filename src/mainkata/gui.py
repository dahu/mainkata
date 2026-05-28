#!/usr/bin/env python3
from __future__ import annotations

import os
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from mainkata.services.generator import (generate_from_inputs,
                                         resolve_output_path)


class VocabPptxGui(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Mainkata: Term-Definition PPTX Generator")
        self.geometry("780x940")
        self.minsize(740, 760)

        self.csv_path_var = tk.StringVar()
        self.output_var = tk.StringVar()
        self.sets_var = tk.StringVar(value="6")
        self.set_size_var = tk.StringVar(value="10")
        self.seed_var = tk.StringVar(value="42")
        self.primary_side_var = tk.StringVar(value="term")
        self.show_alternate_var = tk.BooleanVar(value=True)
        self.export_selected_terms_var = tk.BooleanVar(value=False)

        self.background_dir_var = tk.StringVar()
        self.background_mode_var = tk.StringVar(value="cycle")
        self.background_image_number_var = tk.StringVar()
        self.background_cycle_start_var = tk.StringVar()
        self.background_cycle_end_var = tk.StringVar()

        self.title_slide_overlay_transparency_var = tk.StringVar(value="0.22")
        self.vocab_slide_overlay_transparency_var = tk.StringVar(value="0.22")
        self.show_title_card_var = tk.BooleanVar(value=True)
        self.title_card_transparency_var = tk.StringVar(value="0.18")
        self.show_vocab_card_var = tk.BooleanVar(value=True)
        self.vocab_card_transparency_var = tk.StringVar(value="0.18")

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

        ttk.Separator(root).grid(row=8, column=0, columnspan=3, sticky="ew", pady=12)

        ttk.Label(root, text="Background image folder").grid(
            row=9, column=0, sticky="w", padx=(0, 10), pady=8
        )
        ttk.Entry(root, textvariable=self.background_dir_var).grid(
            row=9, column=1, sticky="ew", pady=8
        )
        ttk.Button(root, text="Browse…", command=self.choose_background_dir).grid(
            row=9, column=2, sticky="ew", pady=8
        )

        ttk.Label(root, text="Background mode").grid(
            row=10, column=0, sticky="w", padx=(0, 10), pady=8
        )
        self.background_mode_combo = ttk.Combobox(
            root,
            textvariable=self.background_mode_var,
            values=["cycle", "fixed"],
            state="readonly",
            width=12,
        )
        self.background_mode_combo.grid(row=10, column=1, sticky="w", pady=8)
        self.background_mode_combo.bind(
            "<<ComboboxSelected>>", self._on_background_mode_changed
        )

        ttk.Label(root, text="Fixed image number").grid(
            row=11, column=0, sticky="w", padx=(0, 10), pady=8
        )
        self.fixed_image_spinbox = ttk.Spinbox(
            root,
            from_=1,
            to=9999,
            textvariable=self.background_image_number_var,
            width=10,
        )
        self.fixed_image_spinbox.grid(row=11, column=1, sticky="w", pady=8)

        ttk.Label(root, text="Cycle image range").grid(
            row=12, column=0, sticky="w", padx=(0, 10), pady=8
        )
        cycle_frame = ttk.Frame(root)
        cycle_frame.grid(row=12, column=1, columnspan=2, sticky="w", pady=8)

        self.cycle_start_spinbox = ttk.Spinbox(
            cycle_frame,
            from_=1,
            to=9999,
            textvariable=self.background_cycle_start_var,
            width=8,
        )
        self.cycle_start_spinbox.pack(side="left")

        ttk.Label(cycle_frame, text="to").pack(side="left", padx=8)

        self.cycle_end_spinbox = ttk.Spinbox(
            cycle_frame,
            from_=1,
            to=9999,
            textvariable=self.background_cycle_end_var,
            width=8,
        )
        self.cycle_end_spinbox.pack(side="left")

        ttk.Label(root, text="Title-slide overlay transparency").grid(
            row=13, column=0, sticky="w", padx=(0, 10), pady=8
        )
        self.title_slide_overlay_transparency_entry = ttk.Entry(
            root,
            textvariable=self.title_slide_overlay_transparency_var,
            width=10,
        )
        self.title_slide_overlay_transparency_entry.grid(
            row=13, column=1, sticky="w", pady=8
        )

        ttk.Label(root, text="Vocab-slide overlay transparency").grid(
            row=14, column=0, sticky="w", padx=(0, 10), pady=8
        )
        self.vocab_slide_overlay_transparency_entry = ttk.Entry(
            root,
            textvariable=self.vocab_slide_overlay_transparency_var,
            width=10,
        )
        self.vocab_slide_overlay_transparency_entry.grid(
            row=14, column=1, sticky="w", pady=8
        )

        ttk.Checkbutton(
            root,
            text="Show white card on title slides",
            variable=self.show_title_card_var,
            command=self._on_title_card_toggle,
        ).grid(row=15, column=0, columnspan=3, sticky="w", pady=6)

        ttk.Label(root, text="Title card transparency").grid(
            row=16, column=0, sticky="w", padx=(0, 10), pady=8
        )
        self.title_card_transparency_entry = ttk.Entry(
            root,
            textvariable=self.title_card_transparency_var,
            width=10,
        )
        self.title_card_transparency_entry.grid(row=16, column=1, sticky="w", pady=8)

        ttk.Checkbutton(
            root,
            text="Show white card on vocab slides",
            variable=self.show_vocab_card_var,
            command=self._on_vocab_card_toggle,
        ).grid(row=17, column=0, columnspan=3, sticky="w", pady=6)

        ttk.Label(root, text="Vocab card transparency").grid(
            row=18, column=0, sticky="w", padx=(0, 10), pady=8
        )
        self.vocab_card_transparency_entry = ttk.Entry(
            root,
            textvariable=self.vocab_card_transparency_var,
            width=10,
        )
        self.vocab_card_transparency_entry.grid(row=18, column=1, sticky="w", pady=8)

        help_text = (
            "CSV must contain Term and Definition headers.\n"
            "Header matching is case-insensitive, and one CSV file is processed at a time.\n"
            "The file must also contain enough unique rows for the selected set size.\n\n"
            "Optional background folder: if one image is present it is used for all slides;\n"
            "otherwise choose fixed or cycle mode. In cycle mode, leave the range blank to\n"
            "cycle through all images, or provide both start and end image numbers.\n\n"
            "Transparency values must be between 0.0 (opaque) and 1.0 (fully transparent)."
        )
        ttk.Label(root, text=help_text, justify="left").grid(
            row=19, column=0, columnspan=3, sticky="w", pady=(12, 8)
        )

        ttk.Button(root, text="Generate PPTX", command=self.generate).grid(
            row=20, column=0, columnspan=3, sticky="ew", pady=(8, 12)
        )

        ttk.Label(root, text="Status").grid(row=21, column=0, sticky="nw", pady=(4, 6))
        status_frame = ttk.Frame(root)
        status_frame.grid(row=21, column=1, columnspan=2, sticky="nsew", pady=(4, 6))
        root.rowconfigure(21, weight=1)
        status_frame.columnconfigure(0, weight=1)
        status_frame.rowconfigure(1, weight=1)

        ttk.Label(status_frame, textvariable=self.status_var).grid(
            row=0, column=0, sticky="w", pady=(0, 6)
        )
        self.log = tk.Text(status_frame, height=10, wrap="word", state="disabled")
        self.log.grid(row=1, column=0, sticky="nsew")

        self._on_background_mode_changed()
        self._on_title_card_toggle()
        self._on_vocab_card_toggle()

    def _require_int(self, value: str, field_name: str) -> int:
        text = value.strip()
        if not text:
            raise ValueError(f"{field_name} is required.")
        try:
            return int(text)
        except ValueError as exc:
            raise ValueError(f"{field_name} must be a whole number.") from exc

    def _optional_int(self, value: str, field_name: str) -> int | None:
        text = value.strip()
        if not text:
            return None
        try:
            return int(text)
        except ValueError as exc:
            raise ValueError(f"{field_name} must be a whole number.") from exc

    def _require_float(self, value: str, field_name: str) -> float:
        text = value.strip()
        if not text:
            raise ValueError(f"{field_name} is required.")
        try:
            return float(text)
        except ValueError as exc:
            raise ValueError(f"{field_name} must be a number.") from exc

    def _validate_inputs(
        self,
    ) -> tuple[
        str,
        str | None,
        int,
        int,
        int,
        str,
        bool,
        bool,
        str | None,
        str,
        int | None,
        int | None,
        int | None,
        float,
        float,
        bool,
        float,
        bool,
        float,
    ]:
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

        background_dir = self.background_dir_var.get().strip() or None
        background_mode = self.background_mode_var.get().strip() or "cycle"
        background_image_number = self._optional_int(
            self.background_image_number_var.get(),
            "Fixed image number",
        )
        background_cycle_start = self._optional_int(
            self.background_cycle_start_var.get(),
            "Cycle start image number",
        )
        background_cycle_end = self._optional_int(
            self.background_cycle_end_var.get(),
            "Cycle end image number",
        )

        title_slide_overlay_transparency = self._require_float(
            self.title_slide_overlay_transparency_var.get(),
            "Title-slide overlay transparency",
        )
        vocab_slide_overlay_transparency = self._require_float(
            self.vocab_slide_overlay_transparency_var.get(),
            "Vocab-slide overlay transparency",
        )
        show_title_card = self.show_title_card_var.get()
        title_card_transparency = self._require_float(
            self.title_card_transparency_var.get(),
            "Title card transparency",
        )
        show_vocab_card = self.show_vocab_card_var.get()
        vocab_card_transparency = self._require_float(
            self.vocab_card_transparency_var.get(),
            "Vocab card transparency",
        )

        if set_count < 1:
            raise ValueError("Sets must be at least 1.")
        if set_size < 1:
            raise ValueError("Set size must be at least 1.")
        if primary_side not in {"term", "definition"}:
            raise ValueError("Large text shows must be either Term or Definition.")
        if background_mode not in {"fixed", "cycle"}:
            raise ValueError("Background mode must be either fixed or cycle.")
        if not 0.0 <= title_slide_overlay_transparency <= 1.0:
            raise ValueError(
                "Title-slide overlay transparency must be between 0.0 and 1.0."
            )
        if not 0.0 <= vocab_slide_overlay_transparency <= 1.0:
            raise ValueError(
                "Vocab-slide overlay transparency must be between 0.0 and 1.0."
            )
        if not 0.0 <= title_card_transparency <= 1.0:
            raise ValueError("Title card transparency must be between 0.0 and 1.0.")
        if not 0.0 <= vocab_card_transparency <= 1.0:
            raise ValueError("Vocab card transparency must be between 0.0 and 1.0.")

        if background_dir is None:
            background_image_number = None
            background_cycle_start = None
            background_cycle_end = None

        return (
            csv_file,
            output,
            set_count,
            set_size,
            seed,
            primary_side,
            show_alternate,
            export_selected_terms,
            background_dir,
            background_mode,
            background_image_number,
            background_cycle_start,
            background_cycle_end,
            title_slide_overlay_transparency,
            vocab_slide_overlay_transparency,
            show_title_card,
            title_card_transparency,
            show_vocab_card,
            vocab_card_transparency,
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

    def _on_background_mode_changed(self, event=None) -> None:
        mode = self.background_mode_var.get().strip() or "cycle"
        has_dir = bool(self.background_dir_var.get().strip())

        fixed_state = "normal" if has_dir and mode == "fixed" else "disabled"
        cycle_state = "normal" if has_dir and mode == "cycle" else "disabled"
        overlay_state = "normal" if has_dir else "disabled"

        self.fixed_image_spinbox.configure(state=fixed_state)
        self.cycle_start_spinbox.configure(state=cycle_state)
        self.cycle_end_spinbox.configure(state=cycle_state)
        self.title_slide_overlay_transparency_entry.configure(state=overlay_state)
        self.vocab_slide_overlay_transparency_entry.configure(state=overlay_state)

    def _on_title_card_toggle(self) -> None:
        state = "normal" if self.show_title_card_var.get() else "disabled"
        self.title_card_transparency_entry.configure(state=state)

    def _on_vocab_card_toggle(self) -> None:
        state = "normal" if self.show_vocab_card_var.get() else "disabled"
        self.vocab_card_transparency_entry.configure(state=state)

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

    def choose_background_dir(self) -> None:
        initial = self.background_dir_var.get().strip()
        dirname = filedialog.askdirectory(
            title="Choose background image folder",
            initialdir=initial if initial else "",
            mustexist=True,
        )
        if not dirname:
            return
        self.background_dir_var.set(dirname)
        self.status_var.set("Background folder selected.")
        self._append_log(f"Background folder: {dirname}")
        self._on_background_mode_changed()

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
                background_dir,
                background_mode,
                background_image_number,
                background_cycle_start,
                background_cycle_end,
                title_slide_overlay_transparency,
                vocab_slide_overlay_transparency,
                show_title_card,
                title_card_transparency,
                show_vocab_card,
                vocab_card_transparency,
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

            if background_dir:
                self._append_log(f"Background folder: {background_dir}")
                self._append_log(f"Background mode: {background_mode}")
                self._append_log(
                    "Title-slide overlay transparency: "
                    f"{title_slide_overlay_transparency}"
                )
                self._append_log(
                    "Vocab-slide overlay transparency: "
                    f"{vocab_slide_overlay_transparency}"
                )
                if background_mode == "fixed" and background_image_number is not None:
                    self._append_log(
                        f"Fixed background image number: {background_image_number}"
                    )
                elif background_mode == "cycle":
                    if (
                        background_cycle_start is not None
                        and background_cycle_end is not None
                    ):
                        self._append_log(
                            "Cycle background image range: "
                            f"{background_cycle_start} to {background_cycle_end}"
                        )
                    else:
                        self._append_log("Cycle background image range: all images")

            self._append_log(f"Show title card: {show_title_card}")
            if show_title_card:
                self._append_log(f"Title card transparency: {title_card_transparency}")

            self._append_log(f"Show vocab card: {show_vocab_card}")
            if show_vocab_card:
                self._append_log(f"Vocab card transparency: {vocab_card_transparency}")

            pptx_path, csv_out = generate_from_inputs(
                csv_file=csv_file,
                output=str(output_path),
                set_count=set_count,
                set_size=set_size,
                seed=seed,
                primary_side=primary_side,
                show_alternate=show_alternate,
                export_selected_terms=export_selected_terms,
                background_dir=background_dir,
                background_mode=background_mode,
                background_image_number=background_image_number,
                background_cycle_start=background_cycle_start,
                background_cycle_end=background_cycle_end,
                title_slide_overlay_transparency=title_slide_overlay_transparency,
                vocab_slide_overlay_transparency=vocab_slide_overlay_transparency,
                show_title_card=show_title_card,
                title_card_transparency=title_card_transparency,
                show_vocab_card=show_vocab_card,
                vocab_card_transparency=vocab_card_transparency,
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
                "Please check the CSV file, background folder, and output location, "
                "then try again.",
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
