#!/usr/bin/env python3
from __future__ import annotations

import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from mainkata.domain import BackgroundOptions, GenerationOptions, VisualOptions
from mainkata.io import resolve_output_path
from mainkata.services.generator import generate_from_inputs


class VocabPptxGui(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Mainkata Term-Definition PPTX Generator")
        self.geometry("860x1080")
        self.minsize(820, 900)

        self.csv_path_var = tk.StringVar()
        self.output_var = tk.StringVar()
        self.style_config_var = tk.StringVar()

        self.set_count_var = tk.StringVar(value="6")
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

        self.override_title_slide_overlay_var = tk.BooleanVar(value=False)
        self.title_slide_overlay_transparency_var = tk.StringVar(value="0.22")
        self.override_vocab_slide_overlay_var = tk.BooleanVar(value=False)
        self.vocab_slide_overlay_transparency_var = tk.StringVar(value="0.22")

        self.override_title_card_var = tk.BooleanVar(value=False)
        self.show_title_card_var = tk.BooleanVar(value=True)
        self.override_title_card_transparency_var = tk.BooleanVar(value=False)
        self.title_card_transparency_var = tk.StringVar(value="0.18")

        self.override_vocab_card_var = tk.BooleanVar(value=False)
        self.show_vocab_card_var = tk.BooleanVar(value=True)
        self.override_vocab_card_transparency_var = tk.BooleanVar(value=False)
        self.vocab_card_transparency_var = tk.StringVar(value="0.18")

        self.status_var = tk.StringVar(value="Choose a CSV file to begin.")

        self.build_ui()

    def build_ui(self) -> None:
        root = ttk.Frame(self, padding=16)
        root.pack(fill="both", expand=True)
        root.columnconfigure(1, weight=1)

        ttk.Label(root, text="CSV file").grid(
            row=0, column=0, sticky="w", padx=(0, 10), pady=8
        )
        ttk.Entry(root, textvariable=self.csv_path_var).grid(
            row=0, column=1, sticky="ew", pady=8
        )
        ttk.Button(root, text="Browse", command=self.choose_csv).grid(
            row=0, column=2, sticky="ew", pady=8
        )

        ttk.Label(root, text="Output PPTX").grid(
            row=1, column=0, sticky="w", padx=(0, 10), pady=8
        )
        ttk.Entry(root, textvariable=self.output_var).grid(
            row=1, column=1, sticky="ew", pady=8
        )
        ttk.Button(root, text="Save as", command=self.choose_output).grid(
            row=1, column=2, sticky="ew", pady=8
        )

        ttk.Label(root, text="Style config TOML").grid(
            row=2, column=0, sticky="w", padx=(0, 10), pady=8
        )
        ttk.Entry(root, textvariable=self.style_config_var).grid(
            row=2, column=1, sticky="ew", pady=8
        )
        ttk.Button(root, text="Browse", command=self.choose_style_config).grid(
            row=2, column=2, sticky="ew", pady=8
        )

        ttk.Label(root, text="Sets").grid(
            row=3, column=0, sticky="w", padx=(0, 10), pady=8
        )
        ttk.Spinbox(
            root, from_=1, to=999, textvariable=self.set_count_var, width=10
        ).grid(row=3, column=1, sticky="w", pady=8)

        ttk.Label(root, text="Set size").grid(
            row=4, column=0, sticky="w", padx=(0, 10), pady=8
        )
        ttk.Spinbox(
            root, from_=1, to=999, textvariable=self.set_size_var, width=10
        ).grid(row=4, column=1, sticky="w", pady=8)

        ttk.Label(root, text="Seed").grid(
            row=5, column=0, sticky="w", padx=(0, 10), pady=8
        )
        ttk.Entry(root, textvariable=self.seed_var, width=12).grid(
            row=5, column=1, sticky="w", pady=8
        )

        ttk.Label(root, text="Large text shows").grid(
            row=6, column=0, sticky="w", padx=(0, 10), pady=8
        )
        primary_frame = ttk.Frame(root)
        primary_frame.grid(row=6, column=1, columnspan=2, sticky="w", pady=8)
        ttk.Radiobutton(
            primary_frame,
            text="Term",
            value="term",
            variable=self.primary_side_var,
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
        ).grid(row=7, column=0, columnspan=3, sticky="w", pady=6)

        ttk.Checkbutton(
            root,
            text="Generate selected-terms CSV",
            variable=self.export_selected_terms_var,
        ).grid(row=8, column=0, columnspan=3, sticky="w", pady=6)

        ttk.Separator(root).grid(row=9, column=0, columnspan=3, sticky="ew", pady=12)

        ttk.Label(root, text="Background image folder").grid(
            row=10, column=0, sticky="w", padx=(0, 10), pady=8
        )
        ttk.Entry(root, textvariable=self.background_dir_var).grid(
            row=10, column=1, sticky="ew", pady=8
        )
        ttk.Button(root, text="Browse", command=self.choose_background_dir).grid(
            row=10, column=2, sticky="ew", pady=8
        )

        ttk.Label(root, text="Background mode").grid(
            row=11, column=0, sticky="w", padx=(0, 10), pady=8
        )
        self.background_mode_combo = ttk.Combobox(
            root,
            textvariable=self.background_mode_var,
            values=("cycle", "fixed"),
            state="readonly",
            width=12,
        )
        self.background_mode_combo.grid(row=11, column=1, sticky="w", pady=8)
        self.background_mode_combo.bind(
            "<<ComboboxSelected>>", self.on_background_mode_changed
        )

        ttk.Label(root, text="Fixed image number").grid(
            row=12, column=0, sticky="w", padx=(0, 10), pady=8
        )
        self.fixed_image_spinbox = ttk.Spinbox(
            root,
            from_=1,
            to=9999,
            textvariable=self.background_image_number_var,
            width=10,
        )
        self.fixed_image_spinbox.grid(row=12, column=1, sticky="w", pady=8)

        ttk.Label(root, text="Cycle image range").grid(
            row=13, column=0, sticky="w", padx=(0, 10), pady=8
        )
        cycle_frame = ttk.Frame(root)
        cycle_frame.grid(row=13, column=1, columnspan=2, sticky="w", pady=8)
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

        ttk.Separator(root).grid(row=14, column=0, columnspan=3, sticky="ew", pady=12)

        ttk.Label(root, text="Style overrides").grid(
            row=15, column=0, sticky="w", padx=(0, 10), pady=(2, 8)
        )

        ttk.Checkbutton(
            root,
            text="Override title-slide overlay transparency",
            variable=self.override_title_slide_overlay_var,
            command=self.on_style_override_toggle,
        ).grid(row=16, column=0, columnspan=2, sticky="w", pady=4)
        self.title_slide_overlay_entry = ttk.Entry(
            root,
            textvariable=self.title_slide_overlay_transparency_var,
            width=10,
        )
        self.title_slide_overlay_entry.grid(row=16, column=2, sticky="w", pady=4)

        ttk.Checkbutton(
            root,
            text="Override vocab-slide overlay transparency",
            variable=self.override_vocab_slide_overlay_var,
            command=self.on_style_override_toggle,
        ).grid(row=17, column=0, columnspan=2, sticky="w", pady=4)
        self.vocab_slide_overlay_entry = ttk.Entry(
            root,
            textvariable=self.vocab_slide_overlay_transparency_var,
            width=10,
        )
        self.vocab_slide_overlay_entry.grid(row=17, column=2, sticky="w", pady=4)

        ttk.Checkbutton(
            root,
            text="Override title card visibility",
            variable=self.override_title_card_var,
            command=self.on_style_override_toggle,
        ).grid(row=18, column=0, columnspan=2, sticky="w", pady=4)
        self.show_title_card_checkbutton = ttk.Checkbutton(
            root,
            text="Show white card on title slides",
            variable=self.show_title_card_var,
        )
        self.show_title_card_checkbutton.grid(row=18, column=2, sticky="w", pady=4)

        ttk.Checkbutton(
            root,
            text="Override title card transparency",
            variable=self.override_title_card_transparency_var,
            command=self.on_style_override_toggle,
        ).grid(row=19, column=0, columnspan=2, sticky="w", pady=4)
        self.title_card_transparency_entry = ttk.Entry(
            root,
            textvariable=self.title_card_transparency_var,
            width=10,
        )
        self.title_card_transparency_entry.grid(row=19, column=2, sticky="w", pady=4)

        ttk.Checkbutton(
            root,
            text="Override vocab card visibility",
            variable=self.override_vocab_card_var,
            command=self.on_style_override_toggle,
        ).grid(row=20, column=0, columnspan=2, sticky="w", pady=4)
        self.show_vocab_card_checkbutton = ttk.Checkbutton(
            root,
            text="Show white card on vocab slides",
            variable=self.show_vocab_card_var,
        )
        self.show_vocab_card_checkbutton.grid(row=20, column=2, sticky="w", pady=4)

        ttk.Checkbutton(
            root,
            text="Override vocab card transparency",
            variable=self.override_vocab_card_transparency_var,
            command=self.on_style_override_toggle,
        ).grid(row=21, column=0, columnspan=2, sticky="w", pady=4)
        self.vocab_card_transparency_entry = ttk.Entry(
            root,
            textvariable=self.vocab_card_transparency_var,
            width=10,
        )
        self.vocab_card_transparency_entry.grid(row=21, column=2, sticky="w", pady=4)

        help_text = (
            "CSV must contain Term and Definition headers. Header matching is case-insensitive, "
            "and one CSV file is processed at a time. The file must also contain enough unique rows "
            "for the selected set size.\n\n"
            "Optional style config: choose a TOML file to define visual defaults.\n"
            "Style override checkboxes let you override only selected values in the GUI. "
            "Unchecked style override fields pass None to the generator so TOML/config defaults apply.\n\n"
            "Optional background folder: if one image is present it is used for all slides; otherwise choose fixed or cycle mode. "
            "In cycle mode, leave the range blank to cycle through all images, or provide both start and end image numbers."
        )
        ttk.Label(root, text=help_text, justify="left").grid(
            row=22,
            column=0,
            columnspan=3,
            sticky="w",
            pady=(12, 8),
        )

        ttk.Button(root, text="Generate PPTX", command=self.generate).grid(
            row=23,
            column=0,
            columnspan=3,
            sticky="ew",
            pady=(8, 12),
        )

        ttk.Label(root, text="Status").grid(row=24, column=0, sticky="nw", pady=(4, 6))
        status_frame = ttk.Frame(root)
        status_frame.grid(row=24, column=1, columnspan=2, sticky="nsew", pady=(4, 6))
        root.rowconfigure(24, weight=1)
        status_frame.columnconfigure(0, weight=1)
        status_frame.rowconfigure(1, weight=1)

        ttk.Label(status_frame, textvariable=self.status_var).grid(
            row=0, column=0, sticky="w", pady=(0, 6)
        )
        self.log = tk.Text(status_frame, height=10, wrap="word", state="disabled")
        self.log.grid(row=1, column=0, sticky="nsew")

        self.on_background_mode_changed()
        self.on_style_override_toggle()

    def require_int(self, value: str, field_name: str) -> int:
        text = value.strip()
        if not text:
            raise ValueError(f"{field_name} is required.")
        try:
            return int(text)
        except ValueError as exc:
            raise ValueError(f"{field_name} must be a whole number.") from exc

    def optional_int(self, value: str, field_name: str) -> int | None:
        text = value.strip()
        if not text:
            return None
        try:
            return int(text)
        except ValueError as exc:
            raise ValueError(f"{field_name} must be a whole number.") from exc

    def require_float(self, value: str, field_name: str) -> float:
        text = value.strip()
        if not text:
            raise ValueError(f"{field_name} is required.")
        try:
            return float(text)
        except ValueError as exc:
            raise ValueError(f"{field_name} must be a number.") from exc

    def collect_inputs(
        self,
    ) -> tuple[
        str, str | None, str | None, GenerationOptions, BackgroundOptions, VisualOptions
    ]:
        csv_file = self.csv_path_var.get().strip()
        if not csv_file:
            raise ValueError("Please choose a CSV file.")

        output = self.output_var.get().strip() or None
        style_config_file = self.style_config_var.get().strip() or None

        generation = GenerationOptions(
            set_count=self.require_int(self.set_count_var.get(), "Sets"),
            set_size=self.require_int(self.set_size_var.get(), "Set size"),
            seed=self.require_int(self.seed_var.get(), "Seed"),
            primary_side=self.primary_side_var.get().strip(),
            show_alternate=self.show_alternate_var.get(),
            export_selected_terms=self.export_selected_terms_var.get(),
        )

        background_dir = self.background_dir_var.get().strip() or None
        background = BackgroundOptions(
            background_dir=background_dir,
            background_mode=self.background_mode_var.get().strip() or "cycle",
            background_image_number=self.optional_int(
                self.background_image_number_var.get(),
                "Fixed image number",
            ),
            background_cycle_start=self.optional_int(
                self.background_cycle_start_var.get(),
                "Cycle start image number",
            ),
            background_cycle_end=self.optional_int(
                self.background_cycle_end_var.get(),
                "Cycle end image number",
            ),
        )

        if background_dir is None:
            background = BackgroundOptions()

        visual = VisualOptions(
            title_slide_overlay_transparency=(
                self.require_float(
                    self.title_slide_overlay_transparency_var.get(),
                    "Title-slide overlay transparency",
                )
                if self.override_title_slide_overlay_var.get()
                else None
            ),
            vocab_slide_overlay_transparency=(
                self.require_float(
                    self.vocab_slide_overlay_transparency_var.get(),
                    "Vocab-slide overlay transparency",
                )
                if self.override_vocab_slide_overlay_var.get()
                else None
            ),
            show_title_card=(
                self.show_title_card_var.get()
                if self.override_title_card_var.get()
                else None
            ),
            title_card_transparency=(
                self.require_float(
                    self.title_card_transparency_var.get(),
                    "Title card transparency",
                )
                if self.override_title_card_transparency_var.get()
                else None
            ),
            show_vocab_card=(
                self.show_vocab_card_var.get()
                if self.override_vocab_card_var.get()
                else None
            ),
            vocab_card_transparency=(
                self.require_float(
                    self.vocab_card_transparency_var.get(),
                    "Vocab card transparency",
                )
                if self.override_vocab_card_transparency_var.get()
                else None
            ),
        )

        return csv_file, output, style_config_file, generation, background, visual

    def validate_output_path(self, csv_file: str, output: str | None) -> Path:
        csv_path = Path(csv_file).expanduser().resolve()
        output_path = resolve_output_path(csv_path, output)

        if output_path.exists() and output_path.is_dir():
            raise ValueError(
                f"Output path is a folder, not a .pptx file: {output_path}"
            )

        return output_path

    def on_background_mode_changed(self, event=None) -> None:
        mode = self.background_mode_var.get().strip() or "cycle"
        has_dir = bool(self.background_dir_var.get().strip())

        fixed_state = "normal" if has_dir and mode == "fixed" else "disabled"
        cycle_state = "normal" if has_dir and mode == "cycle" else "disabled"

        self.fixed_image_spinbox.configure(state=fixed_state)
        self.cycle_start_spinbox.configure(state=cycle_state)
        self.cycle_end_spinbox.configure(state=cycle_state)

    def on_style_override_toggle(self) -> None:
        self.title_slide_overlay_entry.configure(
            state="normal"
            if self.override_title_slide_overlay_var.get()
            else "disabled"
        )
        self.vocab_slide_overlay_entry.configure(
            state="normal"
            if self.override_vocab_slide_overlay_var.get()
            else "disabled"
        )
        self.show_title_card_checkbutton.configure(
            state="normal" if self.override_title_card_var.get() else "disabled"
        )
        self.title_card_transparency_entry.configure(
            state="normal"
            if self.override_title_card_transparency_var.get()
            else "disabled"
        )
        self.show_vocab_card_checkbutton.configure(
            state="normal" if self.override_vocab_card_var.get() else "disabled"
        )
        self.vocab_card_transparency_entry.configure(
            state="normal"
            if self.override_vocab_card_transparency_var.get()
            else "disabled"
        )

    def choose_csv(self) -> None:
        filename = filedialog.askopenfilename(
            title="Choose CSV file",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
        )
        if not filename:
            return

        self.csv_path_var.set(filename)

        if not self.output_var.get().strip():
            output_path = resolve_output_path(
                Path(filename).expanduser().resolve(), None
            )
            self.output_var.set(str(output_path))

        self.status_var.set("CSV file selected.")
        self.append_log(f"CSV: {filename}")

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
        self.append_log(f"Output: {filename}")

    def choose_style_config(self) -> None:
        initial = self.style_config_var.get().strip()
        filename = filedialog.askopenfilename(
            title="Choose style config file",
            initialdir=str(Path(initial).parent) if initial else "",
            filetypes=[("TOML files", "*.toml"), ("All files", "*.*")],
        )
        if not filename:
            return

        self.style_config_var.set(filename)
        self.status_var.set("Style config file selected.")
        self.append_log(f"Style config: {filename}")

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
        self.append_log(f"Background folder: {dirname}")
        self.on_background_mode_changed()

    def generate(self) -> None:
        self.status_var.set("Validating inputs...")

        try:
            (
                csv_file,
                output,
                style_config_file,
                generation,
                background,
                visual,
            ) = self.collect_inputs()
            output_path = self.validate_output_path(csv_file, output)

            files_to_overwrite: list[Path] = []
            if output_path.exists():
                files_to_overwrite.append(output_path)

            if generation.export_selected_terms:
                csv_out_path = output_path.with_name(
                    f"{output_path.stem}_selectedterms.csv"
                )
                if csv_out_path.exists():
                    files_to_overwrite.append(csv_out_path)

            if files_to_overwrite:
                message = "The following files already exist:\n\n"
                message += "\n".join(str(path) for path in files_to_overwrite)
                message += "\n\nDo you want to overwrite them?"
                ok = messagebox.askyesno(
                    "Overwrite existing files?", message, parent=self
                )
                if not ok:
                    self.status_var.set("Generation cancelled.")
                    self.append_log("Cancelled: output file already exists.")
                    return

            self.status_var.set("Generating PPTX...")
            self.append_log(f"Generating: {output_path}")

            if style_config_file:
                self.append_log(f"Style config: {style_config_file}")

            if background.background_dir:
                self.append_log(f"Background folder: {background.background_dir}")
                self.append_log(f"Background mode: {background.background_mode}")

                if (
                    background.background_mode == "fixed"
                    and background.background_image_number is not None
                ):
                    self.append_log(
                        f"Fixed background image number: {background.background_image_number}"
                    )
                elif background.background_mode == "cycle":
                    if (
                        background.background_cycle_start is not None
                        and background.background_cycle_end is not None
                    ):
                        self.append_log(
                            "Cycle background image range: "
                            f"{background.background_cycle_start} to {background.background_cycle_end}"
                        )
                    else:
                        self.append_log("Cycle background image range: all images")

            if visual.title_slide_overlay_transparency is not None:
                self.append_log(
                    "Title-slide overlay transparency override: "
                    f"{visual.title_slide_overlay_transparency}"
                )
            if visual.vocab_slide_overlay_transparency is not None:
                self.append_log(
                    "Vocab-slide overlay transparency override: "
                    f"{visual.vocab_slide_overlay_transparency}"
                )
            if visual.show_title_card is not None:
                self.append_log(f"Show title card override: {visual.show_title_card}")
            if visual.title_card_transparency is not None:
                self.append_log(
                    f"Title card transparency override: {visual.title_card_transparency}"
                )
            if visual.show_vocab_card is not None:
                self.append_log(f"Show vocab card override: {visual.show_vocab_card}")
            if visual.vocab_card_transparency is not None:
                self.append_log(
                    f"Vocab card transparency override: {visual.vocab_card_transparency}"
                )

            result = generate_from_inputs(
                csv_file=csv_file,
                output=str(output_path),
                style_config_file=style_config_file,
                generation=generation,
                background=background,
                visual=visual,
            )

            self.status_var.set("Generation complete.")
            self.append_log(f"Created: {result.pptx_path}")
            if result.selected_terms_csv_path:
                self.append_log(f"Created: {result.selected_terms_csv_path}")
                done_message = (
                    f"Created:\n{result.pptx_path}\n{result.selected_terms_csv_path}"
                )
            else:
                done_message = f"Created:\n{result.pptx_path}"

            messagebox.showinfo("Done", done_message, parent=self)

        except ValueError as exc:
            self.status_var.set("Please fix the input values.")
            self.append_log(f"Validation error: {exc}")
            messagebox.showerror("Check your inputs", str(exc), parent=self)
        except PermissionError as exc:
            self.status_var.set("Cannot write to the selected output location.")
            self.append_log(f"Permission error: {exc}")
            messagebox.showerror("Output location not writable", str(exc), parent=self)
        except FileNotFoundError as exc:
            self.status_var.set("Input file not found.")
            self.append_log(f"File error: {exc}")
            messagebox.showerror("File not found", str(exc), parent=self)
        except Exception as exc:
            self.status_var.set("Generation failed.")
            self.append_log(f"Unexpected error: {exc}")
            messagebox.showerror(
                "Generation failed",
                "Something went wrong while creating the PowerPoint. "
                "Please check the CSV file, style config file, background folder, "
                "and output location, then try again.",
                parent=self,
            )

    def append_log(self, message: str) -> None:
        self.log.configure(state="normal")
        self.log.insert("end", message + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")


def main() -> None:
    app = VocabPptxGui()
    app.mainloop()


if __name__ == "__main__":
    main()
