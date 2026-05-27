# Term-definition PPTX generator

This project generates PowerPoint vocabulary or content-review decks from a CSV file with two columns: `Term` and `Definition`.[file:31]

## Project structure

- `generate.py` — core generation logic shared by both interfaces.
- `cli.py` — command-line wrapper.
- `gui.py` — Tkinter desktop GUI.
- `requirements.txt` — pip-installable Python dependency list.

## CSV format

Your CSV must have exactly these two headers, case-insensitive:

```csv
Term,Definition
anggota,member
antri,to queue/line up
```

The script reads one CSV file at a time and requires enough unique rows to satisfy the chosen `--set-size` value.[file:31]

## Install

Create and activate a virtual environment:

### Linux / macOS

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Tkinter note

Tkinter is the standard Python interface to Tcl/Tk and is available on most Unix platforms, but it is not installed via `pip` and may need to be installed separately depending on how Python was installed on the machine.[web:35][web:44]

If `gui.py` fails with `No module named tkinter`:

- Ubuntu/Debian: install `python3-tk` from the system package manager.[web:44]
- macOS with Homebrew Python: install `python-tk` with Homebrew.[web:40][web:44]
- macOS: using a current Python.org installer is also a common way to get a working Tkinter/Tcl-Tk setup.[web:42][web:48]

Quick test:

```bash
python3 -c "import tkinter; tkinter._test()"
```

## Command-line usage

Basic usage:

```bash
python cli.py topic_1.csv
```

This creates:

- `topic_1_vocab_sets.pptx`

Specify the output filename:

```bash
python cli.py topic_1.csv -o topic1_games.pptx
```

Change the randomisation seed:

```bash
python cli.py topic_1.csv --seed 99
```

Optional controls:

```bash
python cli.py topic_1.csv --sets 6 --set-size 10
```

Generate the companion selected-terms CSV:

```bash
python cli.py topic_1.csv --export-selected-terms
```

Show definitions in large text instead of terms:

```bash
python cli.py topic_1.csv --primary-side definition
```

Hide the alternate side from slides:

```bash
python cli.py topic_1.csv --hide-alternate
```

Combine options:

```bash
python cli.py topic_1.csv --primary-side definition --hide-alternate --export-selected-terms
```

## GUI usage

Launch the desktop app:

```bash
python gui.py
```

The GUI lets you:

- Choose the input CSV file.
- Choose the output PPTX filename.
- Set `Sets`, `Set size`, and `Seed`.
- Choose whether large text shows `Term` or `Definition`.
- Choose whether the alternate side is shown in smaller text.
- Choose whether to generate the companion selected-terms CSV.
- Generate the PPTX without using the command line.[file:31]

## What it does

- Reads one CSV file with `Term` and `Definition` columns.[file:31]
- Randomly creates the requested number of sets with unique items inside each set.[file:31]
- Allows items to reappear across different sets because each set is sampled independently.[file:31]
- Adds a title slide before each set.[file:31]
- Shows either the `Term` or `Definition` in large font depending on the selected option.[file:31]
- Optionally shows the alternate side in smaller text.[file:31]
- Optionally writes a companion CSV showing which terms were selected into each set.[file:31]

## Packaging note

If you later want to distribute this to non-technical Mac and Linux users as a double-clickable app, you would normally package the app separately on each target platform even though the Python source remains shared.[web:26]
