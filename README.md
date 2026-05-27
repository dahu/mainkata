# Mainkata - A Term-Definition PPTX Generator

Mainkata generates PowerPoint vocabulary or content-review decks from a CSV file with two columns: `Term` and `Definition`.

It provides:

- a command-line interface: `mainkata`
- a desktop GUI: `mainkata-gui`

## CSV format

Your CSV must include these two headers, case-insensitive:

```csv
Term,Definition
anggota,member
antri,to queue/line up
```

Mainkata reads one CSV file at a time and requires enough unique rows to satisfy the selected set size.

## Installation

### Requirements

- Python 3.10 or newer
- Git, if installing directly from GitHub
- PowerPoint is **not** required to generate `.pptx` files

### Recommended: install with pipx

`pipx` is the easiest way to install Mainkata as an app. It creates an isolated Python environment for the program and makes the `mainkata` and `mainkata-gui` commands available in your terminal.

### Linux

1. Install Python 3 and Git if they are not already installed.
2. Install `pipx`:

```bash
python3 -m pip install --user pipx
python3 -m pipx ensurepath
```

3. Restart your terminal.
4. Install Mainkata from GitHub:

```bash
pipx install "git+https://github.com/dahu/mainkata"
```

### macOS

If you use Homebrew:

```bash
brew install pipx
pipx ensurepath
pipx install "git+https://github.com/dahu/mainkata"
```

Or with Python only:

```bash
python3 -m pip install --user pipx
python3 -m pipx ensurepath
pipx install "git+https://github.com/dahu/mainkata"
```

Restart Terminal if `pipx` or `mainkata` is not found immediately.

### Windows 11

1. Install Python and ensure the Python Launcher (`py`) is available.
2. Open **PowerShell**.
3. Install `pipx`:

```powershell
py -m pip install --user pipx
py -m pipx ensurepath
```

4. Close and reopen PowerShell.
5. Install Mainkata from GitHub:

```powershell
pipx install "git+https://github.com/dahu/mainkata"
```

## Run Mainkata

### Command-line app

Generate a deck from a CSV file:

```bash
mainkata topic_1.csv
```

This creates:

- `topic_1_vocab_sets.pptx`

Specify the output filename:

```bash
mainkata topic_1.csv -o topic1_games.pptx
```

Change the randomisation seed:

```bash
mainkata topic_1.csv --seed 99
```

Adjust the number and size of sets:

```bash
mainkata topic_1.csv --sets 6 --set-size 10
```

Generate the companion selected-terms CSV:

```bash
mainkata topic_1.csv --export-selected-terms
```

Show definitions in large text instead of terms:

```bash
mainkata topic_1.csv --primary-side definition
```

Hide the alternate side from slides:

```bash
mainkata topic_1.csv --hide-alternate
```

Overwrite existing output files without prompting:

```bash
mainkata topic_1.csv --force
```

See all options:

```bash
mainkata --help
```

### Desktop GUI

Launch the GUI:

```bash
mainkata-gui
```

The GUI lets you:

- choose the input CSV file
- choose the output PPTX filename
- set `Sets`, `Set size`, and `Seed`
- choose whether large text shows `Term` or `Definition`
- choose whether the alternate side is shown in smaller text
- choose whether to generate the companion selected-terms CSV
- generate the PPTX without using the command line

## Tkinter note

The GUI uses Tkinter. Tkinter is included with many Python installations, but not all.

If `mainkata-gui` fails because Tkinter is missing:

- Ubuntu/Debian: install `python3-tk`
- Fedora: install `python3-tkinter`
- macOS: the Python.org installer usually includes a working Tkinter setup
- Windows: the standard Python.org installer typically includes Tkinter

Quick test:

```bash
python3 -c "import tkinter; tkinter._test()"
```

On Windows:

```powershell
py -c "import tkinter; tkinter._test()"
```

## Upgrade

If installed with `pipx`:

```bash
pipx upgrade mainkata
```

## Uninstall

```bash
pipx uninstall mainkata
```

## Install from source

Use this if you want to clone the repository and run Mainkata locally.

### Linux / macOS

```bash
git clone https://github.com/dahu/mainkata
cd mainkata
python3 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
python -m pip install .
```

### Windows 11

```powershell
git clone https://github.com/dahu/mainkata
cd mainkata
py -m venv .venv
.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install .
```

Then run:

```bash
mainkata --help
```

or:

```bash
mainkata-gui
```
