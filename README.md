# Mainkata

Mainkata generates classroom vocabulary PowerPoint slide decks from a simple CSV file.

It is designed for teachers who want a quick way to create vocabulary games and revision slides without manually building each slide in PowerPoint.

## Features

- Create PowerPoint decks from a CSV with `Term` and `Definition` columns.
- Generate multiple randomised sets from one source file.
- Choose whether the primary slide text is the term or the definition.
- Optionally hide the alternate side.
- Export a companion CSV of selected terms.
- Use built-in backgrounds or your own image backgrounds.
- Apply optional style settings from a TOML config file.
- Launch from either the command line or the GUI.

## Install

### Recommended: pipx

```bash
pipx install mainkata
```

This is the easiest way to install Mainkata as an app-like tool without managing a separate virtual environment.

### Alternative: pip

```bash
pip install mainkata
```

## Quick start

Create a CSV file like this:

```csv
Term,Definition
Algorithm,A step-by-step procedure for solving a problem
Variable,A named value that can change in a program
Loop,A structure that repeats instructions
Function,A reusable block of code
```

Then run:

```bash
mainkata my_vocab.csv
```

This creates a PowerPoint deck based on your CSV file.

## Common usage

Generate 6 sets of 10 slides:

```bash
mainkata my_vocab.csv --sets 6 --set-size 10
```

Use definitions as the main slide text:

```bash
mainkata my_vocab.csv --primary-side definition
```

Hide the alternate text:

```bash
mainkata my_vocab.csv --hide-alternate
```

Export a selected-terms CSV as well:

```bash
mainkata my_vocab.csv --export-selected-terms
```

Use a custom style config:

```bash
mainkata my_vocab.csv --style-config ~/config/mainkata/style.toml
```

Use a folder of background images:

```bash
mainkata my_vocab.csv --background-dir ./backgrounds
```

Launch the GUI:

```bash
mainkata-gui
```

## CSV format

Your CSV file must include these headers:

- `Term`
- `Definition`

Header matching is case-insensitive.

Blank rows are ignored. Incomplete rows are rejected. Duplicate term-definition pairs are removed before slide generation.

## Style config

Mainkata can load an optional TOML style configuration file.

If no explicit `--style-config` file is given, Mainkata can look for:

- `$XDG_CONFIG_HOME/mainkata/style.toml`
- `~/.config/mainkata/style.toml`

If no config file is found, built-in defaults are used.

## Who it is for

Mainkata is especially useful for:

- classroom teachers
- learning support teachers
- teacher aides
- homeschool educators
- tutors

## Requirements

- Python 3.10 or newer

## Development

Build the package:

```bash
python -m build
```

Run tests:

```bash
pytest
```

Upload to TestPyPI:

```bash
python -m twine upload --repository testpypi dist/*
```

## License

MIT
