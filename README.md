# PoLab Analyzer (`polab_analyzer`)

A professional, cross-platform command-line toolkit for automated embryo stripe and fluorescence intensity profile analysis. This package bundles two primary analysis workflows, **`eve`** and **`hb`**, making them directly executable from your terminal on macOS, Windows, and Linux.

---

## Table of Contents

- [Installation](#installation)
- [Usage & Commands](#usage--commands)
  - [`hb` — Hunchback Profile Batch Analyzer](#hb--hunchback-profile-batch-analyzer)
  - [`eve` — Even-skipped Stripe Batch Analyzer](#eve--even-skipped-stripe-batch-analyzer)
- [Parameter Reference](#parameter-reference)
- [Input File Requirements](#input-file-requirements)
- [Troubleshooting](#troubleshooting)

---

## Installation

### Globally via PyPI (Recommended)

Once published, anyone can install the toolkit and all its dependencies directly from the terminal:

```bash
pip install polab_analyzer
```

### Locally from a Wheel File

If you are distributing the package directly within the lab via a `.whl` file:

```bash
pip install polab_analyzer-0.1.0-py3-none-any.whl
```

---

## Usage & Commands

After installation, the commands `eve` and `hb` are injected automatically into your system's `PATH`. You do not need to type `python` to run them.

### `hb` — Hunchback Profile Batch Analyzer

The `hb` tool processes Hunchback fluorescence intensity profiles to extract the spatial midpoint where the signal drops off.

**Batch Folder Mode** — process a directory filled with single-sample Excel files:

```bash
hb --folder /path/to/excel_folder/
```

**Single File Mode** — process a single multi-sheet Excel workbook where each sheet (`s1`, `s2`, etc.) represents an embryo sample:

```bash
hb --file /path/to/dataset.xlsx
```

### `eve` — Even-skipped Stripe Batch Analyzer

The `eve` tool utilizes advanced peak detection algorithms (`savgol_filter` and `find_peaks`) to extract and analyze standard spatial segmentation stripes in embryos. It includes an auto-tuning optimization pipeline to dynamically determine the best processing parameters.

**Single File Mode (with parameter grid-search):**

```bash
eve --file dataset.xlsx --lower 15 --upper 85 --test true
```

**Folder Mode (with fixed parameters):**

```bash
eve --folder /path/to/excel_folder/ --lower 10 --upper 90 --test false --distance 15 --prominence 0.03 --height 0.1
```

---

## Parameter Reference

| Flag | Short | Description | Default |
|------|-------|-------------|---------|
| `--folder` | — | Path to a directory of single-sample Excel files to batch process | — |
| `--file` | — | Path to a single multi-sheet Excel workbook | — |
| `--lower` | `-l` | Spatial lower boundary limit (% Length) used to filter false-positive peaks near the edges | — |
| `--upper` | `-u` | Spatial upper boundary limit (% Length) used to filter false-positive peaks | — |
| `--test` | — | `true` to run a hyperparameter sweep over a grid of peak boundaries; `false` to lock parameters | `true` |
| `--distance` | — | Minimum spacing (in samples) required between detected peaks | — |
| `--prominence` | — | Minimum prominence required for a peak to be counted | — |
| `--height` | — | Minimum normalized height required for a peak to be counted | — |

> **Note:** `--distance`, `--prominence`, and `--height` are only used by `eve` when `--test false` is set, since fixed-parameter mode requires explicit peak-detection settings instead of the auto-tuned grid search.

---

## Input File Requirements

Ensure your incoming biological input datasets are properly structured to match expectations:

- **Column Order:** Data arrays must be structured sequentially — Column 0 must represent the spatial axis coordinates (e.g., Length / Distance), and Column 1 must contain raw tracking observations (Fluorescence Intensity).
- **Sheet Naming:** Multi-sample single files expect sequential numbering format tags (e.g., `s1`, `s2`, `s3`, ...).

---

## Troubleshooting

### Command Not Found / Not Recognized

If you run `eve -h` or `hb -h` right after installation and get an execution error, your Python environment's binary directory is missing from your system PATH.

**macOS / Linux:**

Open your shell configuration profile (e.g., `~/.zshrc` or `~/.bashrc`) and add the following entry:

```bash
export PATH="$PATH:$HOME/.local/bin"
```

Then reload your shell configuration:

```bash
source ~/.zshrc
```

**Windows:**

Open the Environment Variables manager, locate the system `Path` variable under your user account, and append your active Python installation's `Scripts\` directory (e.g., `C:\Users\YourName\AppData\Local\Programs\Python\Python313\Scripts\`).

### File Format / Layout Errors

If processing fails or produces unexpected results, double-check that your Excel files follow the [Input File Requirements](#input-file-requirements) above — specifically column order and sheet naming conventions.

---

## License

*(Add license information here, e.g., MIT, GPL-3.0, or internal lab use only.)*

## Contact

*(Add maintainer name/email or lab contact information here.)*
