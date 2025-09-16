<div align="center">

# UE A/B Performance Comparison

**Excel post‑processing tool for analyzing Unreal Engine pipeline variants (A vs B)**

[![Python](https://img.shields.io/badge/Python-3.10+-blue.svg)](https://www.python.org/)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![Status](https://img.shields.io/badge/Status-Research--grade-success.svg)]()

**Author:** Marvin Schubert  
**Version:** 1.0.0  
**Date:** September 2025  
**License:** MIT

</div>

## Overview
This repository contains the second (post‑processing) stage in a two‑step performance analysis pipeline for Unreal Engine:

1. **Upstream Tool (CSV → Aggregation)**: Robustly parses raw UE profiler `EXP_*.csv` files with variable column counts and produces one or two aggregation Excel workbooks containing per‑run metrics (means & p95 values).
2. **This Tool (Aggregation → A/B Comparison)**: Consumes those workbook(s) and generates a consolidated A/B comparison report with per‑scene deltas.

Result: A research‑grade Excel workbook showing structured performance differences between two rendering/content pipeline variants (A vs B) with reproducible aggregation logic.

## Scientific Background

### Motivation
Raw UE CSV exports are noisy and structurally inconsistent. Percentiles (p95) and mean values must be computed from data sets whose integrity is preserved in the upstream phase. This tool assumes that stage has already delivered trustworthy per‑run aggregates and focuses on transparent cross‑variant comparison.

### Core Aggregation Principle
All metrics in the input sheets are already per‑run aggregates. This script computes only the **arithmetic mean across runs per scene and variant**, plus a guarded percentage delta:

\(\Delta = \frac{B - A}{A} \times 100\) (only if A is finite and non‑zero)

### Why Recompute FPS?
Mean FPS reported upstream might be an arithmetic mean of instantaneous FPS values (statistically biased). Here (default ON) FPS is recomputed from each run's mean frametime:  
`FPS_run = 1000 / FrametimeMeanRun_ms` → scene FPS = mean(FPS_run). Disable with `--no-recompute-fps` if you need the original upstream mean.

## Features
- 🔄 Dual input modes: single combined workbook or two variant‑specific workbooks
- 🧮 Explicit delta calculation with division‑by‑zero safeguards
- 🎯 Metric ordering + robust German label alias matching
- 🧵 Locale‑aware numeric parsing (German & English; multi‑dot thousands)
- 🕒 Optional FPS recomputation from frametime means (default)
- 📊 Per‑scene comparison sheets + global summary sheet
- 🇩🇪 German‑style number formatting (decimal comma, NBSP thousands)
- 🧪 Debug diagnostics (`--debug`) for run counts & missing metrics

## Metrics (German Labels Retained)
| Label (German)        | Meaning (English)            | Aggregation (here)                 |
|-----------------------|------------------------------|------------------------------------|
| N                     | Frame count                  | Mean of per‑run counts             |
| Frametime Ø [ms]      | Mean frametime               | Mean of per‑run means              |
| Frametime p95 [ms]    | 95th percentile frametime    | Mean of per‑run p95 values         |
| FPS Ø [#]             | Mean FPS (recomputed option) | Mean of derived per‑run FPS        |
| GPU-Zeit Ø [ms]       | Mean GPU time                | Mean                                |
| GPU-Zeit p95 [ms]     | 95th percentile GPU time     | Mean of per‑run p95 values         |
| Draw Calls Ø [#]      | Draw calls                   | Mean                                |
| Primitives Ø [#]      | Visible primitives           | Mean                                |
| Local VRAM [MB]       | Local VRAM usage             | Mean                                |
| Shader Mem [MB]       | Shader memory usage          | Mean                                |

## Installation
```bash
pip install -r requirements.txt
```
Dependencies: `pandas`, `numpy`, `openpyxl` (Python 3.10+ recommended).

Optional virtual environment:
```bash
python -m venv venv
./venv/Scripts/Activate.ps1   # PowerShell
pip install -r requirements.txt
```

## Input Expectations
You must first generate aggregation workbook(s) using the upstream CSV analyzer.

Accepted inputs for this stage:

1. Single file: `messungen_auswertung.xlsx` containing sheets named `Scene{n}_A` and `Scene{n}_B`.
2. Two files: `messungen_auswertung_a.xlsx` & `messungen_auswertung_b.xlsx` with sheets either `Scene{n}` (variant inferred) or suffixed.

Sheet structure:
```
Column A : Metric label (German)
Columns B..K : Numeric run values (locale variants tolerated)
```

## Usage
Single workbook mode:
```bash
python ue_ab.py
```
Two explicit workbooks:
```bash
python ue_ab.py --a messungen_auswertung_a.xlsx --b messungen_auswertung_b.xlsx
```
Auto-detect variant files:
```bash
python ue_ab.py --auto
```
Custom output filename:
```bash
python ue_ab.py --auto --out vergleich_report.xlsx
```
Disable FPS recomputation:
```bash
python ue_ab.py --no-recompute-fps --auto
```
Debug diagnostics (run-level presence & missing metrics hints):
```bash
python ue_ab.py --debug --auto
```

## Output
Default output: `messungen_ab_vergleich.xlsx`

Per scene:
```
Scene{n}_Agg_A        Aggregated metrics (variant A)
Scene{n}_Agg_B        Aggregated metrics (variant B)
Scene{n}_Vergleich    Comparison with Δ B vs A [%] + note row explaining formula & rounding
```
Global sheet: `Gesamtübersicht` (vertical list of all metrics across scenes).

Formatting rules:
* German decimal comma, NBSP thousands grouping (e.g. `12 345,678`)
* Integer metrics: N, Draw Calls, Primitives (0 decimals)
* Other metrics & Δ: 3 decimals
* Δ formula note embedded in each comparison sheet

## Reproducibility & Data Quality
| Aspect                | Implementation Detail                                   |
|-----------------------|----------------------------------------------------------|
| Run repetition        | Upstream ensures ≥3 valid runs per scene/variant          |
| Window length         | 20s capture windows (upstream)                           |
| Percentiles           | Computed upstream per run (p95), averaged here            |
| FPS method            | Derived from frametime means (default)                    |
| Delta safety          | Guarded division (A finite & ≠ 0)                         |
| Locale parsing        | Mixed German/English formats, multi-dot thousands         |
| Missing values        | Ignored; metric becomes NaN if no valid run values        |
| Label robustness      | Normalisation + tolerant alias patterns                   |

## Example (Conceptual)
Notional scene comparison excerpt:
```
Metric                 A (Ø)        B (Ø)        Δ B vs A [%]
Frametime p95 [ms]     13,421       12,978       -3,303
Draw Calls Ø [#]       5 214        5 198        -0,307
Primitives Ø [#]       12 454 397   12 612 112   +1,266
FPS Ø [#]              83,742       84,105       +0,434
```

## Technical Implementation Highlights
- Pandas-based Excel ingestion with permissive label selection
- Safe numeric parsing with locale heuristics (dots+commas, multi-dot) 
- Structured workbook writer using `openpyxl`
- Modular aggregation + optional FPS recomputation
- Debug mode surfaces run distributions for transparency

## Extensibility Ideas
- Add confidence intervals (if run count sufficient)
- Optional median aggregation
- Export CSV summary alongside Excel
- Additional percentile columns (p99) when provided upstream

## Citation
If this tool contributes to academic work, cite it together with the upstream analyzer:
```bibtex
@software{schubert2025_ue_ab_comparison,
	author  = {Schubert, Marvin},
	title   = {UE A/B Performance Comparison Tool},
	year    = {2025},
	version = {1.0.0},
	url     = {https://github.com/REPO_PLACEHOLDER/ue-ab-comparison}
}
```

## License
MIT License – see `LICENSE`.

## Disclaimer
Assumes upstream correctness of per‑run aggregates. Does not recompute percentiles from raw frames; focuses on aggregation consistency and presentation.

## Changelog
| Version | Date        | Notes                                                     |
|---------|-------------|-----------------------------------------------------------|
| 1.0.0   | 2025-09     | Initial release: dual input, FPS recompute, locale parse  |

## Contact / Support
Questions about methodology / implementation: please reference accompanying thesis material first; future issue tracker (if repo public) for bug reports.

---
*Part of a Bachelor's thesis research pipeline on real‑time rendering performance analysis.*
