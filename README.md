# Unreal Engine A/B Metric Comparison Tool

Author: Marvin Schubert  
Version: 1.0.0  
License: MIT

---
## 1. Purpose
This tool assembles per‑scene performance aggregates for two Unreal Engine rendering / content pipeline variants ("A" and "B") and produces an Excel workbook with:
* Per‑scene aggregated metrics for Variant A and Variant B
* A per‑scene comparison sheet including the relative delta (Δ B vs A in %)
* A global summary sheet across all scenes

It is designed to be transparent and reproducible in line with the methodological principles of the accompanying bachelor thesis (method section excerpt provided by the author). The implementation emphasises:
* Explicit aggregation rules (mean of per‑run metrics; percentiles already computed per run upstream)
* Stable metric ordering
* Robust label matching (tolerant aliases)
* Optional FPS recomputation from frametime means (to avoid averaging instantaneous FPS directly)

---
## 2. Methodological Foundations (Summary)
The underlying study evaluates two rendering / content pipeline configurations:
* **Pipeline A**: Modern UE5 features (e.g. Lumen for Global Illumination & Reflections, Nanite for virtualised geometry)
* **Pipeline B**: Classical baked pipeline (Lightmass GI, Reflection Captures + SSR, conventional LOD chains)

### 2.1 Experimental Design
* Fixed capture windows: **20 seconds per run**.
* Per scene & variant: **≥ 3 repeated runs** (identical conditions).
* Triggering: Automated Level Blueprint event (key 5) executes `csvprofile start` / `csvprofile stop`.
* Non‑varied parameters (camera path, resolution, project profiles) held constant to isolate pipeline effects.

### 2.2 Primary & Secondary Indicators
* **Primary metric:** Frametime p95 (95th percentile in ms) — emphasises tail latency while reducing single outlier influence.
* **Secondary metrics:** GPU Time (mean & p95), Draw Calls, visible Primitives, Local VRAM, Shader Memory, mean Frametime, derived FPS, frame count (N).
* **Scaling aspect:** Scene 2 acts as a micro‑benchmark to observe geometric scaling (Δ p95 vs visible Primitives / Draw Calls / VRAM).

### 2.3 Aggregation Rules Implemented Here
* Input sheets already contain per‑run metrics (including p95 values). This tool computes **arithmetic means across runs** for each metric (e.g. scene mean of per‑run p95 values — *not* recomputing percentile over concatenated raw frames).
* Delta (B vs A) = `(B - A) / A * 100%` when A is finite.
* Optional FPS recomputation (default ON): For each run: `FPS_run = 1000 / FrametimeMeanRun_ms`; reported scene FPS = mean(FPS_run). This avoids discouraging statistical pitfalls of averaging instantaneous FPS.

### 2.4 Data Quality & Exclusion
Run filtering (if applied upstream) follows the thesis rules: discard runs disrupted by user error, transient system activity, or inconsistent configuration. This script assumes only valid runs remain in the Excel sources.

### 2.5 Validity & Reliability (Context)
* **Internal validity:** Controlled environment; unchanged render/config parameters; fixed camera path; constant measurement window length.
* **Reliability:** Repetition (≥3 runs) and explicit aggregation rules.
* **Transparency:** Deterministic parsing and formatting; reproducible arithmetic operations; consistent file naming.

---
## 3. Metrics (German Labels Retained)
| Canonical Label            | Meaning (English)                              | Aggregation Here            |
|----------------------------|-----------------------------------------------|-----------------------------|
| N                          | Frame count in window                         | Mean of per‑run frame counts|
| Frametime Ø [ms]           | Mean frametime                                | Mean of per‑run means       |
| Frametime p95 [ms]         | 95th percentile frametime (per run upstream)  | Mean of per‑run p95 values  |
| FPS Ø [#]                  | Mean FPS (recomputed option)                  | Mean of derived per‑run FPS |
| GPU-Zeit Ø [ms]            | Mean GPU time                                 | Mean                        |
| GPU-Zeit p95 [ms]          | 95th percentile GPU time                      | Mean of per‑run p95 values  |
| Draw Calls Ø [#]           | Mean Draw Calls                               | Mean                        |
| Primitives Ø [#]           | Mean visible primitives                       | Mean                        |
| Local VRAM [MB]            | Mean local GPU memory usage                   | Mean                        |
| Shader Mem [MB]            | Mean shader memory                            | Mean                        |

---
## 4. Input Formats
### 4.1 Single Workbook Mode
A file `messungen_auswertung.xlsx` containing sheets named:
```
Scene1_A, Scene1_B, Scene2_A, Scene2_B, ...
```
Each sheet:
* Column A: metric labels (German)
* Columns B..K: numeric values per run (German formatting tolerated: thousand dot, decimal comma)

### 4.2 Two Workbook Mode
Two files (auto‑detected or explicitly provided):
* `messungen_auswertung_a.xlsx`
* `messungen_auswertung_b.xlsx`

Sheets may be named either `Scene{n}` (variant inferred from file) or `Scene{n}_A` / `Scene{n}_B`.

### 4.3 Label Matching
The tool normalises labels (lowercase, removing diacritics & punctuation) and matches them via tolerant alias sets. Unrecognised rows are ignored.

---
## 5. Output
Generated workbook (default): `messungen_ab_vergleich.xlsx`
Per scene:
* `Scene{n}_Agg_A` – aggregated metrics for Variant A
* `Scene{n}_Agg_B` – aggregated metrics for Variant B
* `Scene{n}_Vergleich` – comparison (Metric | A (Ø) | B (Ø) | Δ B vs A [%])

Global sheet `Gesamtübersicht` summarises all scenes vertically, separated by blank lines.

Numeric formatting uses German style (thousand separator '.' and decimal comma ',') to remain consistent with upstream artefacts.

---
## 6. Installation
### 6.1 Requirements
Python 3.10+ (recommended). Install dependencies:
```bash
pip install -r requirements.txt
```
`requirements.txt` includes:
```
pandas
numpy
openpyxl
```

### 6.2 Virtual Environment (Optional but Recommended)
```bash
python -m venv venv
# Windows PowerShell
./venv/Scripts/Activate.ps1
pip install -r requirements.txt
```

---
## 7. Usage
From the project directory:

### 7.1 Single Workbook
```bash
python ue_ab.py
```
(Requires `messungen_auswertung.xlsx`.)

### 7.2 Two Explicit Workbooks
```bash
python ue_ab.py --a messungen_auswertung_a.xlsx --b messungen_auswertung_b.xlsx
```

### 7.3 Automatic Detection
If both `messungen_auswertung_a.xlsx` and `messungen_auswertung_b.xlsx` are present:
```bash
python ue_ab.py --auto
```
(Autodetect also triggers without `--auto` if both files exist and single workbook is absent.)

### 7.4 Custom Output Name
```bash
python ue_ab.py --a a.xlsx --b b.xlsx --out comparison.xlsx
```

### 7.5 Disable FPS Recalculation
```bash
python ue_ab.py --no-recompute-fps
```

### 7.6 Exit Codes / Errors
* Missing expected input files → non‑zero exit with explanatory message.
* No matching sheets parsed → runtime error message.

---
## 8. Reproducibility & Data Quality Checklist
| Aspect                | Practice Implemented |
|-----------------------|----------------------|
| Fixed window length   | 20 s per run (upstream) |
| Runs per condition    | ≥ 3 (enforced upstream) |
| Percentile method     | Per run upstream (NumPy p95 linear) averaged here |
| FPS derivation        | Derived from per‑run frametime means (default) |
| Missing values        | Ignored in per‑run lists; NaN if no valid values |
| Label robustness      | Alias normalisation & substring pattern matching |
| Delta computation     | (B − A)/A * 100% with guard against NaN/inf |

---
## 9. Extensibility
Potential extensions:
* Additional percentile columns (p99) if upstream provided.
* CSV export of comparison tables.
* Confidence intervals via run variance (if run count high enough).
* Median‑based aggregation toggle for robustness against skew.

---
## 10. Citation Guidance
If you cite or reference the tool in academic work, consider:
```
Schubert, M. (2025). Unreal Engine A/B Metric Comparison Tool (Version 1.0.0) [Software]. MIT License.
```
And, where appropriate, reference the methodology section of the associated thesis for experimental design.

---
## 11. License
Released under the MIT License (see LICENSE file). Copyright (c) 2025 Marvin Schubert.

---
## 12. Disclaimer
This tool assumes upstream correctness of per‑run metrics (especially p95). It does not validate raw frame time series or recompute percentiles from raw frames; it focuses on consistent aggregation and presentation.

---
## 13. Changelog
* 1.0.0: Initial public release, bilingual metric labels retained, FPS recomputation added (default ON), dual input mode.

---
## 14. Support
For clarifications or methodological questions, consult the associated thesis methodology chapter. Bug reports can be documented via issue tracking (if repository hosting is added in future).
