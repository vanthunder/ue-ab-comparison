"""\
A/B comparison generator for Unreal Engine CSV -> Excel aggregations.

Purpose
-------
Creates per–scene aggregated metrics (mean across runs) for two variants (A,B) and
builds a comparison sheet including the relative delta (B vs A in percent).

Input Modes
-----------
1) Single workbook (legacy): 'messungen_auswertung.xlsx' containing sheets "Scene{n}_A" / "Scene{n}_B".
2) Two-workbook mode: separate files for each variant
            - 'messungen_auswertung_a.xlsx' (or provided via --a PATH)
            - 'messungen_auswertung_b.xlsx' (or provided via --b PATH)
     Sheets may be named either "Scene{n}" (variant inferred from file) or already "Scene{n}_A" / "Scene{n}_B".

Per Sheet Format
----------------
First column: metric labels (German labels retained intentionally to match upstream export).
Columns 2..k: numeric run values (German decimal notation tolerated: thousand separators '.' and decimal comma ',').

Output
------
Workbook 'messungen_ab_vergleich.xlsx' with sheets per scene:
    - Scene{scene}_Agg_A
    - Scene{scene}_Agg_B
    - Scene{scene}_Vergleich  (tabular comparison: metric | A mean | B mean | Δ B vs A [%])
Plus a summary sheet (Gesamtübersicht) listing all scenes.

Methodological Notes
--------------------
* Each metric is aggregated as the arithmetic mean of per‑run values (runs are assumed pre‑computed
    for percentiles such as p95; thus we average the per‑run p95 values, not raw frame samples).
* Optional recomputation of FPS mean is provided (default ON) by deriving FPS per run from mean
    frametime (FPS_run = 1000 / FrametimeMeanRun_ms) and averaging those, avoiding arithmetic means of
    instantaneous FPS series.
* Delta (B vs A) = (B - A) / A * 100 %, computed when A is finite.

CLI
---
    python ue_ab.py                               # single workbook mode
    python ue_ab.py --a a.xlsx --b b.xlsx          # two separate workbooks
    python ue_ab.py --auto                        # auto-detect *_a.xlsx and *_b.xlsx
    python ue_ab.py --out result.xlsx             # custom output filename
    --no-recompute-fps                            # keep original FPS if present instead of recomputing

Author: Marvin Schubert (c) 2025
Version: 1.0.0
License: MIT
"""

from __future__ import annotations
import re
from pathlib import Path
import math
from typing import Dict, List, Tuple, Optional
import argparse

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter

# --------- Configuration ---------

# Expected metric ordering (kept in German to match upstream export labels)
ORDERED_LABELS = [
    "N",
    "Frametime Ø [ms]",
    "Frametime p95 [ms]",
    "FPS Ø [#]",
    "GPU-Zeit Ø [ms]",
    "GPU-Zeit p95 [ms]",
    "Draw Calls Ø [#]",
    "Primitives Ø [#]",
    "Local VRAM [MB]",
    "Shader Mem [MB]",
]

# Tolerant alias mapping for row labels (robust matching against slight variations)
# key = canonical target label, values = patterns (normalized)
LABEL_ALIASES: Dict[str, List[str]] = {
    "N": ["n"],
    "Frametime Ø [ms]": ["frametimeøms", "frametimeoms", "frametime ms", "frametimeø", "frametimeoems"],
    "Frametime p95 [ms]": ["frametimep95ms", "frametime p95", "p95 ms", "p95frametime"],
    "FPS Ø [#]": ["fpsø#", "fpsoe#", "fps", "fpsø"],
    "GPU-Zeit Ø [ms]": ["gpu-zeitøms", "gpuzeitøms", "gpuzeitoms", "gpuzeit ms", "gpu ø"],
    "GPU-Zeit p95 [ms]": ["gpu-zeitp95ms", "gpuzeitp95ms", "gpu p95"],
    "Draw Calls Ø [#]": ["drawcallsø#", "drawcalls", "draw calls"],
    "Primitives Ø [#]": ["primitivesø#", "primitives"],
    "Local VRAM [MB]": ["localvrammb", "vram", "gpu memory", "gpu mem", "local vram"],
    "Shader Mem [MB]": ["shadermemmb", "shader mem", "shader memory"],
}

# Output formatting precision rules (decimal places)
DECIMALS = {
    "N": 0,
    "FPS Ø [#]": 3,
    "Draw Calls Ø [#]": 0,
    "Primitives Ø [#]": 1,
    "Local VRAM [MB]": 3,
    "Shader Mem [MB]": 3,
    # default für Zeitmaße (ms):
}

# --------- Helpers ---------

def normalize(s: str) -> str:
    """Lowercase and strip non-alphanumerics for robust label matching."""
    return re.sub(r"[^a-z0-9]+", "", str(s).strip().lower())

def best_label_match(raw_label: str) -> Optional[str]:
    """Map a raw row label to the canonical target label, or None if no match."""
    n = normalize(raw_label)
    for target, patterns in LABEL_ALIASES.items():
        if n == normalize(target):
            return target
        if any(p in n for p in patterns):
            return target
    return None


def parse_number(val) -> Optional[float]:
    """Lenient numeric parser supporting German formatting and blanks.

    Rules:
        * Accept ints / floats directly (reject NaN/inf).
        * For strings: remove spaces, remove thousand separators '.', replace decimal comma ',' with '.'
        * Strip any trailing unit characters.
        * Return None if value is empty, dash or unparsable.
    """
    if val is None:
        return None
    if isinstance(val, (int, float, np.number)):
        if isinstance(val, float) and (math.isnan(val) or math.isinf(val)):
            return None
        return float(val)
    s = str(val).strip()
    if not s or s == "-":
        return None
    # Entferne Leerzeichen
    s = s.replace(" ", "")
    # Entferne Tausendertrennpunkte (.) und ersetze Dezimalkomma (,)
    # Achtung: erst Punkte raus, dann Komma -> Punkt
    s = s.replace(".", "").replace(",", ".")
    # Entferne evtl. Einheitensuffixe in Zahlenfeldern (sollten nicht vorkommen)
    s = re.sub(r"[^0-9.\-eE+]", "", s)
    try:
        return float(s)
    except Exception:
        return None

def fmt_de(label: str, val: Optional[float]) -> str:
    """Format value using German thousands separator and decimal comma with label-specific precision."""
    if val is None or (isinstance(val, float) and (math.isnan(val) or math.isinf(val))):
        return ""
    # Standard: 3 Nachkommastellen für ms / sonst DECIMALS
    prec = DECIMALS.get(label, 3)
    # Tausenderpunkt (deutsch), Dezimalkomma:
    # Wir formatieren zunächst mit US-Konvention und wandeln danach.
    if prec == 0:
        s = f"{val:,.0f}"
    else:
        s = f"{val:,.{prec}f}"
    return s.replace(",", "_").replace(".", ",").replace("_", ".")

def extract_scene_variant(sheet_name: str) -> Tuple[Optional[str], Optional[str]]:
    m = re.search(r"scene\s*(\d+)\s*[_-]\s*([ab])", sheet_name, re.IGNORECASE)
    if m:
        return m.group(1), m.group(2).upper()
    # Fallbacks
    m2 = re.search(r"scene\s*(\d+)", sheet_name, re.IGNORECASE)
    var = "A" if re.search(r"[_-]a\b", sheet_name, re.IGNORECASE) else ("B" if re.search(r"[_-]b\b", sheet_name, re.IGNORECASE) else None)
    return (m2.group(1) if m2 else None), var

# --------- Kernfunktionen ---------

def read_sheet_aggregation(xls_path: Path, sheet: str) -> Tuple[Dict[str, float], Dict[str, List[float]]]:
    """Read one worksheet and compute per-metric mean across run columns.

    Returns:
        agg:  metric label -> mean across runs (NaN if no valid values)
        runs: metric label -> list of individual run values
    """
    df = pd.read_excel(xls_path, sheet_name=sheet, header=None, engine="openpyxl")
    agg: Dict[str, float] = {}
    runs: Dict[str, List[float]] = {}
    for i in range(len(df)):
        raw_label = df.iat[i, 0]
        if not isinstance(raw_label, str):
            continue
        target_label = best_label_match(raw_label)
        if target_label is None:
            continue
        row_vals: List[float] = []
        for j in range(1, df.shape[1]):
            v = parse_number(df.iat[i, j])
            if v is not None:
                row_vals.append(v)
        runs[target_label] = row_vals
        agg[target_label] = float(np.mean(row_vals)) if row_vals else float("nan")
    return agg, runs

def build_comparison_for_scene(scene: str,
                               agg_A: Dict[str, float],
                               agg_B: Dict[str, float]) -> pd.DataFrame:
    """Create comparison dataframe for a scene (Metric | A mean | B mean | Delta [%])."""
    rows = []
    for label in ORDERED_LABELS:
        a = agg_A.get(label, float("nan"))
        b = agg_B.get(label, float("nan"))
        delta = float("nan")
        if a is not None and not (isinstance(a, float) and math.isnan(a)):
            delta = (b - a) / a * 100.0
        rows.append({
            "Kennzahl": label,
            "A (Ø)": a,
            "B (Ø)": b,
            "Δ B vs A [%]": delta
        })
    return pd.DataFrame(rows, columns=["Kennzahl", "A (Ø)", "B (Ø)", "Δ B vs A [%]"])

def write_scene_sheets(wb: Workbook,
                       scene: str,
                       agg_A: Dict[str, float],
                       agg_B: Dict[str, float],
                       cmp_df: pd.DataFrame) -> None:
    """Write three sheets per scene: aggregated A, aggregated B, and comparison."""
    def _write_agg(sheet_title: str, agg: Dict[str, float], variant: str):
        ws = wb.create_sheet(title=sheet_title)
        ws["A1"] = f"Aggregated metrics – Variant {variant}"
        ws["A1"].font = Font(bold=True)
        ws["A1"].alignment = Alignment(horizontal="left")

        ws["A2"] = "Metric"
        ws["B2"] = "Mean over runs"
        ws["A2"].font = ws["B2"].font = Font(bold=True)
        ws["A2"].alignment = ws["B2"].alignment = Alignment(horizontal="center")

        r = 3
        for label in ORDERED_LABELS:
            ws.cell(row=r, column=1, value=label)
            val = agg.get(label, float("nan"))
            ws.cell(row=r, column=2, value=fmt_de(label, val))
            r += 1

        # Breiten
        ws.column_dimensions["A"].width = 24
        ws.column_dimensions["B"].width = 16

    # Agg A
    _write_agg(f"Scene{scene}_Agg_A", agg_A, "A")
    # Agg B
    _write_agg(f"Scene{scene}_Agg_B", agg_B, "B")

    # Vergleich
    ws = wb.create_sheet(title=f"Scene{scene}_Vergleich")
    headers = ["Metric", "A (Ø)", "B (Ø)", "Δ B vs A [%]"]
    for ci, h in enumerate(headers, start=1):
        c = ws.cell(row=1, column=ci, value=h)
        c.font = Font(bold=True)
        c.alignment = Alignment(horizontal="center")

    # Write data rows; iterrows() keeps original column names (with spaces / symbols).
    for excel_row, (_, row) in enumerate(cmp_df.iterrows(), start=2):
        lab = row["Kennzahl"]
        ws.cell(row=excel_row, column=1, value=lab)
        ws.cell(row=excel_row, column=2, value=fmt_de(lab, row["A (Ø)"]))
        ws.cell(row=excel_row, column=3, value=fmt_de(lab, row["B (Ø)"]))
        ws.cell(row=excel_row, column=4, value=fmt_de("Δ", row["Δ B vs A [%]"]))

    # Spaltenbreiten:
    widths = [24, 16, 16, 16]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

def _collect_from_single_file(path: Path) -> List[Dict[str, object]]:
    """Collect A/B variant sheets from a single workbook (legacy mode)."""
    xls = pd.ExcelFile(path, engine="openpyxl")
    recs: List[Dict[str, object]] = []
    for sheet in xls.sheet_names:
        scene, var = extract_scene_variant(sheet)
        if scene is None or var not in ("A", "B"):
            continue
        agg, runs = read_sheet_aggregation(path, sheet)
        recs.append({
            "scene": str(scene),
            "variant": var,
            "agg": agg,
            "runs": runs,
            "sheet": sheet,
            "source": path.name,
        })
    return recs

def _collect_from_two_files(path_a: Path, path_b: Path) -> List[Dict[str, object]]:
    """Collect scene data from two workbooks (one per variant)."""
    def collect_one(p: Path, variant: str) -> List[Dict[str, object]]:
        if not p.exists():
            raise FileNotFoundError(f"Datei nicht gefunden: {p}")
        xls = pd.ExcelFile(p, engine="openpyxl")
        out: List[Dict[str, object]] = []
        for sheet in xls.sheet_names:
            scene, var_in_sheet = extract_scene_variant(sheet)
            # If sheet lacks variant suffix we assign the variant from the file context
            if scene is None:
                # Versuche einfache Szeneextraktion ohne Variantensuffix
                m = re.search(r"scene\s*(\d+)", sheet, re.IGNORECASE)
                if m:
                    scene = m.group(1)
            if scene is None:
                continue
            effective_variant = variant if var_in_sheet is None else var_in_sheet
            # Enforce passed variant if embedded suffix disagrees
            if effective_variant != variant:
                effective_variant = variant
            agg, runs = read_sheet_aggregation(p, sheet)
            out.append({
                "scene": str(scene),
                "variant": effective_variant,
                "agg": agg,
                "runs": runs,
                "sheet": sheet,
                "source": p.name,
            })
        return out
    recs_a = collect_one(path_a, "A")
    recs_b = collect_one(path_b, "B")
    return recs_a + recs_b

def main():
    parser = argparse.ArgumentParser(description="A/B comparison for Unreal Engine CSV metric aggregations")
    parser.add_argument("--a", dest="file_a", type=Path, help="Workbook for variant A", required=False)
    parser.add_argument("--b", dest="file_b", type=Path, help="Workbook for variant B", required=False)
    parser.add_argument("--auto", action="store_true", help="Auto-detect 'messungen_auswertung_a.xlsx' and '_b.xlsx' if present")
    parser.add_argument("--out", dest="out", type=Path, default=Path("messungen_ab_vergleich.xlsx"), help="Output workbook path")
    parser.add_argument("--no-recompute-fps", dest="recompute_fps", action="store_false", help="Do NOT recompute FPS from frametime means")
    parser.set_defaults(recompute_fps=True)
    args = parser.parse_args()

    records: List[Dict[str, object]] = []

    single_default = Path("messungen_auswertung.xlsx")
    default_a = Path("messungen_auswertung_a.xlsx")
    default_b = Path("messungen_auswertung_b.xlsx")

    mode = "single"

    if args.file_a or args.file_b:
        if not (args.file_a and args.file_b):
            raise SystemExit("--a and --b must be provided together.")
        records = _collect_from_two_files(args.file_a, args.file_b)
        mode = "two-files-cli"
    elif args.auto and default_a.exists() and default_b.exists():
        records = _collect_from_two_files(default_a, default_b)
        mode = "two-files-auto"
    elif default_a.exists() and default_b.exists():
        # Auto-Erkennung ohne --auto, falls beide exakt so heißen
        records = _collect_from_two_files(default_a, default_b)
        mode = "two-files-autodetect"
    else:
        if not single_default.exists():
            raise FileNotFoundError("No input workbooks found: expected either 'messungen_auswertung.xlsx' OR both 'messungen_auswertung_a.xlsx' & 'messungen_auswertung_b.xlsx'.")
        records = _collect_from_single_file(single_default)
        mode = "single"

    if not records:
        raise RuntimeError("No matching sheets found (Scene{n}_A / Scene{n}_B or Scene{n} in A/B files).")

    # Szenen-Liste
    scenes = sorted(sorted({r["scene"] for r in records}), key=lambda s: int(re.findall(r"\d+", s)[0]))
    by_scene_var: Dict[Tuple[str, str], Dict[str, float]] = {}
    # Optional FPS-Recompute (aus Frametime je Lauf)
    if args.recompute_fps:
        for r in records:
            runs = r.get("runs", {})  # type: ignore
            ft_runs = runs.get("Frametime Ø [ms]", [])  # type: ignore
            if ft_runs:
                fps_runs = [1000.0 / v for v in ft_runs if isinstance(v, (int, float)) and v > 0]
                if fps_runs:
                    new_fps = float(np.mean(fps_runs))
                    orig_fps = r["agg"].get("FPS Ø [#]")  # type: ignore
                    r["agg"]["FPS Ø [#]"] = new_fps  # type: ignore
                    # Optional: Differenzprüfung (still, aber könnte protokolliert werden)
                    r["_fps_diff"] = (orig_fps, new_fps)  # type: ignore
    for r in records:
        by_scene_var[(r["scene"], r["variant"])] = r["agg"]  # type: ignore

    wb = Workbook()
    if wb.active and wb.active.title == "Sheet":
        wb.remove(wb.active)

    ws_sum = wb.create_sheet(title="Gesamtübersicht")  # Keeping German sheet name for continuity
    ws_sum.append(["Scene", "Metric", "A (Ø)", "B (Ø)", "Δ B vs A [%]"])
    for c in range(1, 6):
        ws_sum.cell(row=1, column=c).font = Font(bold=True)
        ws_sum.cell(row=1, column=c).alignment = Alignment(horizontal="center")

    sum_row = 2
    warn_incomplete: List[str] = []
    for scene in scenes:
        agg_A = by_scene_var.get((scene, "A"), {})
        agg_B = by_scene_var.get((scene, "B"), {})
        if not agg_A or not agg_B:
            warn_incomplete.append(scene)
        cmp_df = build_comparison_for_scene(scene, agg_A, agg_B)
        write_scene_sheets(wb, scene, agg_A, agg_B, cmp_df)
        for _, row in cmp_df.iterrows():
            ws_sum.cell(row=sum_row, column=1, value=f"Scene{scene}")
            lab = row["Kennzahl"]  # original German label retained
            ws_sum.cell(row=sum_row, column=2, value=lab)
            ws_sum.cell(row=sum_row, column=3, value=fmt_de(lab, row["A (Ø)"]))
            ws_sum.cell(row=sum_row, column=4, value=fmt_de(lab, row["B (Ø)"]))
            ws_sum.cell(row=sum_row, column=5, value=fmt_de("Δ", row["Δ B vs A [%]"]))
            sum_row += 1
        sum_row += 1

    for col, w in zip(range(1, 6), [14, 24, 16, 16, 16]):
        ws_sum.column_dimensions[get_column_letter(col)].width = w

    out_path = args.out
    wb.save(out_path)
    msg = f"✓ A/B comparison exported ({mode}): {out_path.resolve()}"
    if args.recompute_fps:
        diffs = [r for r in records if r.get("_fps_diff") and all(isinstance(x, (int,float)) for x in r.get("_fps_diff"))]
        if diffs:
            # Prüfe, ob Differenz signifikant (>0.1 fps)
            changed = []
            for r in diffs:
                orig, new = r["_fps_diff"]  # type: ignore
                if orig is not None and abs(orig - new) > 0.1:
                    changed.append(r["scene"]+r["variant"])  # type: ignore
            if changed:
                msg += f" | FPS recomputed from frametime: deviation >0.1 at {', '.join(changed)}"
    if warn_incomplete:
        msg += f" | Incomplete scenes (only one variant): {', '.join('Scene'+s for s in warn_incomplete)}"
    print(msg)

if __name__ == "__main__":
    main()

