"""Compact submission report (<= 20 pages, portrait A4, Times New Roman).

    python scripts/10_build_short_report.py

Writes report/PathSnap_Report_PT-044.docx and (if MS Word is installed) report/PathSnap_Report_PT-044.pdf.
All tables are read from results/; all figures are the PNGs produced by scripts 01-07 and 03b.
The long version (scripts/08 + 09) remains in report/ as extended documentation.
"""
import _bootstrap  # noqa: F401

import json
import re
from pathlib import Path

import pandas as pd
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Inches, Pt, RGBColor

from pathsnap.config import ROOT

RES = ROOT / "results"
REP = ROOT / "report"
OUT_DOCX = REP / "PathSnap_Report_PT-044.docx"
OUT_PDF = REP / "PathSnap_Report_PT-044.pdf"
FONT = "Times New Roman"
TW = 6.27            # text width (in) on A4 portrait with 1-inch margins
BODY = 12
TEXT = json.load(open(REP / "report_text.json", encoding="utf-8"))


# ----------------------------------------------------------------------------- styling helpers
def fontify(run, size=BODY, bold=None, italic=None, color=None):
    run.font.name = FONT
    rpr = run._element.get_or_add_rPr()
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts"); rpr.append(rf)
    for a in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rf.set(qn(a), FONT)
    run.font.size = Pt(size)
    if bold is not None: run.font.bold = bold
    if italic is not None: run.font.italic = italic
    if color: run.font.color.rgb = RGBColor.from_string(color)


def setup(doc):
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
    for s in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
        setattr(sec, s, Cm(2.54))
    st = doc.styles["Normal"]
    st.font.name = FONT; st.font.size = Pt(BODY)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    st.paragraph_format.line_spacing = 1.15
    st.paragraph_format.space_after = Pt(4)
    for name, size, before, after in [("Heading 1", 14, 12, 4), ("Heading 2", 12, 8, 3), ("Heading 3", 11.5, 6, 2)]:
        s = doc.styles[name]
        s.font.name = FONT; s.font.size = Pt(size); s.font.bold = True; s.font.italic = False
        s.font.color.rgb = RGBColor(0, 0, 0)
        rf = s.element.get_or_add_rPr().find(qn("w:rFonts"))
        if rf is None:
            rf = OxmlElement("w:rFonts"); s.element.get_or_add_rPr().append(rf)
        for a in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
            rf.set(qn(a), FONT)
        for a in ("w:asciiTheme", "w:hAnsiTheme", "w:cstheme", "w:eastAsiaTheme"):
            if rf.get(qn(a)) is not None:
                del rf.attrib[qn(a)]
        s.paragraph_format.space_before = Pt(before); s.paragraph_format.space_after = Pt(after)
        s.paragraph_format.keep_with_next = True
    p = sec.footer.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(); fontify(r, 10)
    for kind, txt in (("begin", None), (None, "PAGE"), ("end", None)):
        el = OxmlElement("w:fldChar" if kind else "w:instrText")
        if kind: el.set(qn("w:fldCharType"), kind)
        else: el.set(qn("xml:space"), "preserve"); el.text = txt
        r._r.append(el)


MD = re.compile(r"(\*\*.+?\*\*|\*.+?\*|`.+?`)")


def para(doc, text, size=BODY, align="justify", after=4, italic=False):
    """Paragraph with light markup: **bold**, *italic*, `code`."""
    p = doc.add_paragraph()
    long_tok = any(len(t) > 28 for t in re.sub(r"[*`]", "", text).split())
    p.alignment = {"justify": WD_ALIGN_PARAGRAPH.LEFT if long_tok else WD_ALIGN_PARAGRAPH.JUSTIFY,
                   "center": WD_ALIGN_PARAGRAPH.CENTER, "left": WD_ALIGN_PARAGRAPH.LEFT}[align]
    p.paragraph_format.space_after = Pt(after)
    for part in MD.split(text):
        if not part:
            continue
        if part.startswith("**"):
            fontify(p.add_run(part[2:-2]), size, bold=True, italic=italic or None)
        elif part.startswith("`"):
            fontify(p.add_run(part[1:-1]), size, italic=italic or None, color="1F3864")
        elif part.startswith("*"):
            fontify(p.add_run(part[1:-1]), size, italic=True)
        else:
            fontify(p.add_run(part), size, italic=italic or None)
    return p


def bullets(doc, items, size=BODY):
    for it in items:
        p = para(doc, it, size, after=2)
        p.paragraph_format.left_indent = Cm(0.7); p.paragraph_format.first_line_indent = Cm(-0.45)
        p.runs[0].text = "•  " + p.runs[0].text


def heading(doc, text, level):
    h = doc.add_heading(level=level)
    fontify(h.add_run(text), {1: 14, 2: 12, 3: 11.5}[level], bold=True)


def _cell_borders_none(table):
    tblpr = table._tbl.tblPr
    b = OxmlElement("w:tblBorders")
    for e in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{e}"); el.set(qn("w:val"), "nil"); b.append(el)
    tblpr.append(b)


def table(doc, df, fmt=None, font=None, widths=None, header_map=None):
    """Portrait table: fixed layout, widths from content, font shrinks until it fits."""
    df = df.copy()
    for c, f in (fmt or {}).items():
        if c in df:
            df[c] = df[c].map(lambda v: "N/A" if pd.isna(v) else f.format(v))
    df = df.astype(str).replace({"nan": "N/A"})
    cols = [header_map.get(c, c) if header_map else c for c in df.columns]
    cols = [c.replace("_", " ") for c in cols]
    n = len(cols)
    font = font or (9.5 if n <= 5 else 8.5 if n <= 8 else 7.5)
    if widths is None:
        def need(f):
            ci = f * 0.56 / 72
            mins, extra = [], []
            for j, c in enumerate(cols):
                vals = df.iloc[:, j].tolist()
                word = max(len(w) for w in c.split())
                num = all(re.fullmatch(r"[-+\d.,%NA/ ]*", v) for v in vals)
                vmax = max(len(v) for v in vals) if vals else 3
                m = max(word, vmax if num else min(vmax, 9), 3)
                mins.append(m * ci + 0.16)
                extra.append(0 if num else max(min(sorted(len(v) for v in vals)[int(0.8 * (len(vals) - 1))], 60) - m, 0))
            return mins, extra
        mins, extra = need(font)
        while sum(mins) > TW and font > 6.5:
            font -= 0.5; mins, extra = need(font)
        free = TW - sum(mins)
        if free < 0:
            widths = [m * TW / sum(mins) for m in mins]
        else:
            te = sum(extra)
            widths = [m + (free * e / te if te else free / n) for m, e in zip(mins, extra)]
    t = doc.add_table(rows=len(df) + 1, cols=n)
    t.style = "Table Grid"; t.alignment = WD_TABLE_ALIGNMENT.CENTER; t.autofit = False
    lay = OxmlElement("w:tblLayout"); lay.set(qn("w:type"), "fixed"); t._tbl.tblPr.append(lay)
    for gc, w in zip(t._tbl.tblGrid.findall(qn("w:gridCol")), widths):
        gc.set(qn("w:w"), str(int(w * 1440)))
    for i in range(len(df) + 1):
        trpr = t.rows[i]._tr.get_or_add_trPr()
        cs = OxmlElement("w:cantSplit"); cs.set(qn("w:val"), "true"); trpr.append(cs)
        if i == 0:
            hd = OxmlElement("w:tblHeader"); hd.set(qn("w:val"), "true"); trpr.append(hd)
        for j in range(n):
            cell = t.cell(i, j); cell.width = Inches(widths[j])
            p = cell.paragraphs[0]; p.paragraph_format.space_after = Pt(0); p.paragraph_format.line_spacing = 1.0
            txt = cols[j] if i == 0 else df.iat[i - 1, j]
            fontify(p.add_run(txt), font, bold=(i == 0))
            if i == 0:
                tcpr = cell._tc.get_or_add_tcPr()
                shd = OxmlElement("w:shd"); shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto"); shd.set(qn("w:fill"), "D9E2F3")
                tcpr.append(shd)
    sp = doc.add_paragraph(); sp.paragraph_format.space_after = Pt(2)
    fontify(sp.add_run(""), 4)
    return t


FIGNO = [0]
FIGS: dict = {}      # key -> figure number (filled in pass 1, used for cross-references in pass 2)


def figure(doc, paths, caption, width=None, key=None):
    """One image, or 2 side by side in a borderless table; numbered italic caption."""
    paths = [Path(p) for p in paths if Path(p).exists()]
    if not paths:
        return
    FIGNO[0] += 1
    if key:
        FIGS[key] = FIGNO[0]
    if len(paths) == 1:
        p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.keep_with_next = True; p.paragraph_format.space_after = Pt(1)
        p.add_run().add_picture(str(paths[0]), width=Inches(width or 5.6))
    else:
        t = doc.add_table(rows=1, cols=len(paths)); _cell_borders_none(t)
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        w = (width or TW) / len(paths)
        for j, pth in enumerate(paths):
            c = t.cell(0, j); c.width = Inches(w)
            cp = c.paragraphs[0]; cp.alignment = WD_ALIGN_PARAGRAPH.CENTER; cp.paragraph_format.space_after = Pt(0)
            cp.add_run().add_picture(str(pth), width=Inches(w - 0.08))
    c = para(doc, f"Figure {FIGNO[0]}: {caption}", 9.5, "center", after=6, italic=True)
    return c


def page_break(doc):
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)


# ----------------------------------------------------------------------------- data
st = pd.read_csv(RES / "eda" / "trip_stats.csv", dtype={"trip_id": str})
tr = st[st.partition == "train"]
lab = pd.read_csv(RES / "labels" / "label_stats.csv", dtype={"trip_id": str})
mr = pd.read_csv(RES / "candidates" / "miss_rate_by_partition.csv")
comp = pd.read_csv(RES / "models" / "comparison.csv")
tv = pd.read_csv(RES / "optuna" / "tuned_vs_untuned.csv")
ob = json.load(open(RES / "optuna" / "best_params.json"))
oi = json.load(open(RES / "optuna" / "run_info.json"))
hb = json.load(open(RES / "hmm" / "best_params.json"))
hc = pd.read_csv(RES / "hmm" / "comparison.csv")
dst = pd.read_csv(RES / "hmm" / "per_trip_delta_student_test.csv", dtype={"trip_id": str})
ev = pd.read_csv(RES / "eval" / "eval_results.csv", dtype={"trip_id": str})
NAMES = {"logreg": "Logistic Regression", "gaussian_nb": "Gaussian Naive Bayes", "knn": "K-Nearest Neighbours",
         "random_forest": "Random Forest", "hist_gb": "HistGradientBoosting", "lightgbm": "LightGBM", "xgboost": "XGBoost",
         "mlp": "MLP (128-64)", "baseline_nearest": "Baseline: nearest arc", "baseline_dist_heading": "Baseline: distance+heading"}
M = {r.model: r for r in comp.itertuples()}


def mj(m):
    return json.load(open(RES / "models" / m / "metrics.json"))


# ----------------------------------------------------------------------------- per-model text (Q1/2, 5/6, 9)
MODEL_TEXT = {
    "xgboost": ("Regularised gradient boosting with depth-limited trees captures the interactions between distance, direction of travel and neighbour context, and is fast at inference.",
                "**Strengths:** best student-test Top-1, macro-F1 and log loss; 6 ms per 1000 points. **Limitations:** many interacting hyperparameters; per-point decisions give only 0.81 route continuity.",
                "Binary log loss + L2 leaf regularisation, second-order boosting.",
                "Train-fold Top-1 0.950 vs CV 0.888 vs ST 0.899. Its training curve (Figure {xgb_hist}) shows the student-test loss flattening at ~0.083 after ~150 trees while the training loss keeps falling - mild over-fitting in late trees, but the held-out loss never rises."),
    "lightgbm": ("Leaf-wise boosting with row/column sub-sampling; efficient on 1 M rows and well calibrated - useful for the HMM.",
                 "**Strengths:** best CV Top-1 (0.889), 18 s training. **Limitations:** 63-leaf trees can over-fit; ~4x slower inference than XGBoost.",
                 "Binary log loss with L2 (reg_lambda = 1), learning rate 0.05.",
                 "Train-fold 0.952 vs CV 0.889 vs ST 0.891; held-out loss plateaus after ~120 iterations while training loss keeps falling - mild, harmless over-fitting."),
    "random_forest": ("Bagged trees model non-linear rules (e.g. heading matters only when moving) without scaling; averaging reduces variance.",
                      "**Strengths:** robust, CV Top-1 0.887. **Limitations:** slowest ensemble to train (94 s), 332 MB model, trees memorise training trips.",
                      "Gini impurity per split; probability = mean leaf frequency over 150 trees.",
                      "Largest generalisation gap: train-fold 0.972 vs CV 0.887 (8.5 pp), log loss 0.040 vs 0.095 - over-fits the training trips, but bagging keeps held-out accuracy high."),
    "hist_gb": ("Histogram gradient boosting: a strong, fast default for tabular data with mixed-scale features.",
                "**Strengths:** ST Top-1 0.890 with 15 s training. **Limitations:** without early stopping its probabilities are less calibrated (log loss 0.097).",
                "Binary log loss, Newton leaf values, 300 iterations at learning rate 0.1.",
                "Train-fold 0.935 vs CV 0.882 vs ST 0.890 - moderate expected gap, no harmful over-fitting."),
    "mlp": ("A feed-forward network can learn arbitrary non-linear feature combinations and outputs smooth probabilities.",
            "**Strengths:** ST Top-1 0.892, fastest inference (3 ms/1k points). **Limitations:** slowest to train (203 s on an 80k-observation sub-sample), higher fold variability (std 0.044).",
            "Binary cross-entropy with Adam (lr 0.001, batch 512, L2 1e-4), early stopping.",
            "Train-fold 0.917 vs CV 0.881 vs ST 0.892. On a trip-grouped split (Figure {mlp_loss}) validation loss is best at epoch 8 (0.112) and then fluctuates while training loss keeps falling - over-fitting after ~8 epochs, which early stopping prevents."),
    "logreg": ("Linear, well-calibrated baseline: the strongest signals (distance, heading, neighbour agreement) are roughly monotonic.",
               "**Strengths:** fast (6 s on 1.08 M rows), stable, interpretable. **Limitations:** cannot model feature interactions; only matches the hand-made distance+heading rule.",
               "Binary log loss with L2 (C = 1), L-BFGS; scaler fitted on training folds only.",
               "Train-fold 0.886 vs CV 0.868 (1.7 pp) - no over-fitting but clear under-fitting (bias): it stays far below the tree models even on training trips."),
    "knn": ("Instance-based: a (point, candidate) pair should share the label of similar stored pairs; no parametric form.",
            "**Strengths:** simple, smooth distance-weighted probabilities. **Limitations:** 684 ms per 1000 points (~120x XGBoost), so the reference set was sub-sampled to 20k observations; all 33 features weigh equally.",
            "No training loss; the 25 nearest neighbours vote with weight 1/distance (scaled features).",
            "Train-fold 0.891 vs CV 0.864 vs ST 0.862 - modest gap; under-fits relative to trees because of the sub-sample and uninformative dimensions."),
    "gaussian_nb": ("Simplest generative probabilistic model; tests whether per-class feature distributions are enough.",
                    "**Strengths:** trains in ~1 s. **Limitations:** features are strongly correlated (d, log_d, rank, context distances), violating independence, so evidence is double-counted.",
                    "Maximum-likelihood per-class Gaussians; Bayes rule.",
                    "Clear under-fitting: train-fold Top-1 only 0.839, log loss ~0.90 even on training trips; precision 0.56 vs recall 0.92 and the highest fold variability (std 0.067)."),
}


# ----------------------------------------------------------------------------- document
def build():
    doc = Document(); setup(doc)
    doc.core_properties.title = "PathSnap Technical Report - Team PT-044"; doc.core_properties.author = "Team PT-044"

    # ---------------- title page
    for _ in range(5): doc.add_paragraph()
    para(doc, "PES UNIVERSITY", 14, "center", 2).runs[0].bold = True
    para(doc, "Department of Computer Science and Engineering", 12, "center", 2)
    para(doc, "UE24CS352A: Machine Learning - Hackathon", 12, "center", 24)
    t = para(doc, "PathSnap: Distributed Map-Matching and Route Reconstruction", 20, "center", 6); t.runs[0].bold = True
    para(doc, "Technical Report", 14, "center", 30)
    para(doc, f"**Team {TEXT['team']['id']}**", 12, "center", 6)
    roster = pd.DataFrame(TEXT["team"]["roster"], columns=["SRN", "Name"])
    table(doc, roster, font=11, widths=[2.2, 2.6])
    para(doc, "Repository: " + TEXT["team"]["repo"], 11, "center", 4)
    page_break(doc)

    # ---------------- 1 overview
    heading(doc, "1. Problem and Approach", 1)
    para(doc, "Vehicle GPS fixes are noisy (multipath, urban canyons, dropouts), so a raw track drifts onto parallel streets and across "
              "buildings. PathSnap recovers the road segments the vehicle actually used: (1) **candidate generation** - road arcs within "
              "50 m of each GPS point; (2) a **classifier** that scores each (point, candidate) pair from 33 engineered features; "
              "(3) **HMM/Viterbi decoding** that combines the classifier probabilities (emissions) with road-network connectivity "
              "(transitions) into one connected route. A FastAPI service (Laptop 1) runs the pipeline and scores it server-side; a Streamlit "
              "UI (Laptop 2) lets the evaluator choose a trip. Partitions: 80 train trips (fitting, CV, tuning), 10 student-test trips "
              "(model comparison), 10 evaluator trips (final evaluation only).")

    # ---------------- 2 EDA
    heading(doc, "2. Dataset Understanding and Exploratory Data Analysis", 1)
    para(doc, "*Partition used: all statistics, plots and design decisions use the train partition (80 trips); student-test was only "
              "inventoried; the evaluator data was not opened during development.* Code: `scripts/01_eda.py`, `02a_candidate_grid.py`, "
              "`02_build_dataset.py`.")
    heading(doc, "2.1 Data organisation and quality", 2)
    para(doc, "Each trip has four header-less files: `.track` (lon, lat, t; float64), `.nodes` (lon, lat; row = node id), `.arcs` "
              "(from_node, to_node; row = arc id, directed; int64) and `.route` (ordered reference arc ids; int64). Every trip has its "
              f"**own road network** (649 - 307 392 arcs, different cities/countries). Train trips contain "
              f"{int(tr.clean_points.sum())} GPS points (median {int(tr.clean_points.median())} per trip, median extent "
              f"{tr.extent_x_km.median():.1f} x {tr.extent_y_km.median():.1f} km).")
    miss = pd.read_csv(RES / "eda" / "missing_values.csv")
    mt = miss.pivot(index="partition", columns="file", values="total_rows").reset_index()
    mt["missing / NaN cells"] = miss.groupby("partition").missing_cells.sum().values
    mt["files missing"] = 0
    para(doc, "**Missing-value analysis** (rows per file type, all files present):", after=2)
    table(doc, mt, font=9.5)
    para(doc, "No NaN/invalid values or duplicate timestamps; all tracks are time-sorted and lie inside their network's bounding box. "
              f"Issues found: {int((tr.gaps_gt_30s > 0).sum())} trips with sampling gaps > 30 s (largest 153 064 s), "
              f"{int(tr.n_speed_outliers.sum())} implied-speed outliers > 70 m/s, {int(tr.n_self_loops.sum())} self-loops, "
              f"{int(tr.n_duplicate_arcs.sum())} duplicate and {int(tr.n_zero_length_arcs.sum())} zero-length arcs, networks split into "
              f"up to {int(tr.weak_components.max())} components, and reference routes covering only part of a trip "
              "(train 00000054: 89 % of points > 40 m from the route).")
    heading(doc, "2.2 Pre-processing", 2)
    bullets(doc, [
        "**Cleaning:** drop non-finite/out-of-range rows, sort by time, drop duplicate timestamps (kept for robustness on unseen trips).",
        "**Projection:** lon/lat are angles, so each trip is projected to metres with a local equirectangular projection; largest "
        "single-step error vs haversine 0.72 m - far below GPS noise.",
        "**Per-point labels:** `.route` lists arcs in order (≈5 GPS points per arc), not per point. Each point is aligned to a route "
        "position by a monotone dynamic programme (forward-only, resolves loops) minimising distance + a heading penalty; points > 40 m "
        f"away are *unmatched* (-1). Matched: {lab[lab.partition == 'train'].n_matched.sum() / lab[lab.partition == 'train'].n_points.sum():.1%} "
        f"of train points (pooled; 94.5 % of the unmatched points come from 4 partial-route trips), {lab[lab.partition == 'student_test'].n_matched.sum() / lab[lab.partition == 'student_test'].n_points.sum():.1%} of student-test points.",
        "**Candidates:** KD-tree over arcs densified every 10 m; K nearest arcs within R metres. R = 50 m, K = 8 were chosen on train "
        f"(cheapest within 0.5 pp of the lowest miss rate): candidate miss rate {mr.miss_rate.iloc[0]:.2%} (train), "
        f"{mr.miss_rate.iloc[1]:.2%} (student-test), ~5.7 candidates per point.",
    ])
    heading(doc, "2.3 Patterns observed and features extracted", 2)
    para(doc, "Key patterns: (1) GPS noise is small but heavy-tailed (median 4.5 m to the route, 95th percentile 12-30 m); (2) ~95 % of "
              "arcs have a reverse twin with identical geometry, so the nearest-arc rule reaches only 0.54 Top-1 - direction of travel "
              "is essential; (3) long time gaps and disconnected networks must be handled by the decoder; (4) neighbouring points are "
              "strongly correlated, motivating context features and trip-grouped CV. We extracted **33 location-independent features** "
              "per (point, candidate) - no absolute coordinates (trips are in different countries) and nothing from the route "
              "(a unit test deletes `.route` and checks the features are unchanged):")
    feats = pd.DataFrame([
        ("Distance (7)", "d, log_d, d_minus_min, d_ratio, rank, softmin, gap12", "the true arc is usually among the closest; relative forms work in dense and sparse areas"),
        ("Heading (3)", "cos_dh, abs_dh, heading_valid", "only direction of travel separates two-way twins"),
        ("Position (4)", "u, dist_to_start, dist_to_end, log_arc_len", "points at a segment end often belong to the next arc"),
        ("Topology (5)", "out_deg_to, in_deg_from, is_two_way, n_cand, n_within_10m", "junction structure and ambiguity of the candidate set"),
        ("Context (10)", "ctx_d_(+-1,3,5), ctx_d_mean, ctx_conn_frac, is_nearest_prev/next", "neighbouring fixes on (or connected to) the same arc"),
        ("Kinematics (4)", "speed, dt_prev, dt_next, turn_rate", "heading reliability, gaps and turns"),
    ], columns=["Group", "Features", "Why"])
    table(doc, feats, font=9, widths=[1.05, 2.6, 2.62])
    figure(doc, [RES / "eda" / "trajectories_train.png"], "GPS trajectory plots of six train trips in projected metres (colour = time).", 6.0)
    figure(doc, [RES / "eda" / "network_00000026.png"],
           "Road-network visualisation around train trip 00000026: network (grey), reference route (green), GPS (blue).", 4.6)
    figure(doc, [RES / "eda" / "gps_route_dist_train.png"],
           "Distance of GPS points to the reference route (train, log scale; the bar at 100 m collects partial-route trips).", 4.4)
    figure(doc, [RES / "eda" / "dt_hist_train.png"], "Sampling interval between consecutive GPS points (train, log-log).", 4.0)
    figure(doc, [RES / "candidates" / "grid.png"],
           "Candidate miss rate (left) and candidate-set size (right) for radius R and K nearest arcs (train).", 6.2)
    figure(doc, [REP / "img" / "eda_map_student_test_00000000.png"],
           "Folium map (student-test 00000000): raw GPS (blue), road network (grey), reference route (green), layer control and legend.", 5.6)

    # ---------------- 3 models
    heading(doc, "3. Model Experimentation and Comparison", 1)
    para(doc, "Task: binary classification of each (GPS point, candidate arc) row; the highest-probability candidate is the prediction "
              "for the point. **Protocol:** 5-fold GroupKFold by trip on the 80 train trips, then refit on all train trips and evaluation "
              "on the 10 student-test (ST) trips; seed 42. **Primary metric: CV Top-1** (share of GPS points whose top candidate is the "
              "reference arc) because it is the actual task and is measured on unseen trips; log loss breaks ties because the HMM uses "
              "the probabilities. Route continuity = share of consecutive arc changes that are physically adjacent (before HMM). "
              "Code: `scripts/03_train_compare.py`.")
    heading(doc, "3.1 Overall model comparison", 2)
    tab = comp.copy()
    tab["Model"] = tab.model.map(NAMES)
    tab["Family"] = tab.family.str.replace(" (generative)", "", regex=False)
    tab["Key configuration"] = tab.model.map({
        "lightgbm": "400 trees, lr 0.05, 63 leaves", "xgboost": "400 trees, lr 0.05, depth 7", "random_forest": "150 trees, min leaf 5",
        "hist_gb": "300 iter, lr 0.1, 31 leaves", "mlp": "128-64 ReLU, Adam, early stop", "logreg": "C = 1, L2, scaled",
        "baseline_dist_heading": "-d + 10 cos(heading diff)", "knn": "k = 25, distance-weighted", "gaussian_nb": "var_smoothing 1e-9",
        "baseline_nearest": "nearest arc"})
    tab = tab[["Model", "Family", "Key configuration", "cv_macro_f1", "cv_top1", "st_route_continuity", "st_macro_f1",
               "st_top1", "st_log_loss", "inference_ms_per_1k_points"]]
    table(doc, tab, {c: "{:.3f}" for c in ["cv_macro_f1", "cv_top1", "st_route_continuity", "st_macro_f1", "st_top1", "st_log_loss"]}
          | {"inference_ms_per_1k_points": "{:.1f}"},
          header_map={"cv_macro_f1": "CV macro-F1", "cv_top1": "CV Top-1", "st_route_continuity": "Route contin.",
                      "st_macro_f1": "ST macro-F1", "st_top1": "ST Top-1", "st_log_loss": "Log loss",
                      "inference_ms_per_1k_points": "Infer. ms/1k pts"})
    para(doc, "Log loss is N/A for the two rule-based baselines (no probabilities). Inference = classifier time per 1000 GPS points.", 9.5)
    heading(doc, "3.2 Per-model answers", 2)
    for i, m in enumerate(["xgboost", "lightgbm", "random_forest", "hist_gb", "mlp", "logreg", "knn", "gaussian_nb"], 1):
        j = mj(m); s = j["student_test"]; why, strength, loss, fit = MODEL_TEXT[m]
        fit = fit.format(xgb_hist=FIGS.get("xgb_hist", "?"), mlp_loss=FIGS.get("mlp_loss", "?"))
        heading(doc, f"3.2.{i} {NAMES[m]} ({j['family']})", 3)
        cfg = ", ".join(f"{k} = {v}" for k, v in j["config"].items())
        para(doc, f"**A. Why suitable:** {why} {strength}", BODY, after=1)
        para(doc, f"**B. Hyperparameters (initial):** {cfg}. **Loss:** {loss}", BODY, after=1)
        para(doc, f"**C. Student-test:** Top-1 {s['top1']:.4f}, macro-F1 {s['macro_f1']:.4f} (P {s['precision']:.3f}, "
                  f"R {s['recall']:.3f}), log loss {s['log_loss']:.4f}; CV Top-1 {j['cv']['top1'][0]:.4f} ± {j['cv']['top1'][1]:.4f}; "
                  f"training {j['fit_time_full_s']:.1f} s. **Fit:** {fit}", BODY, after=3)
    heading(doc, "3.3 Visualisations and model selection", 2)
    figure(doc, [RES / "models" / "cv_macro_f1.png"], "Cross-validation macro-F1 of all models (5-fold GroupKFold, mean ± std).", 6.0)
    figure(doc, [RES / "models" / "student_test_top1_f1.png"], "Student-test Top-1 accuracy and macro-F1 of all models.", 6.0)
    figure(doc, [RES / "models" / "log_loss.png"], "Log loss on CV and student-test (learned models; N/A for baselines).", 5.2)
    figure(doc, [RES / "models" / "generalisation_top1.png"],
           "Top-1 on training-fold trips vs held-out CV vs student-test (over/under-fitting check).", 6.0)
    figure(doc, [RES / "models" / "xgboost" / "training_history.png"],
           "XGBoost per-iteration log loss (student-test passed only to record the curve, not for any decision).", 4.6, key="xgb_hist")
    figure(doc, [RES / "models" / "mlp" / "loss_curve_grouped.png"],
           "MLP epoch-wise training vs validation loss on a trip-grouped split (64 / 16 train trips).", 4.6, key="mlp_loss")
    para(doc, "Only XGBoost, LightGBM (boosting iterations) and the MLP (epochs) produce real training histories, so no training "
              "curves are shown for the other models. Confusion matrices for all models are in `results/models/<model>/`.", 10.5)
    para(doc, "**Selection:** XGBoost and LightGBM tie on CV Top-1 (0.888 vs 0.889, within the fold std of ~0.04). XGBoost was chosen "
              "because it is best on every student-test accuracy metric (Top-1 0.899, macro-F1 0.937, log loss 0.082 - the best input "
              "for the HMM) and ~4x faster at inference. Random Forest over-fits most, the MLP is less stable, and the linear, "
              "instance-based and Naive Bayes models are clearly worse. All learned ensembles beat the distance+heading rule by "
              "2-4 pp and the nearest-arc rule by > 25 pp.")

    # ---------------- 4 Bayesian optimisation
    heading(doc, "4. Bayesian Hyperparameter Optimisation (XGBoost)", 1)
    ss = json.load(open(RES / "optuna" / "search_space.json"))
    para(doc, "**1-2. Search space:** " + "; ".join(f"{k} {v}" for k, v in ss.items()) + ". "
              f"**3. Objective:** {oi['objective']} - the same folds as Section 3; student-test never used in the search. "
              f"Optuna TPE sampler (seed 42, 10 random start-up trials) with median pruning. **4. Trials:** {oi['n_trials_total']} "
              f"({oi['n_complete']} completed, {oi['n_pruned']} pruned), {oi['search_wall_time_s'] / 60:.0f} min on a laptop CPU. "
              "**5. Best:** " + ", ".join(f"{k} = {v:.4g}" if isinstance(v, float) else f"{k} = {v}" for k, v in ob["params"].items())
              + f" (CV Top-1 {ob['best_cv_top1']:.4f}).")
    table(doc, tv, {c: "{:.4f}" for c in tv.columns if c not in ("model", "fit_time_s", "inference_ms_per_1k_points")}
          | {"fit_time_s": "{:.1f}", "inference_ms_per_1k_points": "{:.1f}"},
          header_map={"cv_top1": "CV Top-1", "cv_macro_f1": "CV macro-F1", "cv_log_loss": "CV log loss", "st_top1": "ST Top-1",
                      "st_macro_f1": "ST macro-F1", "st_log_loss": "ST log loss", "st_route_continuity": "ST contin.",
                      "fit_time_s": "Fit s", "inference_ms_per_1k_points": "Infer. ms/1k"})
    para(doc, "**6. Did tuning help?** On the optimised objective slightly: CV Top-1 0.8879 -> 0.8907, CV macro-F1 0.9301 -> 0.9312. All "
              "23 completed trials lie within 0.881-0.891, so XGBoost is insensitive to its hyperparameters here and the gain is "
              "smaller than the fold std (0.037). As a stand-alone classifier the tuned model was *not* better on student-test (Top-1 "
              "0.894 vs 0.899, log loss 0.087 vs 0.082) and is ~4.5x slower; scale_pos_weight = 2.31 improves ranking but hurts "
              "calibration. Because the deployed system is classifier + HMM, we compared both models after decoding on train "
              "out-of-fold predictions (no student-test or evaluator data): 0.9361 (tuned) vs 0.9360 (untuned) - a tie - and on "
              "student-test 0.9513 vs 0.9383, so the tuned model is deployed.")
    figure(doc, [RES / "optuna" / "optimisation_history.png"],
           "Optuna TPE optimisation history: CV Top-1 per trial and best so far (crosses = pruned trials).", 5.0)
    figure(doc, [RES / "optuna" / "tuned_vs_untuned.png"], "Untuned vs tuned XGBoost on the same metrics and partitions.", 6.0)

    # ---------------- 5 HMM
    heading(doc, "5. HMM and Viterbi Route Reconstruction", 1)
    bullets(doc, [
        "**1. Hidden states:** the candidate arcs of each GPS point (≤ 8 within 50 m, plus ties) with the point's projection u on the "
        "arc; candidate rows keep the classifier's order, so every state maps to one arc id and one probability.",
        "**2. Emissions:** log P(o_t | c) = w_e · log p_c, with p_c the XGBoost probability re-normalised over the point's candidates.",
        "**3. Transitions** (Newson-Krumm): log P(a -> b) = -(|d_net - d_gc| + U-turn penalty) / beta, where d_gc is the straight-line "
        "distance between consecutive fixes and d_net the exact driving distance between the two projections (same arc: (u_b - u_a)·len; "
        "otherwise rest of a + bounded-Dijkstra shortest path + start of b). Unreachable pairs get probability 0. Viterbi: "
        "δ_t(b) = log P(o_t|b) + max_a[δ_t-1(a) + log P(a->b)] with back-pointers, in log space.",
        "**4. Missing candidates / gaps:** a point without candidates (none within 200 m) or a time gap > gap_s closes the current chain; "
        "decoding restarts from the emissions at the next point.",
        "**5. Disconnected segments:** if no transition is feasible (disconnected components, GPS jumps) the chain is split instead of "
        "failing; consecutive non-adjacent decoded arcs are joined by the shortest path (≤ 3 km) when the route is stitched.",
    ], BODY)
    para(doc, f"**Decoder tuning** on train out-of-fold probabilities (216-combination grid): beta = {hb['params']['beta']} m, U-turn "
              f"penalty {hb['params']['uturn_m']} m, w_e = {hb['params']['w_e']}, gap_s = {hb['params']['gap_s']} s. "
              "**6. Classifier-only vs classifier + HMM:**", BODY)
    h = hc[["partition", "method", "top1", "strict_accuracy", "direct_continuity", "plausible_continuity", "unbridged_gaps",
            "route_precision_mean", "route_f1_mean"]]
    table(doc, h, {c: "{:.4f}" for c in ["top1", "strict_accuracy", "direct_continuity", "plausible_continuity",
                                          "route_precision_mean", "route_f1_mean"]},
          header_map={"top1": "Top-1", "strict_accuracy": "Strict acc.", "direct_continuity": "Direct contin.",
                      "plausible_continuity": "Plausible contin.", "unbridged_gaps": "Unbridged gaps",
                      "route_precision_mean": "Route precision", "route_f1_mean": "Route F1"})
    para(doc, "HMM decoding raised student-test Top-1 from 0.894 to **0.951** (+5.7 pp) and train-OOF Top-1 from 0.890 to 0.936; "
              "plausible continuity (next arc reachable within 2·d_gc + 50 m) rose from 0.83 to 0.99, unbridged route gaps fell from "
              "72 to 3, and route F1 from 0.81 to 0.96, mainly by removing spurious arcs (route precision 0.69 -> 0.93). Tuning the "
              "decoder added 1.4 pp over the defaults. **Where it did not help:** it improved 9 of 10 student-test trips (up to "
              "+16.9 pp on 00000077) but hurt 00000065 (-3.5 pp), where the chain locked onto the parallel carriageway of a divided "
              "road; on train the worst case is 00000010 (0.683 -> 0.551): slow, many one-way arcs and an already weak classifier "
              "(route F1 still improved 0.37 -> 0.69). Code: `src/pathsnap/hmm.py`, `scripts/05_hmm_eval.py`.", BODY)
    figure(doc, [RES / "hmm" / "clf_vs_hmm_student_test.png"], "Classifier-only vs classifier + HMM/Viterbi on student-test.", 6.0)
    figure(doc, [RES / "hmm" / "per_trip_scatter_student_test.png"],
           "Per-trip Top-1 on student-test; points above the diagonal are trips where HMM decoding helps.", 3.4)

    # ---------------- 6 Folium
    heading(doc, "6. Geographical Visualisation (Folium)", 1)
    para(doc, "`scripts/07_make_maps.py` (code in `src/pathsnap/viz.py`) writes interactive maps to `results/maps/` with toggleable "
              "layers and a legend: road network (grey), reference route (green), classifier-only arcs (orange), predicted HMM route "
              "(red), GPS observations (blue, start/end markers) and wrongly matched points (purple), fitted to the trip extent. "
              "Best trip 00000007: Top-1 99.8 %, only 3 wrong points. Worst trip 00000065 (Figure below): the wrong points form one "
              "run along a divided road where the predicted route follows the parallel carriageway.", BODY)
    figure(doc, [RES / "maps" / "png" / "result_00000065.png"],
           "Folium map of student-test trip 00000065 (worst case: Top-1 87.8 %, route F1 88.7 %).", 5.6)

    # ---------------- 7 API / UI
    heading(doc, "7. API, UI, Deployment and Final Evaluation", 1)
    para(doc, "**Laptop 1** runs FastAPI (`api/main.py`, `uvicorn --host 0.0.0.0 --port 8000`): `GET /health`, `GET /trips` (ids only), "
              "`POST /predict {trip_id}` -> full pipeline (load, project, candidates, features, XGBoost, HMM/Viterbi, stitching). "
              "Only after predicting is the `.route` file read server-side to compute metrics; the response contains GPS points, the "
              "predicted route and aggregate metrics - never the reference route. Every request is logged (client, trip, counts, stage "
              "timings, metrics). **Laptop 2** runs the Streamlit UI (`ui/app.py`), which talks only to the API: it lists the trips, "
              "sends the selected id and draws the predicted route on a Folium map with the metrics. The exported pipeline reproduced "
              "the offline student-test Top-1 exactly on all 10 trips.", BODY)
    para(doc, "**Final run on the evaluator trips.** The model and all parameters were frozen before the evaluator data arrived "
              "(artifacts written 14:30, data received 14:59); the trips were only passed through the API:", BODY, after=2)
    e = ev[["trip_id", "n_points", "top1_classifier_only", "top1_hmm", "plausible_continuity", "route_f1", "api_total_s"]]
    table(doc, e, {"n_points": "{:.0f}", "top1_classifier_only": "{:.4f}", "top1_hmm": "{:.4f}", "plausible_continuity": "{:.4f}",
                   "route_f1": "{:.4f}", "api_total_s": "{:.2f}"},
          header_map={"trip_id": "Trip", "n_points": "GPS points", "top1_classifier_only": "Top-1 classifier only",
                      "top1_hmm": "Top-1 with HMM", "plausible_continuity": "Continuity", "route_f1": "Route F1",
                      "api_total_s": "API time (s)"})
    para(doc, "Pooled Top-1 is **0.960** (classifier alone 0.847), in line with the student-test estimate (0.951). On trip 00000015, "
              "78 % of the classifier's errors were the right road in the wrong direction; HMM decoding lifts it from 0.655 to 0.953.", BODY)
    shots = [p for p in (REP / "img" / "eval_ui.png", REP / "img" / "eval_api_console.png") if p.exists()]
    if shots:
        figure(doc, shots, "Final evaluation: UI on Laptop 2 and API console on Laptop 1 (receiving laptop).")
    else:
        figure(doc, [REP / "img" / "eval_local_ui.png"], "UI showing evaluator trip 00000015 served by the API on the frozen model.", 5.2)
        para(doc, "[Insert screenshot of the API console on Laptop 1 during the final evaluation: save it as "
                  "report/img/eval_api_console.png and re-run scripts/10_build_short_report.py.]", 10, "left").runs[0].font.color.rgb = RGBColor(0xC0, 0, 0)

    # ---------------- 8 conclusion
    heading(doc, "8. Conclusions, Limitations and Disclosure", 1)
    para(doc, "Gradient-boosted trees on location-independent features give a strong per-point classifier (0.89-0.90 Top-1), direction "
              "of travel is the key feature on two-way roads, and sequential HMM decoding contributes more than hyperparameter tuning "
              "(0.951 student-test, 0.960 evaluator). **Limitations:** labels are derived from the reference route and inherit its "
              "errors; arcs are straight segments; parallel one-way carriageways remain the main failure mode; tuning used 30 trials "
              "with marginal gains; 10 student-test trips cannot resolve differences below ~1 pp. **Reproducibility:** all numbers come "
              "from `results/` produced by `scripts/01-07`; the README lists the exact commands. **AI assistance:** " +
              TEXT["ai_disclosure"], BODY)
    doc.save(OUT_DOCX)
    return OUT_DOCX


def to_pdf(docx_path, pdf_path):
    try:
        import win32com.client
    except ImportError:
        return None
    word = win32com.client.Dispatch("Word.Application"); word.Visible = False
    try:
        d = word.Documents.Open(str(docx_path), False, True)
        pages = d.ComputeStatistics(2)
        d.SaveAs(str(pdf_path), FileFormat=17); d.Close(0)
        return pages
    finally:
        word.Quit()


if __name__ == "__main__":
    build()               # pass 1: assign figure numbers
    FIGNO[0] = 0
    p = build()           # pass 2: cross-references resolved
    pages = to_pdf(p, OUT_PDF)
    print("saved", p, "| pages:", pages, "| pdf:", OUT_PDF if pages else "not produced")
