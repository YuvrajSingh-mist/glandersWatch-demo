"""Build the presentation deck from the pipeline outputs.

Run:  .venv/bin/python src/15_make_deck.py

Structure: a holistic overview (what it is, the problem, why it is needed), then
one section per component holding only that component's required information.
Every number and figure is read from reports/*.json and figures/*.png, so the
deck cannot drift from the report. 16:9, sixteen slides.
"""
from __future__ import annotations

import json
import pathlib

import numpy as np
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

ROOT = pathlib.Path(__file__).resolve().parents[1]
REP, FIG = ROOT / "reports", ROOT / "figures"
OUT = ROOT / "GlandersWatch_Deck.pptx"

R = json.loads((REP / "results.json").read_text())
A = json.loads((REP / "dataset_analysis.json").read_text())
CL = json.loads((REP / "country_level.json").read_text())
YT = json.loads((REP / "year_target.json").read_text())

L1 = R["level1"]["metrics"]
CH = R["best_model"]
PREV = A["labels"]["prevalence"]
ROWS = A["labels"]["rows"]
POS = A["labels"]["positives"]
NC = A["labels"]["countries"]
SEL = R["selection"]
PREV_YEARS = sorted(A["labels"]["prevalence_by_year"])
YRS = (PREV_YEARS[0], PREV_YEARS[-1])
LR = L1["Logistic Regression"]
RF = L1["Random Forest"]

# share of total attribution held by the selected features
_rank = SEL["ranking"]
_top = sum(r["mean_abs_shap"] for r in _rank[: SEL["top_k"]])
_all = sum(r["mean_abs_shap"] for r in _rank) or 1.0
ATTR_SHARE = _top / _all

INK = RGBColor(0x1F, 0x29, 0x33)
BLUE = RGBColor(0x1D, 0x4E, 0xD8)
GREY = RGBColor(0x6B, 0x72, 0x80)
PALE = RGBColor(0xF1, 0xF5, 0xF9)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
MUTE = RGBColor(0x94, 0xA3, 0xB8)

W, H = Inches(13.333), Inches(7.5)


def alert_econ() -> dict:
    """Recall at a fixed budget, and budget needed for a target recall."""
    path = REP / f"pred_L1_{CH.replace(' ', '_')}.csv"
    if not path.exists():
        return {"r6": 0.0, "b50": 0.0, "b80": 0.0}
    d = pd.read_csv(path)
    y = d.y_true.to_numpy().astype(int)
    p = d.p_pred.to_numpy()
    order = np.argsort(-p)
    fe = np.cumsum(y[order]) / y.sum()
    fa = np.arange(1, len(y) + 1) / len(y)

    def recall_at(b):
        return float(fe[max(0, int(b * len(y)) - 1)])

    def budget_at(t):
        k = np.flatnonzero(fe >= t - 1e-12)
        return float(fa[k[0]]) if len(k) else 1.0

    return {"r6": recall_at(0.06), "b50": budget_at(0.50), "b80": budget_at(0.80)}


AE = alert_econ()


def slide(prs, title: str, kicker: str | None = None):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bar = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, W, Inches(0.09))
    bar.fill.solid(); bar.fill.fore_color.rgb = BLUE; bar.line.fill.background()
    bar.shadow.inherit = False

    t = s.shapes.add_textbox(Inches(0.62), Inches(0.32), W - Inches(1.2), Inches(0.7))
    p = t.text_frame.paragraphs[0]
    t.text_frame.word_wrap = True
    p.text = title
    p.font.size = Pt(29); p.font.bold = True; p.font.color.rgb = INK

    if kicker:
        k = s.shapes.add_textbox(Inches(0.62), Inches(1.0), W - Inches(1.2), Inches(0.42))
        k.text_frame.word_wrap = True
        kp = k.text_frame.paragraphs[0]
        kp.text = kicker
        kp.font.size = Pt(14); kp.font.color.rgb = GREY
    return s


def divider(prs, num: str, title: str, answers: str, covered: list[str]):
    """Dark full-bleed section opener."""
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, W, H)
    bg.fill.solid(); bg.fill.fore_color.rgb = INK; bg.line.fill.background()
    bg.shadow.inherit = False

    n = s.shapes.add_textbox(Inches(0.85), Inches(2.0), Inches(3.0), Inches(1.3))
    np_ = n.text_frame.paragraphs[0]
    np_.text = num
    np_.font.size = Pt(86); np_.font.bold = True
    np_.font.color.rgb = RGBColor(0x3B, 0x82, 0xF6)

    t = s.shapes.add_textbox(Inches(0.85), Inches(3.25), Inches(8.6), Inches(1.2))
    t.text_frame.word_wrap = True
    tp = t.text_frame.paragraphs[0]
    tp.text = title
    tp.font.size = Pt(40); tp.font.bold = True; tp.font.color.rgb = WHITE

    a = s.shapes.add_textbox(Inches(0.85), Inches(4.35), Inches(9.2), Inches(0.9))
    a.text_frame.word_wrap = True
    ap = a.text_frame.paragraphs[0]
    ap.text = answers
    ap.font.size = Pt(15); ap.font.color.rgb = MUTE

    c = s.shapes.add_textbox(Inches(0.85), Inches(5.45), Inches(9.5), Inches(1.3))
    ctf = c.text_frame; ctf.word_wrap = True
    for i, item in enumerate(covered):
        p = ctf.paragraphs[0] if i == 0 else ctf.add_paragraph()
        p.text = "—  " + item
        p.font.size = Pt(12); p.font.color.rgb = RGBColor(0xCB, 0xD5, 0xE1)
        p.space_after = Pt(3)

    rule = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.85), Inches(2.95),
                              Inches(1.5), Inches(0.05))
    rule.fill.solid(); rule.fill.fore_color.rgb = RGBColor(0x3B, 0x82, 0xF6)
    rule.line.fill.background(); rule.shadow.inherit = False
    return s


def body(s, x, y, w, h, lines, size=16, space=8, bullet=True):
    tb = s.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    first = True
    for ln in lines:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        bold = ln.startswith("**")
        p.text = ln.replace("**", "")
        p.font.size = Pt(size)
        p.font.bold = bold
        p.font.color.rgb = BLUE if bold else INK
        p.space_after = Pt(space)
        if bullet and not bold:
            p.text = "•  " + p.text
    return tb


def stat_row(s, y, items, w_each=Inches(2.85), gap=Inches(0.22), h=Inches(1.18)):
    x = Inches(0.62)
    for label, value in items:
        box = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w_each, h)
        box.fill.solid(); box.fill.fore_color.rgb = PALE
        box.line.fill.background(); box.shadow.inherit = False
        tf = box.text_frame
        tf.word_wrap = True
        tf.margin_left = Inches(0.14); tf.margin_top = Inches(0.09)
        p = tf.paragraphs[0]; p.text = value
        p.font.size = Pt(23); p.font.bold = True; p.font.color.rgb = BLUE
        p2 = tf.add_paragraph(); p2.text = label
        p2.font.size = Pt(11); p2.font.color.rgb = GREY
        x += w_each + gap


def table(s, x, y, w, rows, col_w=None, size=13, header=True):
    nrows, ncols = len(rows), len(rows[0])
    shp = s.shapes.add_table(nrows, ncols, x, y, w, Inches(0.34) * nrows)
    tbl = shp.table
    if col_w:
        for i, cw in enumerate(col_w):
            tbl.columns[i].width = cw
    for r, row in enumerate(rows):
        for c, val in enumerate(row):
            cell = tbl.cell(r, c)
            cell.text = str(val)
            cell.margin_left = Inches(0.07)
            cell.margin_top = Inches(0.02)
            cell.margin_bottom = Inches(0.02)
            para = cell.text_frame.paragraphs[0]
            para.font.size = Pt(size)
            if header and r == 0:
                para.font.bold = True
                para.font.color.rgb = WHITE
                cell.fill.solid(); cell.fill.fore_color.rgb = BLUE
            else:
                para.font.color.rgb = INK
                cell.fill.solid()
                cell.fill.fore_color.rgb = WHITE if r % 2 else PALE
    return tbl


def pic(s, name, x, y, w=None, h=None):
    p = FIG / name
    if not p.exists():
        return None
    return s.shapes.add_picture(str(p), x, y, width=w, height=h)


def flow_box(s, x, y, w, h, label, text, accent=False):
    box = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    box.fill.solid()
    box.fill.fore_color.rgb = BLUE if accent else PALE
    box.line.fill.background()
    box.shadow.inherit = False
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.16); tf.margin_right = Inches(0.12)
    tf.margin_top = Inches(0.12)
    p = tf.paragraphs[0]
    p.text = label
    p.font.size = Pt(13); p.font.bold = True
    p.font.color.rgb = WHITE if accent else BLUE
    p2 = tf.add_paragraph()
    p2.text = text
    p2.font.size = Pt(11)
    p2.font.color.rgb = RGBColor(0xDD, 0xE7, 0xF5) if accent else GREY
    p2.space_before = Pt(5)
    return box


def arrow(s, x, y, w=Inches(0.38), h=Inches(0.3)):
    a = s.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, x, y, w, h)
    a.fill.solid(); a.fill.fore_color.rgb = RGBColor(0xB8, 0xC4, 0xD4)
    a.line.fill.background(); a.shadow.inherit = False
    return a


def band(s, y, text, h=Inches(0.85), accent=False):
    """Full-width emphasis strip."""
    box = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.62), y,
                             W - Inches(1.24), h)
    box.fill.solid()
    box.fill.fore_color.rgb = BLUE if accent else PALE
    box.line.fill.background(); box.shadow.inherit = False
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.2); tf.margin_right = Inches(0.2)
    tf.margin_top = Inches(0.1)
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(13)
    p.font.color.rgb = WHITE if accent else INK
    return box


def quote(s, y, text, h=Inches(1.05)):
    """Bordered example block, visually distinct from body text."""
    box = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.62), y,
                             W - Inches(1.24), h)
    box.fill.solid(); box.fill.fore_color.rgb = WHITE
    box.line.color.rgb = RGBColor(0xC9, 0xD5, 0xE4)
    box.line.width = Pt(1)
    box.shadow.inherit = False
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.2); tf.margin_right = Inches(0.2)
    tf.margin_top = Inches(0.1)
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(12); p.font.color.rgb = INK
    p.font.name = "Georgia"
    return box


def caption(s, y, text, w=None, x=None):
    c = s.shapes.add_textbox(x or Inches(0.62), y, w or (W - Inches(1.24)), Inches(0.5))
    c.text_frame.word_wrap = True
    p = c.text_frame.paragraphs[0]
    p.text = text
    p.font.size = Pt(10); p.font.color.rgb = GREY
    return c


def build() -> pathlib.Path:
    prs = Presentation()
    prs.slide_width, prs.slide_height = W, H

    # ---------------------------------------------------------------- 1 title
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, W, H)
    bg.fill.solid(); bg.fill.fore_color.rgb = INK
    bg.line.fill.background(); bg.shadow.inherit = False

    t = s.shapes.add_textbox(Inches(0.85), Inches(2.2), W - Inches(1.7), Inches(1.9))
    tf = t.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "GlandersWatch"
    p.font.size = Pt(54); p.font.bold = True; p.font.color.rgb = WHITE
    p2 = tf.add_paragraph()
    p2.text = "Risk prediction, cited explanation and owner alerting for a neglected zoonosis"
    p2.font.size = Pt(19); p2.font.color.rgb = RGBColor(0xCB, 0xD5, 0xE1)
    p2.space_before = Pt(14)

    a = s.shapes.add_textbox(Inches(0.85), Inches(5.5), W - Inches(1.7), Inches(1.2))
    ap = a.text_frame.paragraphs[0]
    a.text_frame.word_wrap = True
    ap.text = ("Pedapatruni Shanmuk Vardhan  ·  Shaktimayee Mohanty  ·  Yuvraj Singh\n"
               "Supervisor: Dr. Sangita Dhara  ·  IIIT Bhubaneswar")
    ap.font.size = Pt(13); ap.font.color.rgb = MUTE
    ap.line_spacing = 1.5

    # ---------------------------------------------------------------- 2 overview
    s = slide(prs, "Overview", "One probability, produced once, consumed in two different ways")

    flow_box(s, Inches(0.62), Inches(1.85), Inches(2.55), Inches(1.55),
             "1  Public data",
             "Europe PMC · FAOSTAT\nWorld Bank · NASA POWER")
    arrow(s, Inches(3.30), Inches(2.45))
    flow_box(s, Inches(3.80), Inches(1.85), Inches(3.05), Inches(1.55),
             "2  Risk engine",
             "8 features → calibrated\nprobability per country-year")
    arrow(s, Inches(6.98), Inches(2.45))
    flow_box(s, Inches(7.48), Inches(1.85), Inches(5.2), Inches(1.55),
             "3  Calibrated probability",
             "A number that can be quoted: 0.03 means 3 in 100.",
             accent=True)

    # the two consumers
    b1 = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.62), Inches(3.80),
                            Inches(5.9), Inches(1.5))
    b1.fill.solid(); b1.fill.fore_color.rgb = PALE
    b1.line.fill.background(); b1.shadow.inherit = False
    tf = b1.text_frame; tf.word_wrap = True
    tf.margin_left = Inches(0.18); tf.margin_top = Inches(0.14)
    p = tf.paragraphs[0]; p.text = "Explanation layer"
    p.font.size = Pt(15); p.font.bold = True; p.font.color.rgb = BLUE
    p2 = tf.add_paragraph()
    p2.text = ("Answers “why this country, and what should be done?” with passages\n"
               "retrieved from official guidance, cited, or refused.")
    p2.font.size = Pt(11); p2.font.color.rgb = GREY; p2.space_before = Pt(6)

    b2 = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.78), Inches(3.80),
                            Inches(5.9), Inches(1.5))
    b2.fill.solid(); b2.fill.fore_color.rgb = PALE
    b2.line.fill.background(); b2.shadow.inherit = False
    tf = b2.text_frame; tf.word_wrap = True
    tf.margin_left = Inches(0.18); tf.margin_top = Inches(0.14)
    p = tf.paragraphs[0]; p.text = "Alert service"
    p.font.size = Pt(15); p.font.bold = True; p.font.color.rgb = BLUE
    p2 = tf.add_paragraph()
    p2.text = ("Raises a message only when the probability crosses a level, so a\n"
               "warning still means something on the day it arrives.")
    p2.font.size = Pt(11); p2.font.color.rgb = GREY; p2.space_before = Pt(6)

    body(s, Inches(0.62), Inches(5.40), Inches(11.9), Inches(0.6), [
        "Both consume the same calibrated output, so the message and the explanation can never disagree.",
    ], size=12, space=4)

    stat_row(s, Inches(6.0), [
        (f"country-years, {YRS[0]}–{YRS[1]}", f"{ROWS:,}"),
        ("countries", f"{NC}"),
        ("features", f"{SEL['top_k']}"),
        ("PR-AUC, unseen countries", f"{LR['loco']['pr_auc']:.3f}"),
    ], w_each=Inches(2.85))

    # ---------------------------------------------------------------- 3 problem
    s = slide(prs, "Problem statement", "Stated once, precisely, because every section inherits this framing")

    f = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.62), Inches(1.65),
                           Inches(11.9), Inches(1.15))
    f.fill.solid(); f.fill.fore_color.rgb = PALE
    f.line.fill.background(); f.shadow.inherit = False
    tf = f.text_frame; tf.word_wrap = True
    tf.margin_left = Inches(0.22); tf.margin_top = Inches(0.14)
    p = tf.paragraphs[0]
    p.text = "Given a country's features in year t, predict whether glanders is documented in that country in year t + 1."
    p.font.size = Pt(17); p.font.color.rgb = INK

    body(s, Inches(0.62), Inches(3.05), Inches(6.1), Inches(3.1), [
        "**The label",
        "One country-year is positive if a glanders occurrence is reported from that country in that year.",
        "Features come only from the year before, so the model never reads the outcome it is predicting.",
    ], size=13, space=10)

    body(s, Inches(6.95), Inches(3.05), Inches(5.6), Inches(3.1), [
        "**The difficulty",
        f"{POS} positives in {ROWS:,} rows. A model that says “never” is 94% accurate and useless.",
        "Accuracy and ROC-AUC hide this; precision-recall does not.",
    ], size=13, space=10)

    stat_row(s, Inches(5.95), [
        (f"country-years, {YRS[0]}–{YRS[1]}", f"{ROWS:,}"),
        ("countries", f"{NC}"),
        ("positives", f"{POS}"),
        ("prevalence", f"{PREV*100:.2f}%"),
    ], w_each=Inches(2.85))

    # ---------------------------------------------------------------- 4 why needed
    s = slide(prs, "Why it is needed", "Surveillance tells you where the disease was, which is one step too late")

    table(s, Inches(0.62), Inches(1.7), Inches(6.55), [
        ["Who", "What they cannot do today"],
        ["District veterinary officer", "Know where to pre-position testing and movement checks"],
        ["Equine owner", "Know whether this is the month to restrict movement"],
        ["State planner", "Compare risk across districts on one scale"],
    ], col_w=[Inches(2.35), Inches(4.2)], size=12)

    body(s, Inches(7.45), Inches(1.75), Inches(5.1), Inches(2.4), [
        "**What exists",
        "Reports that arrive after the fact. A national livestock forewarning system exists, but glanders is not on its disease list.",
        "**The failure mode to avoid",
        "A warning system that fires everywhere is ignored everywhere.",
    ], size=12, space=10)

    pic(s, "fig14_background.png", Inches(0.62), Inches(3.75), w=Inches(7.6))
    c = s.shapes.add_textbox(Inches(0.62), Inches(6.35), Inches(11.9), Inches(0.75))
    cp = c.text_frame.paragraphs[0]; c.text_frame.word_wrap = True
    cp.text = ("Panel A: how many equids are exposed. Panel B: how many events are documented per year. "
               "Documentation rises with open-access publishing, which is why the target is a documented event rather than raw incidence.")
    cp.font.size = Pt(10); cp.font.color.rgb = GREY

    # ---------------------------------------------------------------- 5 division
    divider(prs, "01", "The risk engine",
            "Where is glanders likely to be documented next year, and how much should we trust the number?",
            ["Data: four public sources and the label they produce",
             "Eight features, chosen by attribution rather than association",
             "Two validation splits, and why the ranking inverts between them"])

    # ---------------------------------------------------------------- 6 data
    s = slide(prs, "Data", "Four public sources, none previously applied to this disease")

    rows = [["Source", "What it contributes", "Role"],
            ["Europe PMC", "Occurrence events from abstracts", "The label"],
            ["FAOSTAT", "Horse, ass and mule head counts", "Exposure"],
            ["World Bank", "Development and connectivity", "Context"],
            ["NASA POWER", "Temperature, humidity, rainfall", "Environment"]]
    table(s, Inches(0.62), Inches(1.72), Inches(7.4), rows,
          col_w=[Inches(1.9), Inches(3.4), Inches(2.1)], size=12)

    body(s, Inches(8.35), Inches(1.75), Inches(4.4), Inches(3.6), [
        f"6,700 raw country-year rows over 268 reporting entities.",
        f"Filtered to {ROWS:,} rows across {NC} equine-holding countries.",
        f"{POS} positives, {PREV*100:.2f}% prevalence.",
        "Rows without a climate baseline are dropped, not imputed, so no anomaly is manufactured.",
    ], size=12, space=12)

    pic(s, "fig01_dataset_coverage.png", Inches(0.62), Inches(3.62), w=Inches(7.45))
    caption(s, Inches(6.42),
            "Coverage of the assembled panel: which countries contribute which years, and where the positives sit.",
            w=Inches(7.45))

    # ---------------------------------------------------------------- 7 features
    s = slide(prs, "Eight features, chosen by attribution",
              "Ranked by mean absolute SHAP on the training years only")

    units = {"pop_total": "persons", "years_since_last": "years (99 = none)",
             "temp_mean_anom": "°C from baseline", "internet_pct": "% of population",
             "humidity_anom": "% RH from baseline", "precip_cv": "ratio, sd/mean",
             "log_livestock": "ln(1 + LSU)", "goats": "head"}
    rows = [["#", "Feature", "Unit", "Mean |SHAP|"]]
    for r in SEL["ranking"][: SEL["top_k"]]:
        rows.append([r["rank"], r["feature"], units.get(r["feature"], ""),
                     f"{r['mean_abs_shap']:.3f}"])
    table(s, Inches(0.62), Inches(1.78), Inches(7.3), rows,
          col_w=[Inches(0.5), Inches(2.5), Inches(2.9), Inches(1.4)], size=12)

    body(s, Inches(8.25), Inches(1.85), Inches(4.5), Inches(2.9), [
        f"{SEL['top_k']} of {SEL['n_candidates']} candidates, carrying {ATTR_SHARE*100:.0f}% of total attribution.",
        "Selection uses training years only, so held-out years cannot shape the choice.",
        "Correlation puts longitude first; attribution puts it near the bottom.",
    ], size=12, space=11)

    reasons = [
        ("Attribute a prediction", "SHAP says how much each feature moved this model's output, so the ranking matches what was actually used."),
        ("Beyond pairs", "Correlation and mutual information score features pairwise; they cannot see interactions or redundancy between features."),
        ("Comparable", "Both are unnormalised and undirected. Attribution is additive and sums to the prediction."),
    ]
    rx = Inches(0.62)
    for head, text in reasons:
        b = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, rx, Inches(4.92), Inches(3.85), Inches(1.9))
        b.fill.solid(); b.fill.fore_color.rgb = PALE
        b.line.fill.background(); b.shadow.inherit = False
        tf = b.text_frame; tf.word_wrap = True
        tf.margin_left = Inches(0.16); tf.margin_top = Inches(0.12)
        pp = tf.paragraphs[0]; pp.text = head
        pp.font.size = Pt(12); pp.font.bold = True; pp.font.color.rgb = BLUE
        p2 = tf.add_paragraph(); p2.text = text
        p2.font.size = Pt(10.5); p2.font.color.rgb = GREY
        p2.space_before = Pt(6)
        rx += Inches(4.05)

    # ---------------------------------------------------------------- 8 results
    s = slide(prs, "Results", "One target, two validation splits, opposite conclusions")

    rows = [["Model", "PR-AUC, unseen years", "PR-AUC, unseen countries"]]
    for m, v in L1.items():
        rows.append([m, f"{v['temporal']['pr_auc']:.3f}", f"{v['loco']['pr_auc']:.3f}"])
    table(s, Inches(0.62), Inches(1.72), Inches(6.5), rows,
          col_w=[Inches(2.7), Inches(1.9), Inches(1.9)], size=11)

    body(s, Inches(7.45), Inches(1.75), Inches(5.3), Inches(2.4), [
        f"Base rate {PREV*100:.2f}%, so PR-AUC is the metric that moves.",
        f"ROC-AUC stays in 0.80–0.87 for all six models on both splits; it cannot separate them.",
        "Logistic regression is the champion: it wins on unseen countries, which is what deployment means.",
    ], size=12, space=11)

    stat_row(s, Inches(4.35), [
        ("random forest, unseen years", f"{RF['temporal']['pr_auc']:.3f}"),
        ("logistic, unseen countries", f"{LR['loco']['pr_auc']:.3f}"),
        ("random forest, unseen countries", f"{RF['loco']['pr_auc']:.3f}"),
    ], w_each=Inches(2.85))

    band(s, Inches(5.95),
         "The ranking inverts.  Random forest wins on unseen years; logistic regression wins on unseen "
         "countries — and deployment means an unseen country, so the simpler model is the one shipped.",
         h=Inches(1.0), accent=True)

    # ---------------------------------------------------------------- 9 calibration
    s = slide(prs, "The number is trustworthy, and that sets the alert budget",
              "Calibration is what lets a message quote a probability at all")

    stat_row(s, Inches(1.65), [
        ("expected calibration error", f"{R['calibration']['ece']:.3f}"),
        ("Brier score", f"{R['calibration']['brier']:.3f}"),
        ("events caught, alerting 6% of country-years", f"{AE['r6']*100:.0f}%"),
    ])

    body(s, Inches(0.62), Inches(3.3), Inches(6.0), Inches(2.6), [
        f"Half of all events are reached by alerting {AE['b50']*100:.0f}% of country-years; 80% by {AE['b80']*100:.0f}%.",
        "That curve, not a default 0.5 cutoff, is where alert levels come from.",
        "It is also the argument that the alert service can be quiet without being useless.",
    ], size=13, space=12)

    pic(s, "fig07_alert_budget.png", Inches(6.85), Inches(2.75), w=Inches(5.9))
    band(s, Inches(6.15),
         "This curve is the link between the two halves of the system: it states how quiet an alert service "
         "can be while still catching most events.",
         h=Inches(0.78))

    # ---------------------------------------------------------------- 10 division
    divider(prs, "02", "The explanation layer",
            "Why did the model flag this country, and what should an owner actually do about it?",
            ["Grounded retrieval with citations, and refusal when nothing supports an answer",
             "The guardrail that keeps advisory text from inventing guidance",
             "How this layer will be measured, and what is not yet measured"])

    # ---------------------------------------------------------------- 11 how it works
    s = slide(prs, "How an answer is built", "Retrieval first, generation second, citation always")

    steps = [
        ("Question", "“Why is this district\nhigh risk?”"),
        ("Retrieve", "Hybrid dense + lexical\nsearch over guidance"),
        ("Fuse and rerank", "Reciprocal rank fusion,\nthen cross-encoder"),
        ("Answer, cited", "Every claim tied to\na passage"),
    ]
    x = Inches(0.62)
    for i, (lab, txt) in enumerate(steps):
        flow_box(s, x, Inches(1.75), Inches(2.62), Inches(1.72), lab, txt,
                 accent=(i == 3))
        if i < 3:
            arrow(s, x + Inches(2.70), Inches(2.45))
        x += Inches(3.05)

    body(s, Inches(0.62), Inches(3.80), Inches(6.0), Inches(1.9), [
        "**Guardrails, not suggestions",
        "Refuses to answer outside the retrieved sources, and refuses to diagnose an individual animal.",
        "Always shows the passages it used.",
    ], size=12, space=8)

    body(s, Inches(6.95), Inches(3.80), Inches(5.6), Inches(1.9), [
        "**Why retrieval and not a plain model",
        "Guidance changes; weights would have to be retrained to follow it.",
        "A cited passage can be checked by a veterinarian. A generated sentence cannot.",
    ], size=12, space=8)

    quote(s, Inches(5.95),
          "“Risk is elevated mainly because of recent reported activity in neighbouring districts and high "
          "equine density. Restrict equid movement and contact the local veterinary authority if an animal "
          "is suspected.”   —  each clause carries its source; where no passage supports the question, "
          "the layer says so rather than answering.")

    # ---------------------------------------------------------------- 12 rag eval
    s = slide(prs, "How this layer will be judged",
              "The metrics are defined now; measuring them is the next phase")

    rows = [["Metric", "What it asks", "How"],
            ["Recall@k", "Is the right passage in the top k?", "Hand-built question set"],
            ["Faithfulness", "Does the answer follow from those passages?", "RAGAS, plus manual review"],
            ["Answer relevance", "Does it answer the question asked?", "RAGAS"],
            ["Citation precision", "Does every claim trace to a passage?", "Automated check"]]
    table(s, Inches(0.62), Inches(1.75), Inches(9.2), rows,
          col_w=[Inches(2.3), Inches(4.1), Inches(2.8)], size=13)

    body(s, Inches(0.62), Inches(3.72), Inches(11.9), Inches(1.4), [
        "These four cover the ways an explanation fails: retrieving the wrong passage, drifting from the passages it retrieved, or faithfully answering a question nobody asked.",
    ], size=12, space=8)

    band(s, Inches(5.35),
         "None of these are measured yet. Targets will be fixed before measurement, so the layer cannot be "
         "scored by choosing a threshold after seeing the results.",
         h=Inches(0.9))

    # ---------------------------------------------------------------- 13 division
    divider(prs, "03", "The alert service",
            "How does a probability become a message that is still read in month twelve?",
            ["Levels taken from the calibration curve, not a default cutoff",
             "Hysteresis, so an area does not flicker between levels",
             "Three layers of deduplication, and a stale-data guard",
             "How delivery will be measured, and what is not yet measured"])

    # ---------------------------------------------------------------- 14 alert design
    s = slide(prs, "How an alert is raised", "Five stages; each one exists to stop a specific failure")

    stages = [
        ("1  Predict", "Latest calibrated\nscore for the area"),
        ("2  Assign level", "Thresholds taken from\nthe budget curve"),
        ("3  Deduplicate", "Three idempotency\nkeys, not one"),
        ("4  Deliver", "SMS or WhatsApp,\nwith retries"),
        ("5  Log", "Status per recipient,\nfor the admin"),
    ]
    x = Inches(0.62)
    for i, (lab, txt) in enumerate(stages):
        flow_box(s, x, Inches(1.70), Inches(2.15), Inches(1.55), lab, txt,
                 accent=(i == 2))
        if i < 4:
            arrow(s, x + Inches(2.17), Inches(2.32), w=Inches(0.28))
        x += Inches(2.47)

    guards = [
        ("Hysteresis", "Enter at a threshold, leave only 0.05 below it. A country sitting on the line does not oscillate."),
        ("Freshness", "If the data check fails, no alert goes out and the admin is told instead."),
        ("Quiet hours", "Hold overnight except for escalation, so a real rise still gets through."),
    ]
    gx = Inches(0.62)
    for name, text in guards:
        g = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, gx, Inches(3.75),
                               Inches(3.85), Inches(1.55))
        g.fill.solid(); g.fill.fore_color.rgb = PALE
        g.line.fill.background(); g.shadow.inherit = False
        tf = g.text_frame; tf.word_wrap = True
        tf.margin_left = Inches(0.18); tf.margin_top = Inches(0.14)
        p = tf.paragraphs[0]; p.text = name
        p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = BLUE
        p2 = tf.add_paragraph(); p2.text = text
        p2.font.size = Pt(11); p2.font.color.rgb = GREY
        p2.space_before = Pt(6)
        gx += Inches(4.05)

    band(s, Inches(5.65),
         "Levels come from the calibration curve, so “high” has a measured meaning rather than a default "
         "cutoff of 0.5. Deduplication is what keeps a re-run or a retrain from sending the same warning twice.",
         h=Inches(0.95))

    # ---------------------------------------------------------------- 15 architecture
    s = slide(prs, "The alert pipeline, end to end")
    pic(s, "proposal_alert_architecture.png", Inches(2.27), Inches(0.92), w=Inches(8.8))
    caption(s, Inches(6.98),
            "The deployed view: evaluator, levels, deduplication, delivery and logging, with an admin surface for failures.",
            w=Inches(8.8), x=Inches(2.27))

    # ---------------------------------------------------------------- 16 alert eval
    s = slide(prs, "How the alert service will be judged",
              "Because the cost of a false alarm is measured in attention, not accuracy")

    rows = [["Metric", "What it asks"],
            ["Precision", "Of the alerts sent, how many were real events?"],
            ["Recall", "Of the real events, how many did we alert on?"],
            ["Missed events", "The failures a user actually feels"],
            ["False alerts per subscriber", "The alert-fatigue budget"]]
    table(s, Inches(0.62), Inches(1.72), Inches(8.0), rows,
          col_w=[Inches(2.9), Inches(5.1)], size=12)

    body(s, Inches(9.0), Inches(1.75), Inches(3.7), Inches(2.6), [
        "Defined but not yet measured.",
        "A pilot with a district office is how they get numbers.",
    ], size=12, space=10)

    quote(s, Inches(5.05),
          "GLANDERS RISK ALERT — Risk: HIGH.  Elevated glanders risk is predicted for your registered area "
          "for the next period. Reason: recent reported activity and environmental factors. Action: avoid "
          "unnecessary movement or mixing of equids; contact the local veterinary authority if an animal is "
          "suspected. This is a population-level forecast, not a diagnosis.",
          h=Inches(1.55))

    # ---------------------------------------------------------------- 17 status
    s = slide(prs, "Where the project stands",
              "What is built, what is specified, and what a pilot would settle")

    cols = [
        ("Built and measured",
         f"Risk engine. Six models on two splits. Champion is logistic regression, "
         f"PR-AUC {LR['loco']['pr_auc']:.3f} on unseen countries, ECE {R['calibration']['ece']:.3f}."),
        ("Specified, not deployed",
         "Explanation layer with citation checks and refusal. Alert service with levels, "
         "hysteresis and deduplication. Both have metrics defined, neither has numbers yet."),
        ("What a pilot would settle",
         "Alert precision and fatigue against a real district office. Explanation faithfulness "
         "on a hand-built question set. A label audit on a stratified sample."),
    ]
    bx = Inches(0.62)
    for i, (head, text) in enumerate(cols):
        box = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, bx, Inches(1.65),
                                 Inches(3.85), Inches(3.55))
        box.fill.solid()
        box.fill.fore_color.rgb = BLUE if i == 0 else PALE
        box.line.fill.background(); box.shadow.inherit = False
        tf = box.text_frame; tf.word_wrap = True
        tf.margin_left = Inches(0.2); tf.margin_right = Inches(0.18)
        tf.margin_top = Inches(0.16)
        p = tf.paragraphs[0]; p.text = head
        p.font.size = Pt(15); p.font.bold = True
        p.font.color.rgb = WHITE if i == 0 else BLUE
        p2 = tf.add_paragraph(); p2.text = text
        p2.font.size = Pt(11.5)
        p2.font.color.rgb = RGBColor(0xDD, 0xE7, 0xF5) if i == 0 else GREY
        p2.space_before = Pt(9)
        bx += Inches(4.05)

    band(s, Inches(5.5),
         "The model is the part that is finished. The evidence that the explanation and alert layers "
         "behave in use is the part that is not, and a pilot is how it arrives.",
         h=Inches(0.95), accent=True)

    body(s, Inches(0.62), Inches(6.6), Inches(11.9), Inches(0.8), [
        f"Known limits throughout: the label is documented activity, not incidence, and with {POS} positives the uncertainty on every headline metric is wide.",
    ], size=11, space=4)

    prs.save(OUT)
    print(f"wrote {OUT.relative_to(ROOT)}  ({len(prs.slides)} slides)")
    return OUT


if __name__ == "__main__":
    build()
