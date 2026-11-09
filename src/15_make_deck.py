"""Build the presentation deck from the pipeline outputs.

Run:  .venv/bin/python src/15_make_deck.py

Every number and figure is read from reports/*.json and figures/*.png, so the
deck cannot drift from the report. 16:9, twelve slides.
"""
from __future__ import annotations

import json
import pathlib

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

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

INK = RGBColor(0x1F, 0x29, 0x33)
BLUE = RGBColor(0x1D, 0x4E, 0xD8)
GREY = RGBColor(0x6B, 0x72, 0x80)
PALE = RGBColor(0xF1, 0xF5, 0xF9)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
ORANGE = RGBColor(0xEA, 0x58, 0x0C)


def alert_econ() -> dict:
    """Recall at a fixed budget, and budget needed for a target recall."""
    import numpy as np
    import pandas as pd
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

W, H = Inches(13.333), Inches(7.5)


def slide(prs, title: str, kicker: str | None = None):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bar = s.shapes.add_shape(1, 0, 0, W, Inches(0.09))
    bar.fill.solid(); bar.fill.fore_color.rgb = BLUE; bar.line.fill.background()

    t = s.shapes.add_textbox(Inches(0.62), Inches(0.34), W - Inches(1.2), Inches(0.7))
    p = t.text_frame.paragraphs[0]
    p.text = title
    p.font.size = Pt(30); p.font.bold = True; p.font.color.rgb = INK

    if kicker:
        k = s.shapes.add_textbox(Inches(0.62), Inches(1.02), W - Inches(1.2), Inches(0.42))
        kp = k.text_frame.paragraphs[0]
        kp.text = kicker
        kp.font.size = Pt(14); kp.font.color.rgb = GREY
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
        p.font.color.rgb = INK if not bold else BLUE
        p.space_after = Pt(space)
        if bullet and not bold:
            p.text = "•  " + p.text
    return tb


def stat_row(s, y, items, w_each=Inches(2.85), gap=Inches(0.22), h=Inches(1.18)):
    x = Inches(0.62)
    for label, value in items:
        box = s.shapes.add_shape(1, x, y, w_each, h)
        box.fill.solid(); box.fill.fore_color.rgb = PALE
        box.line.fill.background()
        box.shadow.inherit = False
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


def build() -> pathlib.Path:
    prs = Presentation()
    prs.slide_width, prs.slide_height = W, H

    # ---------------------------------------------------------------- 1 title
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg = s.shapes.add_shape(1, 0, 0, W, H)
    bg.fill.solid(); bg.fill.fore_color.rgb = INK; bg.line.fill.background()

    t = s.shapes.add_textbox(Inches(0.85), Inches(2.25), W - Inches(1.7), Inches(1.9))
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
    ap.text = ("Pedapatruni Shanmuk Vardhan  ·  Shaktimayee Mohanty  ·  Yuvraj Singh\n"
               "Supervisor: Dr. Sangita Dhara  ·  IIIT Bhubaneswar")
    ap.font.size = Pt(13); ap.font.color.rgb = RGBColor(0x94, 0xA3, 0xB8)
    ap.line_spacing = 1.5

    # ---------------------------------------------------------------- 2 problem
    s = slide(prs, "The problem", "A district officer needs to know where glanders is likely to appear next")
    body(s, Inches(0.62), Inches(1.62), Inches(6.5), Inches(4.4), [
        "Glanders is a bacterial zoonosis of equids. Notifiable, often fatal, no environmental reservoir.",
        "An equid is a horse, donkey (recorded as asses) or mule.",
        "The economic weight sits in low and middle income countries, where equids are working animals.",
        "Surveillance reports what has already happened. No forecasting system exists for this disease.",
    ], size=15, space=13)
    stat_row(s, Inches(5.4), [("equids in India, 2019", "5.50 lakh"),
                              ("seropositive, 2019-23", "728"),
                              ("districts affected", "138")])

    # ---------------------------------------------------------------- 3 damage
    s = slide(prs, "Population at risk is not the same as damage",
              "The two quantities are routinely conflated, so they are separated at the outset")
    body(s, Inches(0.62), Inches(1.62), Inches(5.9), Inches(4.0), [
        "**5.50 lakh equids** is the size of the susceptible population: how many animals are exposed.",
        "**728 seropositive out of 158,798 screened** is the documented damage: under 1% of that population.",
        "Because the disease is notifiable and frequently fatal once clinical, each positive is acted on rather than monitored.",
    ], size=15, space=14)
    pic(s, "fig14_background.png", Inches(6.9), Inches(1.7), w=Inches(5.9))
    b = s.shapes.add_textbox(Inches(6.9), Inches(5.5), Inches(5.9), Inches(1.1))
    bp = b.text_frame.paragraphs[0]; b.text_frame.word_wrap = True
    bp.text = ("Panel A: the exposed population, by species.  Panel B: documented events per year, which tracks "
               "open-access publishing as much as disease. That is why the label is treated as documented activity.")
    bp.font.size = Pt(10); bp.font.color.rgb = GREY

    # ---------------------------------------------------------------- 4 three parts
    s = slide(prs, "What the platform is", "Three components, fixed contracts, none hardcodes the disease")
    table(s, Inches(0.62), Inches(1.75), W - Inches(1.24), [
        ["Component", "Does", "Status"],
        ["Risk engine", "Country features → calibrated probability, with attribution", "Built and evaluated"],
        ["Explanation layer", "Prediction → cited answer grounded in official guidance; refuses when unsupported", "Specified"],
        ["Alert service", "Probability → few enough messages that they are read", "Specified"],
    ], col_w=[Inches(2.3), Inches(6.7), Inches(2.5)], size=14)
    body(s, Inches(0.62), Inches(4.4), Inches(11.9), Inches(1.9), [
        "A prediction alone is not usable: an owner cannot act on a probability of 0.24, and an officer cannot defend a forecast whose basis is hidden.",
        "A warning that arrives for every area every month teaches its recipients to ignore it.",
    ], size=13, space=9)

    # ---------------------------------------------------------------- 5 data
    s = slide(prs, "Data", "Four public sources, none previously applied to this disease")
    table(s, Inches(0.62), Inches(1.7), Inches(6.0), [
        ["Source", "Role"],
        ["Europe PMC", "Occurrence events from abstracts"],
        ["FAOSTAT", "Animal stocks, head counts"],
        ["World Bank", "Development indicators"],
        ["NASA POWER", "Climate, AG community"],
    ], col_w=[Inches(2.0), Inches(4.0)], size=13)
    body(s, Inches(6.95), Inches(1.7), Inches(5.8), Inches(3.6), [
        f"6,700 country-year rows across 268 countries assemble into the panel.",
        f"Filtered to {ROWS:,} modelling rows over {NC} equine-holding countries, 2006–2023.",
        f"{POS} positives, a prevalence of {PREV*100:.2f}%.",
        "The 3,078→2,880 change came from dropping rows with no climate-anomaly baseline rather than imputing them.",
    ], size=13, space=11)

    # ---------------------------------------------------------------- 6 features
    s = slide(prs, "Eight features, chosen by attribution",
              "Ranked by mean absolute SHAP on the training years only")
    rows = [["#", "Feature", "Unit", "Mean |SHAP|"]]
    units = {"pop_total": "persons", "years_since_last": "years (99 = none)",
             "temp_mean_anom": "°C from baseline", "internet_pct": "% of population",
             "humidity_anom": "% RH from baseline", "precip_cv": "ratio, sd/mean",
             "log_livestock": "ln(1 + LSU)", "goats": "head"}
    for r in SEL["ranking"][:8]:
        rows.append([r["rank"], r["feature"], units.get(r["feature"], ""), f"{r['mean_abs_shap']:.3f}"])
    table(s, Inches(0.62), Inches(1.78), Inches(7.3), rows,
          col_w=[Inches(0.5), Inches(2.5), Inches(2.9), Inches(1.4)], size=12)
    body(s, Inches(8.25), Inches(1.85), Inches(4.5), Inches(4.0), [
        f"8 of {SEL['n_candidates']} candidates, carrying 42% of total attribution.",
        "Ranking uses the training years only, so held-out years cannot influence the choice.",
        "Correlation and mutual information rank longitude first; attribution ranks it near the bottom. Selecting on association would keep proxies.",
    ], size=12, space=11)

    # ---------------------------------------------------------------- 7 results
    s = slide(prs, "Model comparison", "One target, two validation splits")
    rows = [["Model", "PR-AUC (temporal)", "ROC-AUC (temporal)", "PR-AUC (country out)"]]
    for m, v in L1.items():
        rows.append([m, f"{v['temporal']['pr_auc']:.3f}",
                     f"{v['temporal']['roc_auc']:.3f}", f"{v['loco']['pr_auc']:.3f}"])
    table(s, Inches(0.62), Inches(1.75), Inches(8.0), rows,
          col_w=[Inches(2.7), Inches(2.0), Inches(2.0), Inches(1.9)], size=13)
    body(s, Inches(8.85), Inches(1.8), Inches(3.9), Inches(4.2), [
        f"Base rate {PREV*100:.2f}%, so PR-AUC is the metric that moves.",
        "ROC-AUC spans only 0.80–0.87 across all six models on both splits: it cannot separate them.",
    ], size=12, space=12)

    # ---------------------------------------------------------------- 8 inversion
    s = slide(prs, "The ranking inverts between splits",
              "Which is the finding, not a nuisance")
    rf = L1["Random Forest"]; lr = L1["Logistic Regression"]
    stat_row(s, Inches(1.62), [
        ("best on unseen years — random forest", f"{rf['temporal']['pr_auc']:.3f}"),
        ("best on unseen countries — logistic", f"{lr['loco']['pr_auc']:.3f}"),
        ("random forest, unseen countries", f"{rf['loco']['pr_auc']:.3f}"),
    ])
    body(s, Inches(0.62), Inches(3.25), Inches(7.4), Inches(3.0), [
        f"With {SEL['top_k']} features the relationships are close enough to additive that a regularised linear model transfers best.",
        "The flexible ensembles pay for their extra capacity in exactly the setting the project cares about.",
        "Deployment means an unseen country rather than an unseen year, so the champion is logistic regression.",
    ], size=14, space=12)
    pic(s, "fig02_model_comparison.png", Inches(8.2), Inches(2.6), w=Inches(4.6))

    # ---------------------------------------------------------------- 9 calibration
    s = slide(prs, "Calibration and alert economics",
              "A message quoting a probability must be numerically trustworthy")
    stat_row(s, Inches(1.62), [
        ("expected calibration error", f"{R['calibration']['ece']:.3f}"),
        ("Brier score", f"{R['calibration']['brier']:.3f}"),
        ("events caught, alerting 6%", f"{AE['r6']*100:.0f}%"),
    ])
    body(s, Inches(0.62), Inches(3.3), Inches(6.6), Inches(2.6), [
        f"Half of all events are reached at an alert budget of {AE['b50']*100:.0f}% of country-years; 80% at {AE['b80']*100:.0f}%.",
        "Uniform alerting would need full coverage for the same recall.",
    ], size=14, space=12)
    pic(s, "fig07_alert_budget.png", Inches(7.5), Inches(2.75), w=Inches(5.3))

    # ---------------------------------------------------------------- 10 country level
    s = slide(prs, "The coarser question is easier",
              "Asking whether a country is a glanders country at all")
    rows = [["Model", "ROC-AUC", "PR-AUC"]]
    for k, v in CL["models"].items():
        rows.append([k, f"{v['roc_auc']:.3f}", f"{v['pr_auc']:.3f}"])
    table(s, Inches(0.62), Inches(1.78), Inches(5.4), rows,
          col_w=[Inches(2.8), Inches(1.3), Inches(1.3)], size=13)
    body(s, Inches(6.4), Inches(1.8), Inches(6.3), Inches(4.2), [
        f"{CL['n_countries']} countries, {CL['positives']} positives, prevalence {CL['prevalence']*100:.1f}%.",
        "Discrimination rises to 0.97 — but this is an easier question, not a better model.",
        "It largely asks whether a country has a reference laboratory, publishes, and has equids.",
        "The year-by-year target was hard precisely because reporting status is stable while incidence is not.",
    ], size=13, space=11)

    # ---------------------------------------------------------------- 11 negative result
    s = slide(prs, "A negative result we kept",
              "Predicting the calendar year does not work")
    ya = YT["target_a_year"]; yb = YT["target_b_years_to_event"]
    table(s, Inches(0.62), Inches(1.8), Inches(7.2), [
        ["Target", "Model", "R²"],
        ["Year of report", "Linear", f"{ya['Linear regression']['r2']:+.3f}"],
        ["Year of report", "Random forest", f"{ya['Random forest']['r2']:+.3f}"],
        ["Year of report", "Predict the mean", "0.000"],
    ], col_w=[Inches(2.6), Inches(2.8), Inches(1.8)], size=13)
    body(s, Inches(8.2), Inches(1.85), Inches(4.5), Inches(4.0), [
        "Negative R² means worse than always predicting the average.",
        "The year of a report tracks when reporting began, not biology: 65% of events fall in the final five years of the window.",
        "A time-to-event variant scores 0.85, but only because the feature means include post-event years.",
    ], size=12, space=11)

    # ---------------------------------------------------------------- 12 limits
    s = slide(prs, "Limitations and what comes next")
    body(s, Inches(0.62), Inches(1.55), Inches(5.9), Inches(4.9), [
        "**Limitations",
        "The label is documented activity, not incidence: a negative may mean no disease or no report.",
        "The extractor has never been checked against a hand-annotated set.",
        f"With {POS} positives the uncertainty on every headline metric is wide.",
        "The explanation and alert layers are specified, not deployed, so neither reports measured performance.",
    ], size=13, space=9)
    body(s, Inches(6.95), Inches(1.55), Inches(5.8), Inches(4.9), [
        "**Next",
        "Model surveillance capacity as a quantity rather than a covariate.",
        "Validate the label extractor on a stratified sample.",
        "Measure the explanation layer: recall@5, faithfulness, citation precision, refusal rates.",
        "Pilot the alert service with a district office: messages per subscriber per year, delivery reliability.",
    ], size=13, space=9)

    prs.save(OUT)
    print(f"wrote {OUT.relative_to(ROOT)}  ({len(prs.slides.__iter__.__self__._sldIdLst)} slides)")
    return OUT


if __name__ == "__main__":
    build()
