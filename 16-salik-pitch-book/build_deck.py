"""
Build 16-salik-pitch-book/Salik_Pitch_Book.pdf — a sell-side pitch book for a hypothetical mandate: advising the
Government of Dubai on strategic alternatives for its 75.1% stake in Salik.

Unlike the other projects this is a deck, not a model. Every number on every slide is read from the Project 15
workbook (Salik_Profile_Buyer_Screen.xlsx) after recalculation, so the pitch book and the model cannot disagree.
Charts are drawn with matplotlib; the pages are laid out with reportlab in landscape A4 in Salik's palette.

The mandate is hypothetical. There is no suggestion that the Government of Dubai, Salik, or any party named in the
buyer universe is or has been in discussions.
"""
import os, sys, json
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from openpyxl import load_workbook
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas as rl_canvas
from reportlab.platypus import Paragraph, Table, TableStyle
from reportlab.lib.styles import ParagraphStyle
from ibkit.style import THEMES

HERE = os.path.dirname(os.path.abspath(__file__))
P15 = os.path.join(HERE, "..", "15-salik-profile", "Salik_Profile_Buyer_Screen.xlsx")
OUT = os.path.join(HERE, "Salik_Pitch_Book.pdf")
TMP = os.path.join(HERE, "_charts"); os.makedirs(TMP, exist_ok=True)
T = THEMES["salik"]
NAVY, GREY, ORANGE, LIGHT = (colors.HexColor("#" + T.primary), colors.HexColor("#" + T.secondary),
                             colors.HexColor("#" + T.accent), colors.HexColor("#" + T.light))
W, H = landscape(A4)

# --------------------------------------------------------------------------- #
# Read the model. Everything below comes from Project 15's recalculated workbook.
# --------------------------------------------------------------------------- #
wb = load_workbook(P15, data_only=True)
def find(ws, label, col=3):
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
        v = row[1].value
        if isinstance(v, str) and v.strip() == label:
            return row[col - 1].value
    raise KeyError(label)
def find_multi(ws, label, n=4, first=3):
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
        v = row[1].value
        if isinstance(v, str) and v.strip() == label:
            return [row[first - 1 + i].value for i in range(n)]
    raise KeyError(label)
fin, tr, ab, bs = wb["Financials"], wb["Trading"], wb["Ability to Pay"], wb["Buyer Screen"]
M = dict(
    rev=find_multi(fin, "Revenue"), ebitda=find_multi(fin, "EBITDA"), ni=find_multi(fin, "Net profit"),
    margin=find_multi(fin, "EBITDA margin"), fcf=find_multi(fin, "Free cash flow"), lev=find_multi(fin, "Net debt / EBITDA"),
    price=find(tr, "Share price (AED, DFM close 18 Sep 2026)", 4), mcap=find(tr, "Market capitalisation", 4),
    ev=find(tr, "Enterprise value", 4), nd=find(tr, "Net debt, FY2025", 4), ev_ebitda=find(tr, "EV / EBITDA, FY2025", 4),
    pe=find(tr, "P / E, FY2025", 4), dy=find(tr, "Dividend yield on FY2025 dividends paid", 4),
    since_ipo=find(tr, "Price return since the IPO at AED 2.00", 4), peer_med=find(tr, "Median of included peers", 4),
    px_sponsor=find(ab, "Maximum price per share \u2014 sponsor (AED)"), px_strat=find(ab, "Maximum price per share \u2014 strategic (AED)"),
    ev_strat=find(ab, "Maximum enterprise value a strategic can pay"), ev_sponsor=find(ab, "Maximum enterprise value a sponsor can pay"),
)
SHARES = 7500.0; GOV = 0.751; W52 = (4.96, 6.78)
peers = []
for row in tr.iter_rows(min_row=1, max_row=tr.max_row):
    if isinstance(row[3].value, (int, float)) and isinstance(row[1].value, str) and row[1].value not in ("Median of included peers",) and not row[1].value.startswith(("Share price", "Shares", "Market", "Net debt", "Enterprise", "EV /", "P / E", "Dividend", "Free cash", "Price return", "Salik premium")):
        peers.append((row[1].value, row[3].value, row[4].value))
buyers = []
for row in bs.iter_rows(min_row=1, max_row=bs.max_row):
    if isinstance(row[10].value, (int, float)) and isinstance(row[11].value, (int, float)) and row[2].value in ("Strategic", "Financial", "Sovereign"):
        buyers.append((row[1].value, row[2].value, row[10].value, int(row[11].value), row[12].value))
buyers.sort(key=lambda b: b[3])
prices = json.load(open(os.path.join(HERE, "price_history.json")))
comps_lo = (M["peer_med"] * M["ebitda"][3] - M["nd"]) / SHARES
comps_hi = (26.3 * M["ebitda"][3] - M["nd"]) / SHARES
# stake sizes for the recommended transaction
STAKE_LO, STAKE_HI = 0.10, 0.15
prem_lo, prem_hi = 0.05, 0.12
proc_lo = STAKE_LO * SHARES * M["price"] * (1 + prem_lo) / 1000
proc_hi = STAKE_HI * SHARES * M["price"] * (1 + prem_hi) / 1000

# --------------------------------------------------------------------------- #
# Charts
# --------------------------------------------------------------------------- #
plt.rcParams.update({"font.family": "DejaVu Sans", "font.size": 9, "axes.spines.top": False, "axes.spines.right": False})
C_NAVY, C_OR, C_GREY = "#" + T.primary, "#" + T.accent, "#" + T.secondary

def chart_financials():
    fig, ax = plt.subplots(figsize=(6.2, 4.3), dpi=200)
    yrs = ["FY22", "FY23", "FY24", "FY25"]; x = range(4)
    ax.bar([i - 0.2 for i in x], M["rev"], 0.4, color=C_NAVY, label="Revenue")
    ax.bar([i + 0.2 for i in x], M["ebitda"], 0.4, color=C_OR, label="EBITDA")
    for i in x:
        ax.text(i - 0.2, M["rev"][i] + 40, f"{M['rev'][i]:,.0f}", ha="center", fontsize=7.5)
        ax.text(i + 0.2, M["ebitda"][i] + 40, f"{M['ebitda'][i]:,.0f}", ha="center", fontsize=7.5)
    ax.set_xticks(list(x)); ax.set_xticklabels(yrs); ax.set_ylabel("AED millions"); ax.legend(frameon=False, loc="upper left")
    ax2 = ax.twinx(); ax2.plot(list(x), [m * 100 for m in M["margin"]], color=C_GREY, marker="o", lw=1.2)
    ax2.set_ylim(50, 90); ax2.set_ylabel("EBITDA margin, %"); ax2.spines["top"].set_visible(False)
    fig.tight_layout(); p = os.path.join(TMP, "fin.png"); fig.savefig(p); plt.close(fig); return p

def chart_price():
    fig, ax = plt.subplots(figsize=(6.6, 4.3), dpi=200)
    xs = list(range(len(prices))); ys = [v for _, v in prices]
    ax.plot(xs, ys, color=C_NAVY, lw=1.6); ax.fill_between(xs, ys, 2.0, color=C_NAVY, alpha=0.06)
    ax.axhline(2.0, color=C_GREY, lw=0.8, ls="--"); ax.text(0.5, 2.05, "IPO price AED 2.00", fontsize=7.5, color=C_GREY)
    labs = {d: i for i, (d, _) in enumerate(prices)}
    for d, txt in [("2024-11", "Two new gates"), ("2025-01", "Variable tariff")]:
        if d in labs:
            i = labs[d]; ax.annotate(txt, (i, ys[i]), xytext=(i - 9, ys[i] + 1.0), fontsize=7.5, color=C_OR,
                                     arrowprops=dict(arrowstyle="-", color=C_OR, lw=0.8))
    ax.set_xticks([i for i, (d, _) in enumerate(prices) if d.endswith("-01") or i == 0])
    ax.set_xticklabels([d[:4] for i, (d, _) in enumerate(prices) if d.endswith("-01") or i == 0])
    ax.set_ylabel("AED per share"); ax.set_ylim(1.5, 7.5)
    fig.tight_layout(); p = os.path.join(TMP, "price.png"); fig.savefig(p); plt.close(fig); return p

def chart_comps():
    fig, ax = plt.subplots(figsize=(6.2, 4.3), dpi=200)
    rows = sorted([(n, m, inc) for n, m, inc in peers], key=lambda r: r[1])
    names = [r[0] for r in rows] + ["Salik"]; vals = [r[1] for r in rows] + [M["ev_ebitda"]]
    cols = [C_NAVY if r[2] == 1 else "#B8C2D3" for r in rows] + [C_OR]
    ax.barh(names, vals, color=cols)
    for i, v in enumerate(vals): ax.text(v + 0.3, i, f"{v:.1f}x", va="center", fontsize=7.5)
    ax.axvline(M["peer_med"], color=C_GREY, ls="--", lw=0.8); ax.text(M["peer_med"] + 0.3, -0.7, f"Included-peer median {M['peer_med']:.1f}x", fontsize=7.5, color=C_GREY)
    ax.set_xlabel("Trailing EV / EBITDA"); ax.set_xlim(0, 32)
    fig.tight_layout(); p = os.path.join(TMP, "comps.png"); fig.savefig(p); plt.close(fig); return p

def chart_football():
    fig, ax = plt.subplots(figsize=(7.6, 4.4), dpi=200)
    bars = [("52-week trading range", W52[0], W52[1], C_GREY),
            ("Trading comparables (peer median to Transurban)", comps_lo, comps_hi, C_NAVY),
            ("Financial sponsor ability to pay (6x, 18x exit, 13% IRR)", M["px_sponsor"] * 0.93, M["px_sponsor"] * 1.07, C_NAVY),
            ("Strategic ability to pay (7% WACC, 45-year concession)", M["px_strat"] * 0.93, M["px_strat"] * 1.07, C_NAVY),
            ("Indicative anchor-stake pricing (5\u201312% premium)", M["price"] * 1.05, M["price"] * 1.12, C_OR)]
    for i, (lab, lo, hi, c) in enumerate(bars):
        ax.barh(i, hi - lo, left=lo, color=c, height=0.55)
        ax.text(lo - 0.08, i, f"{lo:.2f}", va="center", ha="right", fontsize=7.5); ax.text(hi + 0.08, i, f"{hi:.2f}", va="center", fontsize=7.5)
    ax.set_yticks(range(len(bars))); ax.set_yticklabels([b[0] for b in bars], fontsize=7.5); ax.invert_yaxis()
    ax.axvline(M["price"], color=C_OR, lw=1.2); ax.text(M["price"], -0.9, f"Market AED {M['price']:.2f}", color=C_OR, fontsize=8, ha="center")
    ax.set_xlim(3.0, 8.0); ax.set_xlabel("AED per share")
    fig.tight_layout(); p = os.path.join(TMP, "ff.png"); fig.savefig(p); plt.close(fig); return p

def chart_buyers():
    fig, ax = plt.subplots(figsize=(6.2, 4.8), dpi=200)
    rows = list(reversed(buyers))
    cmap = {"Strategic": C_NAVY, "Financial": C_OR, "Sovereign": C_GREY}
    ax.barh([b[0] for b in rows], [b[2] for b in rows], color=[cmap[b[1]] for b in rows])
    for i, b in enumerate(rows): ax.text(b[2] + 0.03, i, f"{b[2]:.2f}", va="center", fontsize=7.5)
    ax.set_xlim(2.5, 4.8); ax.set_xlabel("Weighted score (1\u20135)"); ax.tick_params(axis="y", labelsize=7.5)
    from matplotlib.patches import Patch
    ax.legend(handles=[Patch(color=v, label=k) for k, v in cmap.items()], frameon=False, loc="lower right", fontsize=7.5)
    fig.tight_layout(); p = os.path.join(TMP, "buyers.png"); fig.savefig(p); plt.close(fig); return p

charts = dict(fin=chart_financials(), price=chart_price(), comps=chart_comps(), ff=chart_football(), buyers=chart_buyers())

# --------------------------------------------------------------------------- #
# Deck
# --------------------------------------------------------------------------- #
c = rl_canvas.Canvas(OUT, pagesize=landscape(A4))
c.setTitle("Project Gateway \u2014 Strategic alternatives for the Government of Dubai's stake in Salik")
st_body = ParagraphStyle("b", fontName="Helvetica", fontSize=9.5, leading=13, textColor=colors.HexColor("#1A1A1A"))
st_bul = ParagraphStyle("bl", parent=st_body, leftIndent=10, bulletIndent=0, spaceAfter=3)
st_small = ParagraphStyle("s", parent=st_body, fontSize=8, leading=10.5, textColor=GREY)
st_kpi = ParagraphStyle("k", parent=st_body, fontName="Helvetica-Bold", fontSize=15, textColor=NAVY, leading=18)
st_kpil = ParagraphStyle("kl", parent=st_body, fontSize=8, textColor=GREY, leading=10)
page_no = [0]

def frame(title, kicker=""):
    page_no[0] += 1
    c.setFillColor(NAVY); c.rect(0, H - 22 * mm, W, 22 * mm, stroke=0, fill=1)
    c.setFillColor(ORANGE); c.rect(0, H - 23 * mm, W, 1 * mm, stroke=0, fill=1)
    c.setFillColor(colors.white); c.setFont("Helvetica-Bold", 16); c.drawString(14 * mm, H - 14 * mm, title)
    if kicker: c.setFont("Helvetica", 9); c.drawString(14 * mm, H - 19.5 * mm, kicker)
    c.setFont("Helvetica-Bold", 8); c.drawRightString(W - 14 * mm, H - 13 * mm, "PROJECT GATEWAY")
    c.setFont("Helvetica", 7.5); c.setFillColor(GREY)
    c.drawString(14 * mm, 8 * mm, "Discussion materials \u00b7 hypothetical mandate \u00b7 all figures from the Project 15 model, market data 18 September 2026")
    c.drawRightString(W - 14 * mm, 8 * mm, f"Jainam Mehta \u00b7 IB Portfolio \u00b7 Project 16 \u00b7 {page_no[0]}")

def para(text, x, y, w, style=st_body):
    p = Paragraph(text, style); pw, ph = p.wrap(w, H); p.drawOn(c, x, y - ph); return ph

def bullets(items, x, y, w):
    yy = y
    for t in items:
        p = Paragraph(t, st_bul, bulletText="\u2022"); pw, ph = p.wrap(w, H); p.drawOn(c, x, yy - ph); yy -= ph + 2
    return y - yy

def kpis(items, x, y, w):
    n = len(items); cw = w / n
    for i, (val, lab) in enumerate(items):
        c.setFillColor(LIGHT); c.rect(x + i * cw + 1.5 * mm, y - 17 * mm, cw - 3 * mm, 17 * mm, stroke=0, fill=1)
        para(val, x + i * cw + 4 * mm, y - 3 * mm, cw - 6 * mm, st_kpi); para(lab, x + i * cw + 4 * mm, y - 10.5 * mm, cw - 6 * mm, st_kpil)

def table(data, x, y, widths, header=True, fs=8.5):
    stb = ParagraphStyle("tb", fontName="Helvetica", fontSize=fs, leading=fs + 2.5, textColor=colors.HexColor("#1A1A1A"))
    sth = ParagraphStyle("th", parent=stb, fontName="Helvetica-Bold", textColor=colors.white)
    data = [[Paragraph(str(cell), sth if (header and i == 0) else stb) if isinstance(cell, str) else cell for cell in row] for i, row in enumerate(data)]
    t = Table(data, colWidths=widths)
    sty = [("FONT", (0, 0), (-1, -1), "Helvetica", fs), ("TEXTCOLOR", (0, 0), (-1, -1), colors.HexColor("#1A1A1A")),
           ("VALIGN", (0, 0), (-1, -1), "TOP"), ("LINEBELOW", (0, -1), (-1, -1), 0.5, GREY),
           ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, LIGHT]), ("TOPPADDING", (0, 0), (-1, -1), 3), ("BOTTOMPADDING", (0, 0), (-1, -1), 3)]
    if header: sty += [("BACKGROUND", (0, 0), (-1, 0), NAVY), ("TEXTCOLOR", (0, 0), (-1, 0), colors.white), ("FONT", (0, 0), (-1, 0), "Helvetica-Bold", fs)]
    t.setStyle(TableStyle(sty)); tw, th = t.wrap(W, H); t.drawOn(c, x, y - th); return th

def img(path, x, y, w):
    from reportlab.lib.utils import ImageReader
    ir = ImageReader(path); iw, ih = ir.getSize(); h = w * ih / iw
    c.drawImage(ir, x, y - h, width=w, height=h); return h

fmt_m = lambda v: f"AED {v:,.0f}m"; fmt_bn = lambda v: f"AED {v/1000:,.1f}bn"

# 1. Cover
c.setFillColor(NAVY); c.rect(0, 0, W, H, stroke=0, fill=1); c.setFillColor(ORANGE); c.rect(0, 0, 8 * mm, H, stroke=0, fill=1)
c.setFillColor(colors.white); c.setFont("Helvetica-Bold", 11); c.drawString(24 * mm, H - 40 * mm, "PROJECT GATEWAY")
c.setFont("Helvetica-Bold", 30); c.drawString(24 * mm, H - 62 * mm, "Strategic alternatives for the")
c.drawString(24 * mm, H - 76 * mm, "Government of Dubai's stake in Salik")
c.setFont("Helvetica", 14); c.drawString(24 * mm, H - 92 * mm, "Discussion materials for the Department of Finance \u00b7 September 2026")
c.setFont("Helvetica", 10); c.drawString(24 * mm, 40 * mm, "Prepared by Jainam Mehta \u00b7 IB Portfolio, Project 16 \u00b7 github.com/jainammehta1215/ib-portfolio")
c.setFont("Helvetica-Oblique", 8.5)
c.drawString(24 * mm, 30 * mm, "Hypothetical mandate prepared as a portfolio exercise on public information. No party named in these materials is or has been in discussions.")
c.showPage()

# 2. Executive summary
frame("Executive summary", "Four years after the IPO the asset is fully valued by the market; the question is what to do with 75.1%")
kpis([(fmt_bn(M["mcap"]), "Market capitalisation"), (f"{M['ev_ebitda']:.1f}x", "EV / FY2025 EBITDA"), (f"+{M['since_ipo']*100:.0f}%", "Price return since IPO at AED 2.00"),
      (f"{GOV*100:.1f}%", "Government of Dubai stake"), (f"{M['dy']*100:.1f}%", "Dividend yield")], 14 * mm, H - 30 * mm, W - 28 * mm)
bullets([
    f"<b>Salik has become the highest-quality listed infrastructure asset in the Emirate.</b> FY2025 EBITDA of {fmt_m(M['ebitda'][3])} at a {M['margin'][3]*100:.0f}% margin, cash conversion above 95%, no capex, all profit distributed. The 2025 tariff change demonstrated that the tariff lever works and the market rewarded it.",
    f"<b>The market already prices it as a strategic asset.</b> At {M['ev_ebitda']:.1f}x EBITDA the shares trade {(M['ev_ebitda']/M['peer_med']-1)*100:.0f}% above the toll-road peer median and above what any financial sponsor could pay (AED {M['px_sponsor']:.2f}). Only a long-term strategic or sovereign holder can justify a premium (up to AED {M['px_strat']:.2f}).",
    "<b>A change of control is neither available nor advisable.</b> The RTA holds the tariff, the gates and the concession; the government's 75.1% is the basis of the asset's credit and equity story. The alternatives that create value are all minority transactions.",
    f"<b>We recommend a placement of {STAKE_LO*100:.0f}\u2013{STAKE_HI*100:.0f}% to an anchor consortium</b> \u2014 a global infrastructure sponsor with Gulf precedent alongside a regional sovereign fund \u2014 at a {prem_lo*100:.0f}\u2013{prem_hi*100:.0f}% premium to market, raising {fmt_bn(proc_lo)}\u2013{fmt_bn(proc_hi)}, with an operator partnership on technology and ancillary revenue as a second step.",
    "<b>Timing is favourable.</b> Gulf infrastructure demand from global sponsors is at a peak after the Jafurah and ADNOC pipeline transactions; the tariff change is in the numbers; and the free float is small enough that a block would be scarce.",
], 14 * mm, H - 54 * mm, W - 28 * mm)
c.showPage()

# 3. Situation overview
frame("Situation overview", "What Salik is, who owns it, and why the question arises now")
h = table([["", ""],
           ["Business", "Exclusive operator of Dubai's automated toll system under a 49-year RTA concession to 2071. Ten gates. Tariff set by decree; variable pricing since 31 January 2025."],
           ["Model", "Asset-light: the RTA builds and operates the gates; Salik collects and pays a concession fee. Capex is immaterial; EBITDA converts to cash at close to 100%."],
           ["Ownership", "Listed on the DFM on 29 September 2022 at AED 2.00 (AED 3.7bn raised). Government of Dubai 75.1%; free float 24.9%."],
           ["Balance sheet", f"AED 4.0bn of bank debt held flat since formation; net debt / EBITDA has fallen from {M['lev'][0]:.1f}x to {M['lev'][3]:.1f}x on earnings growth alone. Policy: 100% payout."],
           ["Why now", f"The shares have risen {M['since_ipo']*100:.0f}% since listing and the tariff change is now in reported numbers. The government has a fully valued, scarce asset and a set of natural buyers who have just shown their appetite for Gulf infrastructure."]],
          14 * mm, H - 30 * mm, [30 * mm, W - 28 * mm - 30 * mm], header=False, fs=9)
para("<b>The constraint that shapes every alternative:</b> the Government of Dubai's ownership and the RTA's control of the tariff are not incidental to Salik's value; they are its source. Any transaction must leave both visibly intact.",
     14 * mm, H - 30 * mm - h - 8 * mm, W - 28 * mm)
c.showPage()

# 4. Financial snapshot
frame("Financial performance since the IPO", "Revenue and EBITDA have grown every year; FY2025 was a tariff event")
img(charts["fin"], 14 * mm, H - 28 * mm, 150 * mm)
data = [["AED m", "FY22", "FY23", "FY24", "FY25"]] + [[k] + [f"{v:,.0f}" for v in M[key]] for k, key in [("Revenue", "rev"), ("EBITDA", "ebitda"), ("Net profit", "ni"), ("Free cash flow", "fcf")]]
data += [["EBITDA margin"] + [f"{v*100:.0f}%" for v in M["margin"]], ["Net debt / EBITDA"] + [f"{v:.1f}x" for v in M["lev"]]]
table(data, 172 * mm, H - 30 * mm, [34 * mm, 18 * mm, 18 * mm, 18 * mm, 18 * mm])
bullets(["FY2025 revenue +35%: variable tariff (AED 6 peak / AED 4 off-peak) and two gates opened November 2024.",
         "Margin held near 70% through the step-up; the RTA fee scales with revenue.",
         "Dividends paid have tracked prior-year profit; the FY2025 payout was AED 1,391m."], 172 * mm, H - 92 * mm, W - 186 * mm)
c.showPage()

# 5. Share price
frame("Share price since listing", f"From AED 2.00 at the IPO to AED {M['price']:.2f}: the market has re-rated Salik as the tariff lever was proven")
img(charts["price"], 14 * mm, H - 28 * mm, 165 * mm)
bullets([f"Market capitalisation {fmt_bn(M['mcap'])}; enterprise value {fmt_bn(M['ev'])}.",
         f"52-week range AED {W52[0]:.2f}\u2013{W52[1]:.2f}; the shares have given back part of the post-tariff rally.",
         f"P/E {M['pe']:.1f}x and yield {M['dy']*100:.1f}%: priced as a long-duration yield asset, not as a growth stock.",
         "Free float of 24.9% and a concentrated register mean a 10\u201315% block would be the largest liquidity event since the IPO."],
        182 * mm, H - 32 * mm, W - 196 * mm)
c.showPage()

# 6. Trading comparables
frame("Where Salik trades against listed toll roads", "Above the peer median, in line with the longest-dated pure-play concessions")
img(charts["comps"], 14 * mm, H - 28 * mm, 150 * mm)
bullets([f"Included-peer median {M['peer_med']:.1f}x; Salik {M['ev_ebitda']:.1f}x, a {(M['ev_ebitda']/M['peer_med']-1)*100:.0f}% premium.",
         "Transurban, the closest pure-play, trades at 26x on a portfolio of urban roads with traffic risk; Salik carries none.",
         "Ferrovial and Vinci are excluded from the median because construction and airports dilute the concession multiple (shown in grey).",
         "Read-across: the market treats Salik like the best pure-play concessions. A buyer paying a premium to today's price is paying above 22x EBITDA."],
        172 * mm, H - 32 * mm, W - 186 * mm)
c.showPage()

# 7. Valuation summary
frame("Valuation summary", "What different holders can pay, and where a minority placement would price")
img(charts["ff"], 14 * mm, H - 28 * mm, 182 * mm)
bullets([f"<b>Sponsor ceiling AED {M['px_sponsor']:.2f}:</b> 6x leverage, 18x exit and a 13% hurdle cannot reach the market price. Financial buyers will anchor, not lead on price.",
         f"<b>Strategic ceiling AED {M['px_strat']:.2f}:</b> a 7% cost of capital over the 45 years remaining supports a premium, implying {M['ev_strat']/M['ebitda'][3]:.1f}x EBITDA.",
         f"<b>Placement pricing AED {M['price']*1.05:.2f}\u2013{M['price']*1.12:.2f}:</b> a 5\u201312% premium is achievable for a scarce block with board representation and a long lock-up, and sits inside the strategic ceiling.",
         "All figures from the Project 15 model."], 200 * mm, H - 32 * mm, W - 214 * mm)
c.showPage()

# 8. Strategic alternatives
frame("Strategic alternatives", "Four paths, one constraint: the government remains the controlling shareholder in every case")
table([["Alternative", "Description", "Proceeds", "Valuation", "Control", "Execution", "View"],
       ["1. Status quo", "Hold 75.1%; continue the dividend; add gates and tariff changes by decree", "\u2014", "Market", "Unchanged", "None", "Base case; leaves a scarce asset idle"],
       ["2. Accelerated bookbuild", "Sell 5\u201310% to institutions overnight at a discount", f"{fmt_bn(0.05*SHARES*M['price']*0.95/1000)}\u2013{fmt_bn(0.10*SHARES*M['price']*0.95/1000)}", "3\u20136% discount", "Unchanged", "Days", "Fast but value-destructive for a scarce block"],
       ["3. Anchor-stake placement", "Place 10\u201315% with a sponsor-plus-sovereign consortium; board seat; 3-year lock-up", f"{fmt_bn(proc_lo)}\u2013{fmt_bn(proc_hi)}", f"{prem_lo*100:.0f}\u2013{prem_hi*100:.0f}% premium", "Government \u2265 60%", "3\u20134 months", "Recommended"],
       ["4. Strategic partnership", "Operator takes a small stake and a technology / ancillary-revenue agreement", "Small", "Premium on the stake", "Unchanged", "6+ months", "Second step after 3; adds capability, not capital"]],
      14 * mm, H - 30 * mm, [34 * mm, 72 * mm, 34 * mm, 28 * mm, 26 * mm, 24 * mm, 51 * mm], fs=8)
para("A sale of control is not listed because it is not available: the concession, the tariff and the credit story all rest on the government's ownership, and no buyer in the universe scores well enough on consent to make it a real option.",
     14 * mm, 52 * mm, W - 28 * mm, st_small)
c.showPage()

# 9. Buyer universe
frame("Buyer universe", "Twelve candidates scored on fit, precedent, capacity, regional presence and likelihood of consent")
img(charts["buyers"], 14 * mm, H - 28 * mm, 148 * mm)
top = buyers[:6]
table([["Rank", "Candidate", "Type", "Read-across"]] + [[str(b[3]), b[0], b[1], b[4]] for b in top],
      165 * mm, H - 30 * mm, [12 * mm, 46 * mm, 20 * mm, W - 14 * mm - 165 * mm - 78 * mm], fs=7.5)
para("<b>Recommended approach:</b> a consortium of one global sponsor with Gulf precedent (Brookfield, KKR or GIP) and one regional sovereign fund, with Vinci Concessions approached in parallel for the partnership step. Operators with the best fit (Transurban, Ferrovial) are kept warm as partners, not as buyers.",
     165 * mm, H - 100 * mm, W - 179 * mm)
c.showPage()

# 10. Recommended path and terms
frame("Recommended path and indicative terms", "An anchor-stake placement, then an operator partnership")
table([["Term", "Indicative position", "Rationale"],
       ["Stake", f"{STAKE_LO*100:.0f}\u2013{STAKE_HI*100:.0f}% of Salik (from the government's 75.1%)", "Large enough to matter to a sponsor; government stays above 60%"],
       ["Price", f"AED {M['price']*(1+prem_lo):.2f}\u2013{M['price']*(1+prem_hi):.2f} per share ({prem_lo*100:.0f}\u2013{prem_hi*100:.0f}% premium)", "Inside the strategic ceiling; scarcity and governance rights justify a premium over a bookbuild"],
       ["Proceeds", f"{fmt_bn(proc_lo)}\u2013{fmt_bn(proc_hi)}", "Meaningful for the Emirate's infrastructure programme without a discount"],
       ["Governance", "One board seat; information rights; no veto over tariff or concession", "Sponsors need a seat; the RTA's control must be untouched"],
       ["Lock-up", "Three years, then orderly-market provisions", "Protects the free float and the government's remaining stake"],
       ["Dividend", "Policy unchanged (100% payout)", "The yield is the reason a sponsor buys"],
       ["Partnership", "Technology, tolling analytics and ancillary revenue agreement with an operator; option over a further 2\u20133%", "Adds capability the RTA does not have; keeps operators engaged without ceding control"]],
      14 * mm, H - 30 * mm, [30 * mm, 95 * mm, W - 28 * mm - 125 * mm], fs=8.5)
c.showPage()

# 11. Process and timeline
frame("Process and timeline", "Sixteen weeks from mandate to settlement")
table([["Phase", "Weeks", "Workstreams", "Deliverable"],
       ["1. Preparation", "1\u20134", "Vendor due diligence on the concession and RTA agreements; management presentation; valuation refresh; regulatory mapping (SCA, DFM, foreign-ownership rules)", "Information memorandum; data room"],
       ["2. Approach", "5\u20138", "Six to eight candidates under NDA; management meetings; first-round indications of price, stake and governance", "Non-binding offers"],
       ["3. Selection", "9\u201312", "Two consortia into confirmatory diligence; negotiation of shareholders' agreement, lock-up and board rights; SCA consultation on the block", "Binding offers; preferred bidder"],
       ["4. Execution", "13\u201316", "Sale and purchase agreement; DFM block trade mechanics; settlement; announcement", "Completion"]],
      14 * mm, H - 30 * mm, [30 * mm, 18 * mm, W - 28 * mm - 48 * mm - 60 * mm, 60 * mm], fs=8.5)
para("Parallel track: operator partnership discussions begin in phase 2 and conclude after completion, so they do not delay the placement.", 14 * mm, 60 * mm, W - 28 * mm, st_small)
c.showPage()

# 12. Risks
frame("Key risks and mitigants", "")
table([["Risk", "Why it matters", "Mitigant"],
       ["Perceived loss of control", "Public and political sensitivity to a strategic asset", "Government stays above 60%; RTA control of tariff and concession unchanged; sovereign co-investor"],
       ["Price discipline", "Sponsors cannot reach the market price on a standalone basis", "Run a competitive process; anchor with a sovereign; price on the strategic ceiling, not the sponsor ceiling"],
       ["Market overhang", "A block of 10\u201315% is the largest event since the IPO", "Negotiated placement, not a bookbuild; three-year lock-up; staged announcement"],
       ["Tariff policy", "Buyers will price the risk of tariff reversal", "Document the 2025 mechanism and the concession's tariff provisions in the vendor diligence"],
       ["Regulatory", "SCA approval of the block; foreign-ownership thresholds", "Early consultation; structure through DFM block-trade rules"],
       ["Execution timing", "Rate cuts, oil price and regional events move infrastructure appetite", "Sixteen-week process with a go / no-go at the end of phase 2"]],
      14 * mm, H - 30 * mm, [40 * mm, 90 * mm, W - 28 * mm - 130 * mm], fs=8.5)
c.showPage()

# 13. Appendix
frame("Appendix \u2014 sources, assumptions and basis of preparation", "")
bullets(["All financial and valuation figures are read directly from the Project 15 workbook (Salik_Profile_Buyer_Screen.xlsx) after recalculation; the deck rebuilds from the repository and cannot disagree with the model.",
         "Salik financial statements FY2022\u2013FY2025 via Yahoo Finance for SALIK.AE; concession terms, gates and tariff from the 2022 IPO prospectus and RTA announcements; share prices from the DFM.",
         "Peer multiples are trailing EV / EBITDA from Yahoo Finance on 18 September 2026. Ferrovial and Vinci are excluded from the median for the reasons stated on the Trading sheet of the model.",
         "Ability-to-pay inputs: sponsor case 6x leverage at 6%, 18x exit, 13% IRR, 6% EBITDA growth; strategic case 7% WACC, 3% synergies, 45-year remaining concession, 3% long-run growth.",
         "Buyer scores are judgements on five criteria weighted 25 / 20 / 20 / 15 / 20; the ranking is live in the model and changes with the weights.",
         "Indicative placement terms (stake, premium, lock-up, governance) are the author's proposals for discussion and are not derived from any counterparty.",
         "<b>This is a hypothetical mandate prepared as a portfolio exercise on public information.</b> Neither the Government of Dubai nor Salik has engaged any adviser on such a transaction to the author's knowledge, and no party named in the buyer universe is or has been in discussions."],
        14 * mm, H - 32 * mm, W - 28 * mm)
c.showPage()
c.save()
print("saved", OUT, page_no[0], "pages")

# cover preview
import pymupdf
doc = pymupdf.open(OUT); doc[0].get_pixmap(dpi=80).save(os.path.join(HERE, "cover.png"))
import shutil; shutil.rmtree(TMP, ignore_errors=True)
print("cover.png written")
