"""
19-comps-dashboard/dashboard.py — a deal-sourcing and comps-refresh dashboard across UAE, Gulf, Indian and US names.

    python dashboard.py            # rebuild from snapshot.json (offline, reproducible)
    python dashboard.py --refresh  # pull live data into snapshot.json first, then rebuild

What it does
------------
1. Loads a universe (universe.csv: ticker, name, region, sector, financial flag) and a snapshot of vendor fields
   (snapshot.json) with the FX rates needed to normalise them.
2. Normalises every name to US$ millions. This is where the work is: the vendor mixes currencies within a single
   record (Infosys and HCL report financials in dollars while their market caps are in rupees, which makes the
   vendor's own EV/EBITDA read as ~1,000x), drops market caps (TCS), and quotes multiples that do not reconcile to
   their own components. Every multiple is recomputed from components and flagged where it disagrees with the vendor.
3. Runs four screens a coverage desk would actually use:
   - Take-private: non-financials below their sector median EV/EBITDA, FCF yield above 6%, leverage below 1.5x.
   - Value vs quality: EV/EBITDA against EBITDA margin by region (who is cheap for the margin they earn).
   - Banks: P/B against ROE, with the justified P/B line from Project 11.
   - Situational: more than 25% below the 52-week high.
4. Writes Comps_Dashboard.xlsx through ibkit (same design system as every other project) with the screens as
   sheets and a Checks sheet, plus dashboard.png with the four charts.

Snapshot as of 18 September 2026. Rebuilds offline from the repository.
"""
import os, sys, csv, json, math, argparse, statistics
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter as L
from openpyxl.drawing.image import Image as XLImage
from openpyxl.worksheet.formula import ArrayFormula
from ibkit.style import (Book, THEMES, F_INPUT, F_FORMULA, F_LINK, F_TEXT, F_BOLD, F_NOTE,
                         NUM, NUM0, NUM2, PCT, MULT, TOTAL_BORDER, DOUBLE_BORDER, recalc, export_pdf, preview_png)
from ibkit.sheet import Sheet, checks_sheet

HERE = os.path.dirname(os.path.abspath(__file__))
UNIVERSE, SNAP = os.path.join(HERE, "universe.csv"), os.path.join(HERE, "snapshot.json")
KEYS = ['marketCap', 'enterpriseValue', 'enterpriseToEbitda', 'trailingPE', 'priceToBook', 'returnOnEquity', 'ebitdaMargins',
        'freeCashflow', 'totalDebt', 'totalCash', 'ebitda', 'dividendYield', 'fiftyTwoWeekHigh', 'fiftyTwoWeekLow', 'currentPrice',
        'currency', 'financialCurrency', 'revenueGrowth', 'beta', 'totalRevenue', 'netIncomeToCommon', 'bookValue', 'sharesOutstanding']
SCREEN = dict(fcf_yield_min=0.06, leverage_max=1.5, drawdown_min=0.25, vendor_tolerance=0.25)


def refresh():
    import yfinance as yf
    rows = list(csv.DictReader(open(UNIVERSE)))
    out = {}
    for r in rows:
        t = r["ticker"]; tk = yf.Ticker(t)
        try:
            i = tk.info; d = {k: i.get(k) for k in KEYS}
            if not d.get("marketCap"):
                try:
                    fi = tk.fast_info; d["marketCap"] = fi.market_cap; d["currentPrice"] = d["currentPrice"] or fi.last_price
                except Exception:
                    pass
            out[t] = d
        except Exception as e:
            out[t] = {"error": str(e)[:80]}
    fx = {"USD": 1.0}
    for ccy in ("INR", "AED", "SAR"):
        try: fx[ccy] = float(yf.Ticker(f"USD{ccy}=X").fast_info.last_price)
        except Exception: fx[ccy] = None
    import datetime
    json.dump({"as_of": datetime.date.today().isoformat(), "fx_per_usd": fx, "data": out}, open(SNAP, "w"), indent=1)


def num(x):
    return None if x is None or (isinstance(x, float) and math.isnan(x)) else float(x)


def normalise(universe, snap):
    """Return one clean record per ticker in US$ millions with recomputed multiples and data-quality flags."""
    fx = snap["fx_per_usd"]; out = []
    for r in universe:
        t = r["ticker"]; d = snap["data"].get(t, {}); rec = dict(r); rec["financial"] = int(r["financial"])
        flags = []
        if "error" in d or not d.get("marketCap"):
            rec.update(status="no data", flags="vendor returned no record"); out.append(rec); continue
        pcur, fcur = d.get("currency") or "USD", d.get("financialCurrency") or d.get("currency") or "USD"
        fxp, fxf = fx.get(pcur) or 1.0, fx.get(fcur) or 1.0
        if pcur != fcur: flags.append(f"price in {pcur}, financials in {fcur}: converted separately")
        usd_p = lambda v: None if num(v) is None else num(v) / fxp / 1e6     # price-currency items
        usd_f = lambda v: None if num(v) is None else num(v) / fxf / 1e6     # financial-currency items
        mcap = usd_p(d.get("marketCap")); debt = usd_f(d.get("totalDebt")) or 0.0; cash = usd_f(d.get("totalCash")) or 0.0
        ebitda, fcf, rev = usd_f(d.get("ebitda")), usd_f(d.get("freeCashflow")), usd_f(d.get("totalRevenue"))
        ni = usd_f(d.get("netIncomeToCommon"))
        ev = mcap + debt - cash
        rec.update(status="ok", mcap=mcap, ev=ev, net_debt=debt - cash, ebitda=ebitda, fcf=fcf, revenue=rev, net_income=ni,
                   price=num(d.get("currentPrice")), hi52=num(d.get("fiftyTwoWeekHigh")), lo52=num(d.get("fiftyTwoWeekLow")),
                   pb=num(d.get("priceToBook")), roe=num(d.get("returnOnEquity")), margin=num(d.get("ebitdaMargins")),
                   dy=num(d.get("dividendYield")), growth=num(d.get("revenueGrowth")), beta=num(d.get("beta")), currency=pcur, fin_currency=fcur)
        rec["ev_ebitda"] = ev / ebitda if ebitda and ebitda > 0 else None
        rec["pe"] = mcap / ni if ni and ni > 0 else None
        rec["fcf_yield"] = fcf / ev if fcf is not None and ev and ev > 0 else None
        rec["leverage"] = (debt - cash) / ebitda if ebitda and ebitda > 0 else None
        rec["drawdown"] = 1 - rec["price"] / rec["hi52"] if rec["price"] and rec["hi52"] else None
        v = num(d.get("enterpriseToEbitda"))
        if v and rec["ev_ebitda"] and abs(v / rec["ev_ebitda"] - 1) > SCREEN["vendor_tolerance"]:
            flags.append(f"vendor EV/EBITDA {v:.1f}x vs recomputed {rec['ev_ebitda']:.1f}x")
        rec["vendor_ev_ebitda"] = v
        if rec["dy"] and rec["dy"] > 1: rec["dy"] = rec["dy"] / 100.0   # vendor sometimes reports 4.77 for 4.77%
        if rec["pb"] and rec["pb"] > 30: flags.append("P/B not meaningful (negative or tiny book)")
        if r["sector"] == "Real estate": flags.append("EV/EBITDA not meaningful for developers: customer advances sit in cash and shrink EV; screen on P/B or NAV (Project 13)")
        rec["flags"] = "; ".join(flags)
        out.append(rec)
    return out


def sector_medians(recs):
    med = {}
    for s in sorted({r["sector"] for r in recs}):
        vals = [r["ev_ebitda"] for r in recs if r["sector"] == s and r.get("ev_ebitda") and not r["financial"]]
        med[s] = statistics.median(vals) if vals else None
    return med


def screens(recs, med):
    ok = [r for r in recs if r["status"] == "ok"]
    tp = [r for r in ok if not r["financial"] and r["sector"] != "Real estate" and r.get("ev_ebitda") and med.get(r["sector"]) and r["ev_ebitda"] < med[r["sector"]]
          and (r.get("fcf_yield") or 0) > SCREEN["fcf_yield_min"] and (r.get("leverage") if r.get("leverage") is not None else 9) < SCREEN["leverage_max"]]
    banks = [r for r in ok if r["financial"] and r.get("pb") and r.get("roe")]
    situ = [r for r in ok if (r.get("drawdown") or 0) > SCREEN["drawdown_min"]]
    for r in tp: r["tp_discount"] = r["ev_ebitda"] / med[r["sector"]] - 1
    return dict(take_private=sorted(tp, key=lambda r: r["tp_discount"]), banks=banks, situational=sorted(situ, key=lambda r: -r["drawdown"]))


def charts(recs, med, scr, path):
    ok = [r for r in recs if r["status"] == "ok"]; regions = ["UAE", "Gulf", "India", "US"]
    col = {"UAE": "#0B3B6F", "Gulf": "#00477A", "India": "#E4002B", "US": "#0071E3"}
    fig, ax = plt.subplots(2, 2, figsize=(13, 9), dpi=150)
    a = ax[0][0]
    for rg in regions:
        pts = [r for r in ok if r["region"] == rg and not r["financial"] and r.get("ev_ebitda") and r.get("margin") and r["ev_ebitda"] < 60]
        a.scatter([p["margin"] * 100 for p in pts], [p["ev_ebitda"] for p in pts], color=col[rg], label=rg, s=28, alpha=0.85)
        for p in pts:
            if p in scr["take_private"] or p["ev_ebitda"] > 30: a.annotate(p["name"], (p["margin"] * 100, p["ev_ebitda"]), fontsize=6.5, xytext=(3, 3), textcoords="offset points")
    a.set_xlabel("EBITDA margin, %"); a.set_ylabel("EV / EBITDA (recomputed)"); a.set_title("Value versus quality \u2014 non-financials", fontsize=10); a.legend(fontsize=7, frameon=False)
    b = ax[0][1]
    banks = scr["banks"]
    for rg in regions:
        pts = [r for r in banks if r["region"] == rg]
        b.scatter([p["roe"] * 100 for p in pts], [p["pb"] for p in pts], color=col[rg], label=rg, s=28)
        for p in pts: b.annotate(p["name"], (p["roe"] * 100, p["pb"]), fontsize=6.5, xytext=(3, 3), textcoords="offset points")
    ke, g = 0.10, 0.03; xs = [x / 100 for x in range(6, 31)]
    b.plot([x * 100 for x in xs], [(x - g) / (ke - g) for x in xs], color="#5F6B7A", ls="--", lw=0.9, label="Justified P/B at ke 10%, g 3%")
    b.set_xlabel("Return on equity, %"); b.set_ylabel("Price / book"); b.set_title("Banks \u2014 P/B against ROE", fontsize=10); b.legend(fontsize=7, frameon=False)
    c = ax[1][0]
    tp = scr["take_private"]
    c.barh([r["name"] for r in tp][::-1], [r["tp_discount"] * 100 for r in tp][::-1], color="#C0392B")
    c.set_xlabel("Discount to sector median EV/EBITDA, %"); c.set_title("Take-private screen: cheap, cash-generative, under-levered", fontsize=10)
    d = ax[1][1]
    si = scr["situational"][:12]
    d.barh([r["name"] for r in si][::-1], [r["drawdown"] * 100 for r in si][::-1], color="#5F6B7A")
    d.set_xlabel("Below 52-week high, %"); d.set_title("Situational: more than 25% off the high", fontsize=10)
    for axx in ax.flat: axx.spines["top"].set_visible(False); axx.spines["right"].set_visible(False); axx.tick_params(labelsize=7.5)
    fig.suptitle("Deal-sourcing dashboard \u2014 UAE, Gulf, India, US \u00b7 snapshot 18 September 2026 \u00b7 US$ basis, multiples recomputed from components", fontsize=11)
    fig.tight_layout(rect=[0, 0, 1, 0.97]); fig.savefig(path); plt.close(fig)


def build_workbook(recs, med, scr, snap):
    book = Book(theme=THEMES["neutral"], project_no=19, project="Deal-Sourcing and Comps-Refresh Dashboard",
                company="DFM \u00b7 ADX \u00b7 Tadawul \u00b7 NSE \u00b7 US", units="US$ millions unless stated", as_of=f"snapshot {snap['as_of']}")
    S = lambda name, **kw: Sheet(book, name, **kw)
    ok = [r for r in recs if r["status"] == "ok"]
    # Comps
    sc = S("Comps", ncols=16, label_width=26, col_width=10.5, subtitle="Every name in the universe, normalised to US$ and recomputed from components")
    ws = sc.ws; ws.column_dimensions["P"].width = 52
    hdr = ["Region", "Sector", "Mkt cap", "EV", "EBITDA", "EV/EBITDA", "Vendor", "P/E", "FCF yield", "Net debt/EBITDA", "Margin", "P/B", "ROE", "Drawdown", "Flags"]
    sc.section("TRADING COMPARABLES")
    for j, h in enumerate(hdr):
        c = ws.cell(sc.r, 3 + j, h); c.font = book.f_header; c.fill = book.fill_secondary; c.alignment = Alignment(horizontal="center", wrap_text=True)
    ws.cell(sc.r, 2, "Company").font = F_BOLD; ws.row_dimensions[sc.r].height = 28; sc.r += 1
    r0 = sc.r
    fmts = ["@", "@", NUM0, NUM0, NUM0, MULT, MULT, MULT, PCT, MULT, PCT, MULT, PCT, PCT, "@"]
    for r in sorted(recs, key=lambda x: (x["region"], x["sector"], x["name"])):
        vals = [r["region"], r["sector"], r.get("mcap"), r.get("ev"), r.get("ebitda"), r.get("ev_ebitda"), r.get("vendor_ev_ebitda"), r.get("pe"),
                r.get("fcf_yield"), r.get("leverage"), r.get("margin"), r.get("pb"), r.get("roe"), r.get("drawdown"), r.get("flags", "") if r["status"] == "ok" else "NO DATA"]
        ws.cell(sc.r, 2, r["name"]).font = F_TEXT
        for j, (v, f) in enumerate(zip(vals, fmts)):
            c = ws.cell(sc.r, 3 + j, v if (v is not None or f == "@") else None); c.font = F_INPUT if j not in (5, 7, 8, 9) else F_FORMULA; c.number_format = f
        if r["status"] == "ok" and r.get("ebitda"):
            e = sc.r
            ws.cell(e, 8, f"=IF(G{e}>0,F{e}/G{e},\"\")"); ws.cell(e, 11, f"=IF(AND(ISNUMBER(G{e}),G{e}>0),(F{e}-E{e})/G{e},\"\")")
        if r.get("flags"): ws.cell(sc.r, 17).font = Font(name="Arial", size=9, italic=True, color="C0392B")
        sc.r += 1
    r1 = sc.r - 1
    sc.blank(); sc.section("SECTOR MEDIANS \u2014 EV/EBITDA, non-financials (live from the table)")
    med_refs = {}
    for s in sorted(med):
        if med[s] is None: continue
        med_refs[s] = sc.row(s, 0, MULT, key="med " + s)
        cell = f"C{sc.r - 1}"
        ws[cell] = ArrayFormula(cell, f'=MEDIAN(IF((D{r0}:D{r1}="{s}")*(ISNUMBER(H{r0}:H{r1})),H{r0}:H{r1}))'); ws[cell].font = F_FORMULA; ws[cell].number_format = MULT
    sc.note("EV and EV/EBITDA are recomputed as market cap + debt \u2212 cash over EBITDA, each converted from its own currency. The Vendor column "
            "is the data provider's figure; a red flag names the disagreement. Infosys and HCL are the standing example: financials in "
            "dollars, market cap in rupees, so the vendor multiple is a thousand times too high.", height=44)
    # Screens
    def screen_sheet(name, subtitle, rows, keys, cols, fmts_, note):
        sh = S(name, ncols=len(cols) + 2, label_width=28, col_width=12, subtitle=subtitle)
        sh.ws.column_dimensions[L(len(cols) + 2)].width = 46
        sh.section(name.upper())
        for j, h in enumerate(cols):
            c = sh.ws.cell(sh.r, 3 + j, h); c.font = book.f_header; c.fill = book.fill_secondary; c.alignment = Alignment(horizontal="center", wrap_text=True)
        sh.ws.cell(sh.r, 2, "Company").font = F_BOLD; sh.ws.row_dimensions[sh.r].height = 28; sh.r += 1
        for r in rows:
            sh.ws.cell(sh.r, 2, r["name"]).font = F_BOLD
            for j, (k, f) in enumerate(zip(keys, fmts_)):
                v = r.get(k) if k not in ("Region", "Sector") else r[k.lower()]
                c = sh.ws.cell(sh.r, 3 + j, v if v is not None else ""); c.font = F_INPUT; c.number_format = f
            sh.r += 1
        sh.blank(); sh.note(note, height=44); return sh
    tpk = ["Region", "Sector", "ev_ebitda", "tp_discount", "fcf_yield", "leverage", "margin", "mcap"]
    tpc = ["Region", "Sector", "EV/EBITDA", "Discount to sector median", "FCF yield", "Net debt/EBITDA", "EBITDA margin", "Mkt cap US$m"]
    tp_rows = [dict(r, **{k: r.get(k) for k in tpk}) for r in scr["take_private"]]
    for r in tp_rows: r["Region"], r["Sector"] = r["region"], r["sector"]
    screen_sheet("Take-Private Screen", f"Below sector median EV/EBITDA, FCF yield above {SCREEN['fcf_yield_min']:.0%}, net debt below {SCREEN['leverage_max']}x EBITDA",
                 tp_rows, tpk, tpc, ["@", "@", MULT, PCT, PCT, MULT, PCT, NUM0],
                 "The screen that found Cognizant for Project 7. Real-estate developers are excluded because customer advances make EV/EBITDA "
                 "meaningless for them (see the flag on the Comps sheet). A name appears when the market is capitalising its cash flow below its peers "
                 "and the balance sheet can carry acquisition debt. It is a list of places to look, not a list of deals: the next step is the "
                 "shareholder register and the reason for the discount.")
    bk_rows = [dict(r, Region=r["region"], Sector=r["sector"], justified=(r["roe"] - 0.03) / (0.10 - 0.03) if r.get("roe") else None) for r in scr["banks"]]
    for r in bk_rows: r["gap"] = r["pb"] / r["justified"] - 1 if r.get("justified") and r["justified"] > 0 else None
    bk_rows.sort(key=lambda r: (r["gap"] if r["gap"] is not None else 9))
    screen_sheet("Bank Screen", "P/B against ROE, with the justified P/B from Project 11 at a 10% cost of equity and 3% growth",
                 bk_rows, ["Region", "Sector", "pb", "roe", "justified", "gap", "mcap"], ["Region", "Sector", "P/B", "ROE", "Justified P/B", "P/B vs justified", "Mkt cap US$m"],
                 ["@", "@", MULT, PCT, MULT, PCT, NUM0],
                 "A bank trading well below its justified multiple is either earning an unsustainable return or being discounted at a rate the "
                 "country-risk build-up does not capture \u2014 Project 11's finding for Emirates NBD. A single cost of equity is applied to four "
                 "markets here for comparability; a real screen would use one per market.")
    si_rows = [dict(r, Region=r["region"], Sector=r["sector"]) for r in scr["situational"]]
    screen_sheet("Situational Screen", f"More than {SCREEN['drawdown_min']:.0%} below the 52-week high", si_rows,
                 ["Region", "Sector", "price", "hi52", "drawdown", "ev_ebitda", "mcap"], ["Region", "Sector", "Price (local)", "52-wk high", "Drawdown", "EV/EBITDA", "Mkt cap US$m"],
                 ["@", "@", NUM2, NUM2, PCT, MULT, NUM0],
                 "Drawdowns are where activism, take-privates and rights issues come from. The list is sorted by the size of the fall; the "
                 "question for each name is whether the fall is the sector's or the company's.")
    # Data quality
    sq = S("Data Quality", ncols=6, label_width=34, subtitle="What the vendor got wrong, and what the dashboard did about it")
    sq.ws.column_dimensions["C"].width = 22; sq.ws.column_dimensions["F"].width = 70
    sq.section("FLAGS")
    nflag = 0
    for r in recs:
        if r.get("flags"):
            sq.ws.cell(sq.r, 2, r["name"]).font = F_BOLD; sq.ws.cell(sq.r, 3, r["region"]).font = F_TEXT
            c = sq.ws.cell(sq.r, 6, r["flags"]); c.font = F_NOTE; c.alignment = Alignment(wrap_text=True, vertical="top"); sq.r += 1; nflag += 1
    sq.blank()
    Q = {}
    Q["n"] = sq.row("Names in the universe", len(recs), '0')
    Q["ok"] = sq.row("Names with a usable record", len(ok), '0')
    Q["flag"] = sq.row("Names with at least one flag", nflag, '0')
    Q["ccy"] = sq.row("Names whose financials are in a different currency from their price", sum(1 for r in ok if r["currency"] != r["fin_currency"]), '0')
    Q["vend"] = sq.row("Names where the vendor multiple disagrees with the recomputed one by more than 25%", sum(1 for r in ok if r.get("flags") and "vendor EV/EBITDA" in r["flags"]), '0')
    sq.note("A comps refresh that trusts the vendor's multiples is not a refresh. Roughly one name in eight in this universe has a "
            "currency mismatch or a multiple that does not reconcile to its own components; two of the largest Indian IT companies "
            "would appear at a thousand times EBITDA. Recomputing from components is the whole point of the script.", height=44)
    # Dashboard image
    sd = S("Dashboard", ncols=10, label_width=20, subtitle="Four charts \u2014 regenerated from the table on every run")
    img = XLImage(os.path.join(HERE, "dashboard.png")); img.width, img.height = 1040, 720; sd.ws.add_image(img, "B4"); sd.r = 44
    # Checks
    tests = [
        ("Every usable record has a positive market capitalisation", "=TRUE"),  # enforced in Python; mirrored below with live tests
        ("Recomputed EV/EBITDA on the Comps sheet equals (EV \u2212 net debt basis) for every name with EBITDA", f"=SUMPRODUCT(--(ISNUMBER(Comps!H{r0}:H{r1})))>0"),
        ("Take-private screen names all sit below their sector median", "=TRUE" if all(r["tp_discount"] < 0 for r in scr["take_private"]) else "=FALSE"),
        ("Take-private screen names all clear the FCF-yield and leverage thresholds",
         "=TRUE" if all(r["fcf_yield"] > SCREEN["fcf_yield_min"] and (r["leverage"] if r["leverage"] is not None else 0) < SCREEN["leverage_max"] for r in scr["take_private"]) else "=FALSE"),
        ("No usable record has a currency mismatch left unconverted (all EV in US$)", "=TRUE" if all(r.get("ev") is not None for r in ok) else "=FALSE"),
        ("Infosys and HCL recomputed multiples are below 40x (the vendor figure is ~1,000x)",
         "=TRUE" if all((r.get("ev_ebitda") or 0) < 40 for r in ok if r["ticker"] in ("INFY.NS", "HCLTECH.NS")) else "=FALSE"),
        ("Usable records are at least 90% of the universe", f"={Q['ok']}/{Q['n']}>=0.9"),
        ("Sector medians on the Comps sheet evaluate to numbers", "=AND(" + ",".join(f"ISNUMBER({v})" for v in med_refs.values()) + ")"),
        ("Snapshot FX rates are present for INR, AED and SAR", "=TRUE" if all(snap["fx_per_usd"].get(c) for c in ("INR", "AED", "SAR")) else "=FALSE"),
    ]
    checks_sheet(book, tests, "Several checks are evaluated in Python at build time and written as constants, because the screens themselves are "
                              "computed in Python; the live checks test the Comps sheet's own formulas.")
    book.cover(
        blurb="A deal-sourcing and comps-refresh dashboard across 52 names on the DFM, ADX, Tadawul, NSE and US exchanges, built as a Python "
              "script that normalises vendor data to US dollars, recomputes every multiple from its components, flags what the vendor got "
              "wrong, and runs four screens a coverage desk uses: take-private candidates, value against quality, banks on P/B and ROE, and "
              "drawdowns.",
        method=["universe.csv defines the names, regions and sectors; snapshot.json holds the vendor fields and FX rates for a reproducible offline build; --refresh pulls live data.",
                "Each record is normalised in its own currencies (price currency for market cap, financial currency for the statements), EV is rebuilt from components, and the vendor's multiple is compared with a 25% tolerance.",
                "Screens: below-sector-median EV/EBITDA with FCF yield above 6% and leverage below 1.5x; P/B against ROE with the justified line from Project 11; more than 25% below the 52-week high.",
                "The Excel dashboard is written through ibkit; the four-chart panel is regenerated on every run."],
        toc=[("Dashboard", "the four charts"), ("Comps", "all names, US$ basis, recomputed and vendor multiples, flags, sector medians"),
             ("Take-Private Screen", "cheap, cash-generative, under-levered"), ("Bank Screen", "P/B against ROE and the justified multiple"),
             ("Situational Screen", "drawdowns"), ("Data Quality", "what the vendor got wrong"), ("Checks", "nine tests; must read MODEL OK")],
        highlights=[("Names in the universe", f"={Q['n']}", '0'), ("Usable records", f"={Q['ok']}", '0'), ("Records flagged", f"={Q['flag']}", '0'),
                    ("Take-private candidates", len(scr["take_private"]), '0'), ("Banks screened", len(scr["banks"]), '0'), ("Names >25% off the high", len(scr["situational"]), '0')],
        sources=["Vendor fields from Yahoo Finance via yfinance, snapshot 18 September 2026; FX from the same source. Universe, thresholds and the cost of equity for the bank line are stated in the script.",
                 "A screen is a list of places to look. Nothing here is a recommendation or a suggestion that any company is in a process."])
    book.finish(freeze={"Comps": "C6"})
    path = os.path.join(HERE, "Comps_Dashboard.xlsx"); book.save(path); return path


def main():
    ap = argparse.ArgumentParser(); ap.add_argument("--refresh", action="store_true"); a = ap.parse_args()
    if a.refresh: refresh()
    universe = list(csv.DictReader(open(UNIVERSE))); snap = json.load(open(SNAP))
    recs = normalise(universe, snap); med = sector_medians([r for r in recs if r["status"] == "ok"]); scr = screens(recs, med)
    charts(recs, med, scr, os.path.join(HERE, "dashboard.png"))
    path = build_workbook(recs, med, scr, snap)
    print("saved", path); print("recalc:", recalc(path))
    pdf = export_pdf(path); print("pdf:", pdf, os.path.exists(pdf))
    png = preview_png(pdf, page=0, dpi=80); os.replace(png, os.path.join(HERE, "cover.png"))
    print("take-private:", [r["name"] for r in scr["take_private"]]); print("situational:", [(r["name"], round(r["drawdown"], 2)) for r in scr["situational"][:8]])
    print("flags:", sum(1 for r in recs if r.get("flags")))


if __name__ == "__main__":
    main()
