"""
Build guide/IB_Portfolio_Explained.pdf — a plain-English guide to every project in the portfolio,
written for the person who built it, so he can defend it in an interview.
"""
import os
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_LEFT
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, PageBreak, Table, TableStyle,
                                KeepTogether)

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "IB_Portfolio_Explained.pdf")

NAVY = colors.HexColor("#0F2B5B")
GREY = colors.HexColor("#5B6B85")
ACCENT = colors.HexColor("#E31837")
LIGHT = colors.HexColor("#EEF2F8")

ss = getSampleStyleSheet()
S = {}
S["title"] = ParagraphStyle("title", parent=ss["Title"], fontName="Helvetica-Bold", fontSize=26,
                            textColor=NAVY, spaceAfter=6, alignment=TA_LEFT, leading=30)
S["sub"] = ParagraphStyle("sub", parent=ss["Normal"], fontName="Helvetica", fontSize=12,
                          textColor=GREY, spaceAfter=18, leading=16)
S["h1"] = ParagraphStyle("h1", parent=ss["Heading1"], fontName="Helvetica-Bold", fontSize=17,
                         textColor=NAVY, spaceBefore=16, spaceAfter=8, leading=21)
S["h2"] = ParagraphStyle("h2", parent=ss["Heading2"], fontName="Helvetica-Bold", fontSize=12.5,
                         textColor=NAVY, spaceBefore=12, spaceAfter=5, leading=16)
S["h3"] = ParagraphStyle("h3", parent=ss["Heading3"], fontName="Helvetica-BoldOblique", fontSize=10.5,
                         textColor=ACCENT, spaceBefore=9, spaceAfter=3, leading=13)
S["body"] = ParagraphStyle("body", parent=ss["Normal"], fontName="Helvetica", fontSize=10,
                           leading=14.5, spaceAfter=7, textColor=colors.HexColor("#1A1A1A"))
S["bullet"] = ParagraphStyle("bullet", parent=S["body"], leftIndent=12, bulletIndent=2, spaceAfter=4)
S["q"] = ParagraphStyle("q", parent=S["body"], fontName="Helvetica-Bold", textColor=NAVY, spaceAfter=2)
S["a"] = ParagraphStyle("a", parent=S["body"], leftIndent=10, spaceAfter=8)
S["note"] = ParagraphStyle("note", parent=S["body"], fontName="Helvetica-Oblique", fontSize=9.5,
                           textColor=GREY, leading=13)

story = []


def h1(t): story.append(Paragraph(t, S["h1"]))
def h2(t): story.append(Paragraph(t, S["h2"]))
def h3(t): story.append(Paragraph(t, S["h3"]))
def p(t): story.append(Paragraph(t, S["body"]))
def note(t): story.append(Paragraph(t, S["note"]))
def sp(h=6): story.append(Spacer(1, h))
def bul(items):
    for t in items:
        story.append(Paragraph(t, S["bullet"], bulletText="\u2022"))
    sp(4)


def qa(pairs):
    for q, a in pairs:
        story.append(KeepTogether([Paragraph("Q. " + q, S["q"]), Paragraph(a, S["a"])]))


def tbl(data, widths, header=True, small=False):
    t = Table(data, colWidths=widths, hAlign="LEFT")
    style = [("FONT", (0, 0), (-1, -1), "Helvetica", 8.5 if small else 9),
             ("TEXTCOLOR", (0, 0), (-1, -1), colors.HexColor("#1A1A1A")),
             ("VALIGN", (0, 0), (-1, -1), "TOP"),
             ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
             ("LEFTPADDING", (0, 0), (-1, -1), 6), ("RIGHTPADDING", (0, 0), (-1, -1), 6),
             ("LINEBELOW", (0, 0), (-1, -2), 0.4, colors.HexColor("#D5DCE8"))]
    if header:
        style += [("BACKGROUND", (0, 0), (-1, 0), NAVY),
                  ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                  ("FONT", (0, 0), (-1, 0), "Helvetica-Bold", 8.5 if small else 9)]
    t.setStyle(TableStyle(style))
    story.append(t); sp(10)


def callout(title, text):
    inner = [[Paragraph("<b>%s</b>" % title, ParagraphStyle("ct", parent=S["body"], textColor=NAVY, spaceAfter=3))],
             [Paragraph(text, ParagraphStyle("cb", parent=S["body"], fontSize=9.5, leading=13.5, spaceAfter=0))]]
    t = Table(inner, colWidths=[165 * mm], hAlign="LEFT")
    t.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), LIGHT),
                           ("LEFTPADDING", (0, 0), (-1, -1), 10), ("RIGHTPADDING", (0, 0), (-1, -1), 10),
                           ("TOPPADDING", (0, 0), (-1, -1), 7), ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
                           ("LINEBEFORE", (0, 0), (0, -1), 3, ACCENT)]))
    story.append(t); sp(10)


def footer(canvas, doc):
    canvas.saveState()
    canvas.setFont("Helvetica", 7.5)
    canvas.setFillColor(GREY)
    canvas.drawString(20 * mm, 12 * mm, "IB Portfolio Explained  \u00b7  Jainam Mehta  \u00b7  github.com/jainammehta1215/ib-portfolio")
    canvas.drawRightString(190 * mm, 12 * mm, "Page %d" % doc.page)
    canvas.setStrokeColor(colors.HexColor("#D5DCE8"))
    canvas.line(20 * mm, 16 * mm, 190 * mm, 16 * mm)
    canvas.restoreState()


# =========================================================================== #
# COVER
# =========================================================================== #
sp(40)
story.append(Paragraph("IB Portfolio, Explained", S["title"]))
story.append(Paragraph("A plain-English guide to what you built, what the numbers mean, "
                       "and how to answer when someone asks", S["sub"]))
t = Table([[""]], colWidths=[165 * mm], rowHeights=[3])
t.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), ACCENT)]))
story.append(t)
sp(18)
p("This guide has one job: make sure that when an interviewer points at any cell in any of your models, "
  "you can say what it is, where it came from, and why you chose it. A portfolio you cannot defend is worse "
  "than no portfolio, because it turns a neutral first impression into a bad one in about ninety seconds.")
p("It is written in ordinary language. Where a piece of jargon matters, it is explained the first time it "
  "appears and then used normally, because you will need to use it normally too.")
sp(6)
callout("How to use this",
        "Read Part 1 once to fix the vocabulary. Read the Part 2 section for a project the night before "
        "you discuss it. Part 3 is the question bank \u2014 practise saying those answers out loud, not just "
        "reading them. If you can answer Part 3 without looking, you own the portfolio.")
sp(10)
h2("What is in here")
tbl([["Part", "What it covers"],
     ["1", "The vocabulary \u2014 twelve terms that carry almost everything"],
     ["2", "Project by project \u2014 what each one is, in one line and then properly"],
     ["3", "The questions you will actually be asked, with your answers"],
     ["4", "The three rules that make this portfolio different, and why they matter"]],
    [22 * mm, 143 * mm])
note("Figures throughout are the ones actually in your models, as of the 12 September 2026 market date used "
     "across the repository. If you rebuild a model and a number moves, update it here too.")

story.append(PageBreak())

# =========================================================================== #
# PART 1 - VOCABULARY
# =========================================================================== #
h1("Part 1 \u00b7 The vocabulary")
p("Almost all of investment banking valuation runs on about twelve ideas. Everything else is detail.")

h2("1. Equity value and enterprise value")
p("<b>Equity value</b> is what the shareholders own: share price times number of shares. It is what you would "
  "pay to buy all the shares.")
p("<b>Enterprise value</b> is what the whole business is worth regardless of who financed it: equity value "
  "plus debt, minus cash. The logic is that if you buy a company you inherit its debts, so you must pay for "
  "those too \u2014 but you also get its cash, so that comes off.")
callout("The one-line version",
        "Buying a house for 1 million with a 300,000 mortgage attached and 50,000 in cash in the kitchen "
        "drawer: equity value 1 million, enterprise value 1,250,000. You pay the seller, you take on the "
        "mortgage, you keep the cash.")
p("This matters constantly. Profit measures before interest (like EBITDA) belong to everybody \u2014 lenders "
  "and shareholders \u2014 so they pair with enterprise value. Profit measures after interest (like net income) "
  "belong only to shareholders, so they pair with equity value. Mixing them is the most common junior error.")

h2("2. EBITDA")
p("Earnings before interest, tax, depreciation and amortisation. A rough measure of the cash a business "
  "generates from operations before financing and accounting choices get involved. Bankers like it because it "
  "lets you compare two companies with different debt loads and different depreciation policies.")
p("It is not cash flow. It ignores the money spent on equipment and the money tied up in unpaid customer "
  "invoices. Anyone who tells you EBITDA is cash flow has not run a business.")

h2("3. Multiples")
p("A multiple is a price expressed in years of profit. If a company is worth 10 billion and makes 1 billion of "
  "EBITDA, it trades at 10 times EBITDA. Saying \u201c10 times\u201d is quicker than saying the price, and it lets you "
  "compare a large company with a small one.")
p("The two you will use most: <b>EV / EBITDA</b> and <b>P / E</b> (share price divided by earnings per share). "
  "The first values the whole business, the second values the shares only.")

h2("4. Discounted cash flow, or DCF")
p("The idea: a business is worth the cash it will produce in future, but future cash is worth less than cash "
  "today, so you shrink each future year by a discount rate before adding them up.")
p("Money today is worth more for two reasons: you could invest it, and the future is uncertain. The discount "
  "rate carries both.")

h2("5. WACC \u2014 the discount rate")
p("Weighted average cost of capital: the blended return the company's investors require. If a business is "
  "funded half by shareholders wanting 12% and half by lenders charging 6%, its WACC is roughly 9%.")
p("The shareholder half comes from CAPM: risk-free rate, plus beta times the equity risk premium.")
bul(["<b>Risk-free rate</b> \u2014 what the government pays to borrow for ten years. Your safest alternative.",
     "<b>Equity risk premium</b> \u2014 the extra return investors demand for owning shares instead of government bonds. Around 4.5% in developed markets; more where the country is riskier.",
     "<b>Beta</b> \u2014 how much this share moves when the whole market moves. Beta of 1 means it moves with the market; 0.9 means it moves a bit less; 1.5 means it swings harder."])

h2("6. Terminal value")
p("A DCF forecasts five years in detail, then has to account for every year after that. The lump sum "
  "representing all of it is the terminal value. It is usually 60\u201380% of the answer, which is worth admitting "
  "rather than hiding: most of a DCF is an assumption about a distant year, not a forecast.")

h2("7. Free cash flow")
p("The cash left after the business has paid its operating costs, its taxes, bought the equipment it needs, "
  "and funded the extra working capital that growth requires. It is what a DCF discounts, and what an LBO uses "
  "to repay debt.")

h2("8. Working capital")
p("Money stuck in the business day to day \u2014 mostly customers who have not paid yet, less suppliers you have "
  "not paid yet. When a company grows, this number grows, and that growth consumes cash. That is why a "
  "profitable, fast-growing company can still run out of money.")

h2("9. Comparable companies and precedent transactions")
p("<b>Trading comparables</b> ask: what is the stock market paying today for similar businesses? "
  "<b>Precedent transactions</b> ask: what have buyers actually paid to acquire similar businesses?")
p("Precedents are almost always higher, because a buyer taking control pays a premium for it. The gap between "
  "the two is the control premium.")

h2("10. Accretion and dilution")
p("In a takeover: does earnings per share go up or down for the buyer's existing shareholders? If up, the deal "
  "is accretive; if down, dilutive.")
p("The mechanism is simple. The buyer gains the target's profits but usually issues new shares to pay for it. "
  "If profits rise by more than the share count, EPS rises. If not, it falls.")

h2("11. Tangible book value, and why banks are different")
p("Book value is what the accounts say the shareholders own. <b>Tangible</b> book value strips out goodwill and "
  "other intangibles \u2014 things you cannot sell separately.")
p("For banks this is the number that matters, because regulators require a bank to hold real capital against "
  "its loans, and goodwill does not count. A bank deal that destroys tangible book destroys the bank's capacity "
  "to lend, however good the earnings look.")

h2("12. IRR and MOIC")
p("<b>MOIC</b> \u2014 multiple on invested capital. Put in 100, get back 200, that is 2.0 times.")
p("<b>IRR</b> \u2014 internal rate of return: the same result expressed as an annual percentage. 2.0 times over five "
  "years is about 15% a year. Private equity funds typically target 20%+.")
callout("Why both",
        "MOIC ignores time: doubling your money in two years and in ten years are both 2.0 times. IRR fixes "
        "that but flatters short deals. Bankers quote both because neither alone is honest.")

story.append(PageBreak())

# =========================================================================== #
# PART 2 - PROJECTS
# =========================================================================== #
h1("Part 2 \u00b7 Project by project")
p("Twenty projects are built. Each section gives the one-line version, what it actually does, the numbers that "
  "came out, and the single thing that makes it worth showing to someone.")

# --- P1
h2("Project 1 \u00b7 Apple \u2014 three-statement operating model")
h3("One line: a forecast of Apple's profit, balance sheet and cash, wired together so changing one assumption updates all three.")
p("The three statements are the income statement (did we make money), the balance sheet (what do we own and "
  "owe) and the cash flow statement (where did the cash actually go). In a real model they are linked: profit "
  "flows into retained earnings on the balance sheet, and cash flow explains the change in the cash line.")
p("The test of a three-statement model is that the balance sheet balances without you forcing it. If it does, "
  "your links are right. If you plug a number to make it balance, the model is decoration.")
p("Yours has a scenario switch \u2014 one cell changes the whole model between base, upside and downside \u2014 and a "
  "revolver, which is an automatic overdraft that kicks in if the company would otherwise run out of cash.")
h3("Why it matters")
p("This is the foundation model in banking. Everything else \u2014 DCF, LBO, merger \u2014 sits on top of one of these. "
  "Building one that ties is the basic proof that you can do the job.")

# --- P2
h2("Project 2 \u00b7 Apple \u2014 discounted cash flow valuation")
h3("One line: what is Apple actually worth, based on the cash it will generate, rather than what the market says today.")
p("Takes the forecast from Project 1, works out the free cash flow each year, and discounts it back at Apple's "
  "WACC. The WACC was built from scratch rather than assumed: ten peer companies, five years of weekly share "
  "price moves regressed against the market to get each one's beta, adjusted for how much debt each carries, "
  "then the median applied to Apple.")
tbl([["Item", "Value", "Plain English"],
     ["Beta", "1.12", "Apple swings slightly harder than the market"],
     ["Cost of equity", "10.0%", "What shareholders require to hold it"],
     ["WACC", "9.9%", "Apple is almost all equity-funded, so the two are nearly the same"]],
    [40 * mm, 28 * mm, 97 * mm])
h3("Why it matters")
p("Terminal value is calculated two ways \u2014 by assuming perpetual growth, and by assuming a sale at an exit "
  "multiple \u2014 and each is checked against what the other implies. If assuming 3% growth forever implies selling "
  "at 25 times earnings, and you thought 12 times was right, one of your two assumptions is wrong. Most student "
  "DCFs never run that check.")

# --- P3
h2("Project 3 \u00b7 TCS \u2014 trading comparables")
h3("One line: what the stock market pays for IT services companies today, and what that implies TCS is worth.")
p("Eight peers across three tiers: Indian large caps, Indian mid caps, and global names like Accenture and "
  "Cognizant. For each, enterprise value is built up properly and four multiples calculated.")
p("Two technical problems were solved here, and both are worth mentioning in an interview:")
bul(["<b>Currency.</b> Infosys reports its accounts in US dollars but its shares trade in rupees. Multiply those "
     "together carelessly and the multiple is wrong by a factor of 85. The model detects the reporting currency "
     "and normalises everything to dollars before any division happens.",
     "<b>Calendarisation.</b> TCS's year ends in March, Accenture's in August, Cognizant's in December. "
     "Comparing forward multiples across them compares different twelve-month periods. The model weights each "
     "peer's forecast years to line them all up with TCS's March year."])
h3("Why it matters")
p("Anyone can pull multiples off a screen. Knowing that they are not comparable until you fix currency and "
  "fiscal year ends is the difference between a student and an analyst.")

# --- P4
h2("Project 4 \u00b7 IT services \u2014 precedent transactions")
h3("One line: what buyers have actually paid for IT services businesses over ten years.")
p("Fourteen deals from 2014 to 2024 \u2014 Capgemini buying IGATE, Atos buying Syntel, LTI merging with Mindtree, "
  "Cognizant buying Belcan and others \u2014 built from the original announcement press releases and filings rather "
  "than downloaded from a database.")
tbl([["Measure", "25th percentile", "Median", "75th percentile"],
     ["EV / revenue", "1.8x", "2.6x", "3.4x"],
     ["EV / EBITDA", "13.6x", "14.7x", "16.9x"],
     ["One-day premium", "3.5%", "4.6%", "9.0%"]],
    [45 * mm, 40 * mm, 40 * mm, 40 * mm])
h3("Why it matters")
p("The notes say plainly that part of the apparent control premium is not control at all \u2014 it is the cycle. "
  "Those deals were struck when the sector traded at 20\u201330 times earnings; it now trades at 10\u201315. Comparing "
  "deal multiples from one era with trading multiples from another overstates the premium. Spotting that is "
  "the kind of thing that gets remembered.")

# --- P5
h2("Project 5 \u00b7 TCS \u2014 football field and valuation summary")
h3("One line: every method of valuing TCS on one chart, as ranges rather than single numbers.")
p("A football field is the chart bankers put in front of clients: one horizontal bar per valuation method, each "
  "running from low to high. You read the overlaps. Where independent methods agree, you have something.")
tbl([["Method", "Low", "High", "Weight"],
     ["52-week trading range", "1,977", "3,350", "0%"],
     ["Analyst price targets (43)", "1,775", "3,900", "10%"],
     ["DCF", "1,625", "2,504", "35%"],
     ["Trading comps (three multiples)", "1,242", "2,689", "45%"],
     ["Precedent transactions", "1,428", "3,532", "10%"],
     ["Recommended range (weighted)", "1,676", "2,614", ""]],
    [62 * mm, 26 * mm, 26 * mm, 24 * mm])
p("All figures INR per share. The share price on the day was 2,201, which sits 2.6% above the weighted "
  "midpoint of 2,145.")
h3("Why it matters")
p("Two things. First, the weights are inputs on the sheet, not hidden inside a formula \u2014 a client can argue "
  "with them, which is the conversation you want. Second, and better: the standard beta method from Project 2 "
  "was tried here and <b>rejected</b>. Regressing Indian shares against the S&amp;P 500 produced betas of 0.06 to "
  "0.33 with R-squared of 0.00 to 0.04 \u2014 meaning the relationship explains essentially nothing, because the two "
  "markets trade in different hours and different currencies. The regression was redone against the Indian "
  "market instead, giving 0.91 with R-squared 0.28. The rejected attempts are still on the sheet.")
callout("If you remember one thing for interviews",
        "\u201cR-squared\u201d measures how much of one thing's movement is explained by another, from 0 to 1. An "
        "R-squared of 0.01 means your beta is noise. Knowing to throw a result away is a stronger signal than "
        "producing one.")

# --- P6
h2("Project 6 \u00b7 Emirates NBD / Mashreqbank \u2014 merger model")
h3("One line: should Emirates NBD buy Mashreqbank, and the answer is no.")
p("A full accretion/dilution model with real purchase accounting. Bank deals are judged in a specific order, "
  "and the model follows it: tangible book dilution first, regulatory capital second, earnings third.")
tbl([["Test", "Result", "What it means"],
     ["Price paid", "2.16x tangible book", "Emirates NBD's own shares trade at 1.41x"],
     ["EPS, year 3", "(2.3%)", "Still dilutive even at full synergies"],
     ["Break-even synergies", "44.6% of target costs", "Against 25% assumed. Not achievable"],
     ["Tangible book dilution", "(16.7%)", "Shareholders lose a sixth of their tangible book"],
     ["Earnback", "Never", "Earnings never exceed standalone, so it is never recovered"],
     ["Pro forma CET1", "13.0%", "Down 2 points; survives, but only just"],
     ["Return on capital", "8.6%", "Against a 12% required return"]],
    [44 * mm, 38 * mm, 83 * mm], small=True)
h3("Why it matters")
p("An acquirer whose shares trade at 1.4 times tangible book, issuing those shares to buy assets at 2.2 times, "
  "hands value to the seller the moment the deal is announced. No amount of synergy phasing reverses it.")
p("Rather than reporting \u201c2.3% dilutive\u201d and inviting an argument about whether synergies could be higher, "
  "the model solves for the level that would make it work: 44.6%. That closes the discussion. "
  "<b>A merger model that only ever says yes is not being used properly.</b>")
note("CET1 is a bank's core capital as a percentage of its risk-weighted loans. Regulators set a floor; here "
     "it is 11%. Goodwill is deducted in full, which is why an expensive bank acquisition eats capital.")

# --- P7
h2("Project 7 \u00b7 Cognizant \u2014 leveraged buyout")
h3("One line: could a private equity firm take Cognizant private, and what is the most it could pay?")
p("An LBO is buying a company mostly with borrowed money, using the company's own cash flow to repay the debt, "
  "and selling it five years later. The return comes from three places: debt repaid, profit grown, and the "
  "multiple changing.")
tbl([["Item", "Value"],
     ["Offer price (30% premium)", "US$81.36 per share"],
     ["Entry multiple", "9.58x EBITDA"],
     ["Debt raised", "5.25x EBITDA \u2014 US$21.3bn"],
     ["Sponsor equity cheque", "US$18.9bn"],
     ["Exit at 9.00x in year 5", "MOIC 2.00x, IRR 14.9%"],
     ["Versus the 20% hurdle", "Misses by 5.1 points"],
     ["Maximum price at 20% IRR", "US$73.78 \u2014 an 18% premium"]],
    [70 * mm, 95 * mm])
h3("Why it matters")
p("The exit multiple is set <i>below</i> the entry multiple deliberately. Assuming you sell for more than you "
  "paid is how an LBO model is made to produce whatever answer is wanted, and one of the sixteen checks fails "
  "the model if anyone raises it above entry.")
p("The genuinely useful output is not the IRR. It is the maximum price \u2014 US$73.78 \u2014 because that is the floor "
  "under the share price in any sale process. And it sits 18% above where Cognizant actually trades, which "
  "says the market is pricing it below what a leveraged buyer could justify.")


# --- P8
h2("Project 8 \u00b7 Reliance Industries \u2014 debt capacity and capital structure")
h3("One line: how much more could Reliance borrow, and at what rating?")
p("A rating agency does not look at the headline leverage a company reports. It adds back leases, deferred "
  "spectrum payments and other debt-like items, strips one-offs from EBITDA, and asks how much cash is left after "
  "capex to service debt. The model rebuilds Reliance's FY2026 balance sheet on that basis and asks how much "
  "capacity remains at each rating.")
tbl([["Measure", "Result", "What it means"],
     ["Headline net debt / EBITDA", "0.60x", "What the company reports"],
     ["Agency-adjusted net debt / EBITDA", "1.25x", "Leases and spectrum liabilities double it"],
     ["Cash interest cover", "5.0x", "Comfortable, but not the 8x of a decade ago"],
     ["Capacity at an A rating", "Rs 44,272 crore", "Funds-from-operations to debt is the binding test"],
     ["Capacity at BBB", "Rs 1,90,332 crore", "Four times more room one notch down"],
     ["WACC today / at optimum", "12.88% / 12.84%", "Barely any saving from more debt"]],
    [50 * mm, 40 * mm, 75 * mm], small=True)
h3("Why it matters")
p("The finding is that Reliance is already close to its optimal capital structure: the WACC curve is almost flat, "
  "so extra borrowing buys almost nothing. The useful output is the rating map \u2014 which test binds at each notch \u2014 "
  "because that is what a treasurer actually negotiates with. <b>Capacity is set by the weakest ratio, not the average.</b>")

# --- P9
h2("Project 9 \u00b7 DEWA \u2014 rebuilding a real IPO")
h3("One line: was Dubai's largest IPO priced fairly, and what has happened since?")
p("Takes the April 2022 flotation of Dubai's electricity and water utility and rebuilds the pricing from the "
  "prospectus: dividend yield, EBITDA multiple and DCF against the AED 2.25\u20132.48 range, the order book, the "
  "cornerstone investors and the upsizing from 6.5% to 18% of the company.")
tbl([["Item", "Value"],
     ["Priced at", "AED 2.48 \u2014 top of the range"],
     ["Open-book demand", "AED 315bn \u2014 37x covered, reproduced"],
     ["Cornerstones", "AED 13.8bn, 62% of the deal"],
     ["Price versus the midpoint of three methods", "+7.6% \u2014 priced full"],
     ["Day-one close", "AED 2.87, +15.7% \u2014 AED 3.5bn 'left on the table'"],
     ["Total return since, per year", "6.2%, mostly dividends"]],
    [80 * mm, 85 * mm])
h3("Why it matters")
p("An IPO that is 37 times covered and pops 16% looks underpriced. The model says it was priced above the "
  "midpoint of its own valuation methods and that the pop was scarcity: cornerstones took most of the deal and "
  "the free float was small. Four years on, the total return has been a utility's, not a growth stock's. "
  "<b>The order book measures allocation, not value.</b>")

# --- P10
h2("Project 10 \u00b7 Reliance Industries \u2014 sum of the parts")
h3("One line: value each Reliance business on its own peers, add them up, and see what the market is paying for Jio.")
p("Five segments \u2014 Jio, Retail, refining and petrochemicals, oil and gas, and others \u2014 each at the multiple its "
  "listed peers trade on, less net debt and minorities, less a holding-company discount.")
tbl([["Segment", "Multiple", "Peer anchor"],
     ["Jio (digital)", "12x EBITDA", "Bharti Airtel 11.6x"],
     ["Retail", "25x", "Between Walmart and DMart"],
     ["Refining and chemicals", "7.5x", "Indian and US refiners"],
     ["Result", "Rs 1,065 per share", "14.5% below the market"],
     ["Market-implied Jio multiple", "17.3x", "49% premium to Bharti"]],
    [50 * mm, 40 * mm, 75 * mm], small=True)
h3("Why it matters")
p("The number to remember is not the SOTP but the implied multiple: to justify today's price the market must "
  "value Jio at 17 times EBITDA, half again what Bharti trades on, or Retail at 37 times. That reframes the "
  "conversation from 'is Reliance cheap' to 'is Jio worth a 49% premium to its only peer'. <b>A sum of the parts is "
  "most useful run backwards.</b>")

# --- P11
h2("Project 11 \u00b7 Emirates NBD \u2014 valuing a bank")
h3("One line: banks are valued on book value and return on equity, not on EBITDA, and this is the method.")
p("A dividend-discount model and an excess-return model, which are the same thing written two ways: a bank is "
  "worth its book value plus the present value of the returns it earns above its cost of equity. The cost of "
  "equity is built up from the US risk-free rate, a UAE country premium and a bank beta.")
tbl([["Item", "Value"],
     ["Cost of equity", "10.0% (4.1% + 0.9% country + 1.0 \u00d7 5.0%)"],
     ["Return on equity, fading 16.5% to 14%", "Long-run growth 3%"],
     ["Value per share", "AED 41.71 against AED 30.98 \u2014 35% upside"],
     ["ROE the price implies", "12.5%"],
     ["Cost of equity the price implies", "11.1%"]],
    [80 * mm, 85 * mm])
h3("Why it matters")
p("Either the market thinks ENBD's returns will fall to 12.5% \u2014 below anything it has earned since 2016 \u2014 or "
  "it discounts Gulf bank equity at a rate above what a country-risk build-up gives. The second is the honest "
  "reading, and it recurs in Project 19's bank screen. <b>For a bank, price-to-book is ROE against the cost of equity, "
  "and nothing else.</b>")

# --- P12
h2("Project 12 \u00b7 Saudi Aramco \u2014 energy net asset value")
h3("One line: value the barrels under the Kingdom's fiscal terms and find the oil price the share price implies.")
p("Oil companies are valued on the cash their reserves will produce. For Aramco the state's take dominates: a "
  "royalty that steps from 15% to 45% to 80% as Brent passes US$70 and US$100, and a 50% income tax. The model "
  "builds a per-barrel netback under those terms, runs it over the 52-year reserve life, and then across a price deck.")
tbl([["Item", "Value"],
     ["Share of the barrel Aramco keeps at US$70", "33 cents \u2014 US$20.21 on US$60.62"],
     ["Share of a marginal dollar above US$100", "10 cents"],
     ["NAV per share at US$70, 8% discount rate", "SAR 18.06 against SAR 25.56"],
     ["Brent the price implies at 8%", "US$118"],
     ["Discount rate the price implies at US$70", "5.7%"],
     ["Brent at which the base dividend is covered", "US$68.3"],
     ["Model rebuild of FY2025 free cash flow", "Within 6% of reported"]],
    [80 * mm, 85 * mm])
h3("Why it matters")
p("The shares are a claim on volume and cost, not on the oil price: above US$100 the equity gets a dime of each "
  "dollar. The price is not an oil bet either; it implies a discount rate under 6%, which is how the market treats a "
  "quasi-sovereign yield stock. And the dividend, not the NAV, is the live question: it needs high-sixties Brent to "
  "be covered after a US$52bn capex programme. <b>The calibration check \u2014 rebuilding last year's cash flow \u2014 is "
  "what earns the rest.</b>")

# --- P13
h2("Project 13 \u00b7 Emaar Properties \u2014 developer net asset value")
h3("One line: a developer is worth its sales backlog, its rental assets and its land, not its earnings multiple.")
p("Three blocks: the profit embedded in AED 155bn of sold-but-undelivered homes, the recurring income from malls "
  "and hotels at a yield, and the land bank at appraised values with a 50% haircut on the cash it will take to develop.")
tbl([["Item", "Value"],
     ["NAV per share", "AED 20.23 against AED 11.80 \u2014 a 42% discount"],
     ["Development block versus listed Emaar Development", "Within 4.5%"],
     ["Residual land value implied", "AED 767 per square foot of buildable area"]],
    [80 * mm, 85 * mm])
h3("Why it matters")
p("Developers trade at discounts to NAV everywhere, and the discount is the market's view of execution and cycle "
  "risk. The cross-check against the separately listed development arm is what makes the NAV credible: the model's "
  "value for that block lands within 5% of what the market pays for it directly. <b>When part of a sum-of-the-parts is "
  "listed, use it as the check.</b>")

# --- P14
h2("Project 14 \u00b7 Dabur India \u2014 quality of earnings and working-capital peg")
h3("One line: what is the business really earning, and how much working capital does it need?")
p("The report a buyer commissions before agreeing a price. It bridges from the headline EBITDA a seller quotes to "
  "the operating EBITDA a buyer should pay for, tests whether it turns into cash, and sets a normal level of working "
  "capital \u2014 the peg \u2014 so the buyer pays for the business and not for whatever sits on the balance sheet on closing day.")
tbl([["Item", "Value"],
     ["Headline EBITDA, FY2026", "Rs 3,015 crore"],
     ["Adjusted operating EBITDA", "Rs 2,473 crore \u2014 82% of headline"],
     ["What the gap is", "Interest and gains on a Rs 5,148 crore treasury book"],
     ["What the gap is worth at 20x", "Rs 10,800 crore \u2014 twice the cash pile itself"],
     ["Payable days, FY2023 to FY2026", "126 to 157"],
     ["Working-capital peg / closing balance", "Rs 366 crore / Rs 83 crore \u2014 buyer receives Rs 283 crore"]],
    [80 * mm, 85 * mm])
h3("Why it matters")
p("Nearly a fifth of the headline EBITDA is income on cash. In a deal the cash transfers at face value; a buyer who "
  "also capitalises its income at 20 times has paid for the same asset twice. The working-capital question is "
  "smaller in money and larger in negotiation: whether a month of extra payable days is structural or unwinds "
  "after completion. <b>A QoE exists to find the number that is not in the accounts.</b>")

# --- P15 and P16
h2("Projects 15 and 16 \u00b7 Salik \u2014 company profile, buyer screen and pitch book")
h3("One line: who could own Dubai's toll operator, what could they pay, and what would you pitch to the government?")
p("Project 15 is the groundwork: four years of financials, trading against listed toll roads, twelve candidate "
  "buyers scored on five weighted criteria, and what a financial sponsor and a strategic operator could each pay. "
  "Project 16 is the twelve-slide pitch book built on it, with every number read from the Project 15 workbook so "
  "the two cannot disagree.")
tbl([["Item", "Value"],
     ["EV / EBITDA at the market", "20.2x \u2014 a 30% premium to the peer median"],
     ["Price return since the 2022 IPO", "+169%"],
     ["Top of the buyer screen", "Vinci, Brookfield, KKR, Abertis"],
     ["Sponsor ability to pay (6x debt, 18x exit, 13% IRR)", "AED 4.66 \u2014 13% below market"],
     ["Strategic ability to pay (7% WACC, 45-year concession)", "AED 6.05 \u2014 13% above market"],
     ["Pitch book recommendation", "Place 10\u201315% with a sponsor-plus-sovereign anchor at a 5\u201312% premium"]],
    [80 * mm, 85 * mm])
h3("Why it matters")
p("The screen's finding is who cannot buy: the government holds 75% and the regulator sets the tariff, so a "
  "change of control is not on offer and the best operators fall down the list on consent. A sponsor cannot reach "
  "the market price; only a strategic can justify a premium. The pitch follows the model rather than the other way "
  "round. <b>A deck that recommends what its own appendix cannot support is the commonest failure in banking.</b>")

# --- P17
h2("Project 17 \u00b7 Kellanova / Mars \u2014 replicating a fairness opinion")
h3("One line: take a real proxy statement and see whether the bankers' valuation ranges can be rebuilt.")
p("When Mars bought Kellanova for US$83.50 a share, the proxy disclosed how Goldman Sachs and Lazard built their "
  "fairness opinions: projections, discount rates, multiple ranges, premia. The model takes only what was disclosed, "
  "back-solves the two undisclosed inputs from one range, and rebuilds the other eleven.")
tbl([["Analysis", "Proxy range", "Rebuilt"],
     ["Goldman DCF", "$68.42\u201383.82", "$67.79\u201383.19"],
     ["Lazard EBITDA comparables", "$57.20\u201373.80", "$57.01\u201374.10"],
     ["Lazard P/E comparables", "$55.90\u201375.70", "$55.86\u201375.81"],
     ["Both advisors' premia analyses", "as disclosed", "To the cent"],
     ["Lazard DCF", "$69.60\u201379.10", "$61.07\u201378.76 \u2014 low end not reproducible"],
     ["Offer against twelve ranges", "", "Above 7, inside 5, below 0"]],
    [55 * mm, 45 * mm, 65 * mm], small=True)
h3("Why it matters")
p("Almost every range comes back within a dollar or two from a public document. The one that does not \u2014 the "
  "bottom of Lazard's DCF \u2014 needs a terminal cash flow close to NOPAT, which means Lazard normalised the terminal "
  "year and the proxy did not say so. The offer was a full price: top of one DCF range, above the other, above every "
  "trading multiple. <b>A fairness opinion is a set of defensible ranges around a price already negotiated; the "
  "replication shows how tightly they fit.</b>")

# --- P18
h2("Project 18 \u00b7 Dubai World and Nakheel \u2014 restructuring")
h3("One line: the Gulf's largest restructuring rebuilt as liquidity, recoveries and debt-for-equity.")
p("In November 2009 Dubai World asked for a standstill with US$3.5bn due in three weeks against US$0.5bn of cash. "
  "The model shows why a solvent group ran out of time, what each creditor class received in present value once "
  "maturities were pushed out five and eight years at 1% coupons, and what the government's US$8.9bn conversion to "
  "equity was worth.")
tbl([["Item", "Value"],
     ["Asset cover over the plan, forced-sale values", "0.93x with DP World, 0.47x without"],
     ["Bank creditors' recovery, present value", "71 cents \u2014 on a 0% nominal haircut"],
     ["Value transferred from banks by the new terms", "US$2.7bn"],
     ["Nakheel trade creditors (40% cash, 60% 10% sukuk)", "98% \u2014 the best-treated class"],
     ["Government recovery on its conversion, 2010 values", "48%"]],
    [80 * mm, 85 * mm])
h3("Why it matters")
p("Nobody took a nominal haircut except the government, and the banks took a real one: eight years at 1% when the "
  "market required 8% is 70 cents on the dollar, whatever the loan agreement says. Contractors who could stop the "
  "projects did better than lenders who could only sue. <b>In a restructuring, operational leverage beats legal "
  "seniority, and recovery is a present-value number.</b>")

# --- P19
h2("Project 19 \u00b7 Deal-sourcing dashboard")
h3("One line: the Monday-morning tool \u2014 52 companies across five exchanges, normalised, screened, and rebuilt on demand.")
p("A Python script rather than a workbook, because the value is in the refresh. It pulls vendor data for Gulf, "
  "Indian and US names, converts every field in its own currency, recomputes each multiple from components, flags "
  "what the vendor got wrong, and runs four screens into an Excel dashboard and a chart panel.")
tbl([["Item", "Value"],
     ["Records flagged for a data problem", "10 of 48 \u2014 one in five"],
     ["Worst vendor error", "Infosys, HCL at 956x and 1,218x (dollars over rupees); near 10x recomputed"],
     ["Take-private screen", "Cognizant, EPAM, Wipro"],
     ["Banks below the justified price-to-book line", "Dubai Islamic Bank, Emirates NBD, Mashreqbank, HDFC Bank"]],
    [70 * mm, 95 * mm])
h3("Why it matters")
p("A comps refresh that trusts the vendor is not a refresh. The screen that found Cognizant for Project 7 is here as "
  "code, and a developer excluded from it because customer advances make EV/EBITDA meaningless for the sector. "
  "<b>Knowing when a metric does not apply is most of the job.</b>")

# --- P20
h2("Project 20 \u00b7 HDFC Bank / HDFC Ltd \u2014 the capstone")
h3("One line: replicate India's largest merger at announcement, then check what happened.")
p("Reuses the merger toolkit from Project 6 and the bank frame from Project 11 on the April 2022 deal, then does "
  "what a pitch never does: sets the announcement case against three years of reported results and the share price "
  "against the bank index.")
tbl([["Item", "Value"],
     ["Exchange ratio and premium", "1.68x, 3.3%"],
     ["HDFC Ltd holders' ownership / look-through share of profit, book", "41% / 42%, 47%"],
     ["Pro forma EPS before the drag / book accretion", "\u22121.2% / +7.7%"],
     ["Regulatory drag in year one", "Rs 4,006 crore after tax \u2014 29% of HDFC Ltd's profit"],
     ["ROE: standalone FY2023 to FY2026", "15.7% to 12.8%"],
     ["Total shareholder return since the day before announcement", "+3.8% against Nifty Bank +50.9%"],
     ["Price / book, 2022 to today", "3.48x to 2.05x; implied cost of equity 9.4% to 10.1%"]],
    [95 * mm, 70 * mm], small=True)
h3("Why it matters")
p("The ratio was set on market value and favoured the bank. The accretion arithmetic was right: book value per share "
  "compounded at 12%. It was also the wrong question, because the same pro forma said ROE would fall three points, "
  "and a bank's price is book times a multiple set by ROE. The multiple halved; the price went nowhere; the cost of "
  "equity the market applies barely moved. <b>The de-rating was arithmetic, not sentiment \u2014 and it was visible on the "
  "back page of the pitch.</b>")

story.append(PageBreak())

# =========================================================================== #
# PART 3 - QUESTIONS
# =========================================================================== #
h1("Part 3 \u00b7 The questions you will be asked")
p("These are the questions an interviewer reaches for when they see this repository. Practise them out loud.")

h2("General")
qa([("Walk me through a DCF.",
     "Forecast free cash flow for five years. Discount each year back at WACC. Calculate a terminal value for "
     "everything after year five, either by assuming perpetual growth or by applying an exit multiple, and "
     "discount that back too. Add them up to get enterprise value. Add net cash or subtract net debt to get "
     "equity value, then divide by shares for a per-share number."),
    ("Why would you use EBITDA rather than net income?",
     "EBITDA sits above interest and tax, so it is not affected by how a company is financed or where it is "
     "domiciled. That makes two companies with different debt loads comparable. The cost is that it ignores "
     "capital spending and working capital, so it is not cash flow."),
    ("Which is higher, precedent transaction multiples or trading multiples, and why?",
     "Precedents, normally, because an acquirer taking control pays a premium for synergies and for the right "
     "to run the business. In my Project 4 the gap is 2.6 times revenue against 1.8 times. But I would add "
     "that part of that gap is the cycle, not control \u2014 those deals were struck when the sector traded much "
     "higher, and comparing across eras overstates the premium."),
    ("A company has negative free cash flow but is profitable. How?",
     "Usually working capital. If it is growing fast, customers owe it more each month and inventory rises, "
     "and that swallows cash before it ever reaches the bank. Heavy capital spending does the same.")])

h2("Project 2 \u2014 the Apple DCF")
qa([("Where did your WACC come from?",
     "Bottom-up. I took ten peers, regressed five years of weekly returns against the S&amp;P 500 to get each "
     "one's beta, unlevered each at its own net debt to equity, took the median, then relevered at Apple's "
     "capital structure \u2014 which is zero, because Apple is net cash. Then Blume-adjusted, which is two-thirds "
     "raw beta plus one third, the way banks do it. That gave 1.12, a cost of equity of 10.0% on a 4.95% "
     "ten-year Treasury and a 4.5% equity risk premium, and a WACC of 9.9%."),
    ("Why is your WACC almost the same as your cost of equity?",
     "Because equity is about 98% of Apple's capital. There is barely any debt to weight in.")])

h2("Project 5 \u2014 the TCS football field")
qa([("Why is your TCS discount rate 13.3% when Apple's was 9.9%?",
     "Different country and different currency. TCS's cash flows are in rupees, so I discount at a rupee cost "
     "of capital: a 7.02% Indian ten-year government bond rather than a 4.95% US Treasury, and an equity risk "
     "premium of 6.7% rather than 4.5%, because that adds India country risk. Discounting rupee cash flows at "
     "a dollar rate would bury an exchange-rate forecast inside the valuation."),
    ("Why 5.5% terminal growth?",
     "Roughly Indian long-run inflation plus a little real growth, and comfortably below India's nominal GDP "
     "growth \u2014 nothing can grow faster than its economy forever. I also cross-check it: 5.5% implies a 9.1 "
     "times exit multiple, and the 11 times exit multiple implies 6.8% growth. The two are close enough that "
     "neither contradicts the other."),
    ("Your DCF says 1,953 and the stock is 2,201. Is TCS a sell?",
     "I would not put it that strongly. My weighted range is 1,676 to 2,614 and the price sits inside it, "
     "about 2.6% above the midpoint. The honest reading is that the market is paying a modest premium to both "
     "my intrinsic value and the peer group. Whether that is wrong depends on one specific question \u2014 whether "
     "the sector's AI de-rating is permanent \u2014 and my model does not answer that."),
    ("Seventy per cent of your DCF is terminal value. Isn't that a problem?",
     "It is a limitation and I say so on the sheet. It is normal for an asset-light business with low capital "
     "intensity, but it does mean most of the answer is an assumption about 2031. That is exactly why I show "
     "the DCF as a sensitivity range on the football field rather than as a single point.")])

h2("Project 6 \u2014 the bank merger")
qa([("Why do banks care about tangible book value more than EPS?",
     "Because regulatory capital is built from tangible equity. Goodwill is deducted in full from CET1, so "
     "paying a large premium for a bank converts real capital into an intangible that cannot support any "
     "lending. A deal can look accretive to earnings and still leave the bank less able to write loans."),
    ("Your model says the deal fails. Why build it?",
     "Because that is the answer. Emirates NBD trades at 1.4 times tangible book and would be paying 2.2 "
     "times. Rather than stop at \u2018dilutive\u2019, I solved for the synergies needed to break even \u2014 44.6% of "
     "Mashreq's cost base, roughly double what is assumed and above anything disclosed in comparable bank "
     "mergers. In a real process that number is what you take to the client."),
    ("Why did you set revenue synergies to zero?",
     "Because they are the standard way this model gets made to work \u2014 a cross-sell assumption big enough to "
     "close the gap with no mechanism behind it. I left the line in at zero so the choice is visible. It would "
     "not rescue this deal anyway; the shortfall is about AED 700m after tax."),
    ("Why is the consideration 70% stock?",
     "Capital. At that mix pro forma CET1 lands at 13.0%, about two points above the minimum. A cash-heavy "
     "structure would not clear the floor, because cash paid away does not come back as capital.")])

h2("Project 7 \u2014 the LBO")
qa([("What drives returns in an LBO?",
     "Three things: repaying debt with the company's own cash flow, growing EBITDA, and any change in the "
     "multiple between buying and selling. In mine the bridge is roughly US$11.4bn from debt paydown and "
     "US$11.2bn from EBITDA growth, less US$2.4bn lost to multiple contraction."),
    ("Why did you assume a lower exit multiple than entry?",
     "Because assuming multiple expansion is how an LBO model is made to say yes. If the return only works "
     "when you sell higher than you bought, you have not analysed anything. I also put a check in the model "
     "that fails if anyone later raises the exit multiple above entry."),
    ("Why does interest use the opening balance rather than the average?",
     "Average balances make the model circular \u2014 interest drives cash flow, cash flow drives the sweep, the "
     "sweep drives interest. Resolving that needs iterative calculation enabled, and a workbook that only "
     "computes correctly with a setting switched on will break on someone else's machine. Using the opening "
     "balance slightly overstates interest, which is conservative."),
    ("What is a cash sweep?",
     "A term in the loan requiring spare cash to go toward repaying debt rather than to the owners. Mine "
     "applies 100% of free cash flow to the term loan until it is repaid. The notes are bullet, meaning they "
     "cannot be repaid early and are settled at maturity."),
    ("Could a fund really write an US$18.9 billion cheque?",
     "Not alone. It would be a consortium. That changes who signs, not the arithmetic, and I say so in the "
     "model.")])

story.append(PageBreak())

# =========================================================================== #
# PART 4
# =========================================================================== #

h2("Projects 8 to 20")
qa([("Why is your Aramco NAV 30% below the share price? Are you saying it is overvalued?",
     "No. At an 8% discount rate and US$70 Brent the reserves are worth SAR 18 a share, but the same model says the "
     "price implies a 5.7% discount rate, and that is the credible reading: the state owns most of the float and the "
     "market treats the stock as a quasi-sovereign yield instrument. The model reports both. What I would defend is "
     "the netback: it rebuilds Aramco's reported free cash flow within 6%."),
    ("In the Dabur QoE you remove interest income from EBITDA. Isn't that income real?",
     "It is real, and the buyer gets the cash that earns it at face value on closing. What the buyer should not do is "
     "also pay a 20 times multiple for the income on cash it has just bought at one times. Removing it from EBITDA and "
     "adding the cash to the price is the standard treatment; leaving it in double-counts Rs 5,000 crore of cash."),
    ("Your Salik pitch recommends a minority stake, not a sale. Isn't a sale the bigger fee?",
     "The model says no sponsor can reach the market price and the government will not sell control of a tariff-setting "
     "asset. A pitch that recommends a sale of control would contradict its own appendix. The credible mandate is the "
     "anchor placement with an operator partnership behind it, and a banker who pitches the deal that can happen wins "
     "the one after it."),
    ("You could not reproduce Lazard's DCF low end. Doesn't that undermine the replication?",
     "It is the most useful result in it. Eleven of twelve ranges rebuild within a dollar or two. The one that does not "
     "needs a terminal cash flow close to NOPAT rather than the projected free cash flow, which tells you Lazard "
     "normalised the terminal year and the proxy did not disclose it. I would rather report that than force a fit."),
    ("In Dubai World the banks got 100 cents. Why do you say 71?",
     "Because five to eight years at a 1% coupon, when comparable paper yielded 8%, is worth 71 cents on the dollar "
     "today whatever the face value says. The banks accepted it because regulatory accounting let them hold the paper "
     "near par, and that is exactly why they would not take equity. Nominal recovery and economic recovery are "
     "different numbers, and a restructuring adviser prices the second."),
    ("HDFC Bank's merger was accretive to book and EPS. Why has the stock done nothing?",
     "Because a bank is priced on book times a multiple set by ROE, and the merger cut ROE by three points while growing "
     "the book. The multiple went from 3.5 to 2 times; the cost of equity the price implies barely moved. Everything the "
     "announcement promised happened; it just was not what the price depends on. The deal will be judged on whether "
     "deposits replace HDFC Ltd's bonds fast enough to rebuild the return.")])

story.append(PageBreak())
h1("Part 4 \u00b7 The three rules")
p("These are what separate this repository from the hundreds of others with the same project names on them. "
  "Be able to say why each one matters.")

h2("1. Every number is a live formula, and a checks sheet proves it")
p("No answer is typed in. Every model has a Checks sheet that has to read MODEL OK, testing things like: does "
  "the balance sheet balance, does enterprise value equal the sum of its parts, does the money multiple agree "
  "with the IRR, is every forecast cash flow positive.")
p("Some checks are not arithmetic at all \u2014 they are discipline. Project 5 fails if forecast revenue growth "
  "ever exceeds anything TCS achieved in the last three years. Project 7 fails if the exit multiple is set "
  "above entry. Those exist because quietly lifting an assumption until the answer looks right is the most "
  "common way a model goes wrong, and a check is cheaper than a reviewer.")

h2("2. Every source is named, and the weak ones are flagged")
p("Where a number is approximate, it says so. Project 4 marks which target revenues came from rounded press "
  "releases. Project 6 flags that the credit mark is an assumption on an estimated loan book and the largest "
  "judgement in the model. Project 5 states that no broker-target bar could be verified from a free source.")
p("Admitting the weak spots before someone finds them is what makes the strong parts believable.")

h2("3. Every project says what it cannot do")
p("Each notes page has a section on what it would take to do the analysis properly, and what the analysis does "
  "not support. Project 4 says its TCS range is illustrative because no target in the set is remotely TCS's "
  "size. Project 5 says 70% of the DCF is terminal value. Project 7 says the leverage sensitivity flatters "
  "high-leverage cases and explains why.")
sp(4)
callout("The point of all three",
        "An interviewer is not checking whether your model is right. They are checking whether you know where "
        "it is wrong. Every senior banker has seen a confident model blow up. What they are hiring for is "
        "judgement about limits, and this portfolio is built to demonstrate exactly that.")

sp(14)
h2("Before any interview")
bul(["Open the model you are about to discuss and click through the Checks sheet. Know what each check tests.",
     "Know your three headline numbers cold: WACC, the valuation conclusion, and the one assumption the answer is most sensitive to.",
     "Be ready to say what you would do differently with more time. It is the question most candidates fumble and it is written on every notes page.",
     "Never claim a number you cannot source. \u2018I would have to check that\u2019 costs you nothing; a wrong figure stated confidently costs you the interview."])

doc = SimpleDocTemplate(OUT, pagesize=A4,
                        leftMargin=20 * mm, rightMargin=20 * mm,
                        topMargin=18 * mm, bottomMargin=20 * mm,
                        title="IB Portfolio Explained", author="Jainam Mehta")
doc.build(story, onFirstPage=footer, onLaterPages=footer)
print("written", OUT, os.path.getsize(OUT), "bytes")
