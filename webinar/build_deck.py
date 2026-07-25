"""Build the Max Solutions M&A webinar deck.

Every number on a data slide traces to the research pass: Max Data registry
SQL (project "Database") or a cited public source. Sources render bottom-left
on each slide. Missing owner assets (photos, screenshots, QR codes) render as
loud amber PLACEHOLDER boxes.
"""
import sys

sys.path.insert(0, "/home/user/mckinsey-pptx")
sys.path.insert(0, "/home/user/mckinsey-pptx/webinar")

from mckinsey_pptx import PresentationBuilder
from max_theme import MAX_THEME
import max_slides  # noqa: F401  (registers custom templates)

SRC_MAXDATA = "Max Data registry analysis, July 2026 (1.99M Thai juristic persons, 9.8M financial statements)"

b = PresentationBuilder(theme=MAX_THEME)

# Engine templates default footnote to a "1. xx" placeholder; this deck never
# uses numbered footnotes, so default it off globally.
_engine_add = b.add


def _add(slide_type, **kw):
    kw.setdefault("footnote", None)
    return _engine_add(slide_type, **kw)


b.add = _add


def S(marker):
    """Section marker helper."""
    return {"section_marker": marker}


# =====================================================================
# OPEN
# =====================================================================

b.add("max_cover",
      title="Thailand's Hidden M&A Opportunity",
      subtitle="Where deals come from, how to judge them, and the fastest way in",
      event_line="Max Solutions live webinar",
      date="Sunday 26 July 2026 · 20:30 (ICT) · 90 minutes")

b.add("poll_slide", **S("Welcome"),
      title="Before we start",
      question="Which best describes you tonight?",
      options=["Buyer or investor looking for opportunities",
               "Owner thinking about an exit or succession",
               "Advisor, banker, or connector",
               "Exploring M&A for the first time"],
      instruction="Vote now in the poll panel")

b.add("agenda", **S("Agenda"),
      title="Tonight's map: six stops",
      items=["WHY NOW · Thailand's post-COVID reshuffle and the succession wave",
             "WHAT GOOD LOOKS LIKE · 5 green flags, 3 red flags",
             "WHERE DEALS COME FROM · community, marketplace, advisory, data",
             "HOW A DEAL RUNS · the buy-side path in 6 steps",
             "PROOF · two closed deals and the search-fund wave",
             "YOUR NEXT STEP · what to do tomorrow morning"])

b.add("speaker_slide", **S("Welcome"),
      title="Your guides tonight",
      speakers=[
          {"name": "[Max — full name]", "role": "Founder, Max Solutions",
           "bullets": ["Advises 150+ Thai SMEs per year on M&A",
                       "Built Max Data, our AI analytics platform",
                       "[Placeholder: 1-line personal credential]"],
           "photo_label": "Photo: Max"},
          {"name": "[Wipin — full name]", "role": "Head of M&A, Max Solutions",
           "bullets": ["Leads deal execution across 15 industries",
                       "[Placeholder: deals closed / years experience]",
                       "[Placeholder: 1-line personal credential]"],
           "photo_label": "Photo: Wipin"},
      ])

# =====================================================================
# SECTION 1 · WHY NOW
# =====================================================================

b.add("section_divider", section_number="01",
      section_title="Why Thailand, why now",
      subtitle="The market has re-sorted since COVID, and a generational handover is starting")

b.add("column_split_growth", **S("1 · Why now"),
      title="Thai corporate revenue has fully outgrown COVID: ฿68T and climbing",
      categories=[2016, 2017, 2018, 2019, 2020, 2021, 2022, 2023, 2024, 2025],
      values=[42.5, 44.5, 47.7, 48.1, 44.0, 48.6, 58.6, 60.2, 64.6, 68.0],
      split_index=4,
      growth_pct_first="+3.1%",
      growth_pct_second="+9.1%",
      description="Annual growth rates shown are CAGR, pre-COVID (2016-2020) vs recovery (2020-2025)",
      takeaway_header="What it means",
      data_label="Total revenue filed by Thai companies", data_unit="฿ trillion",
      takeaways=["COVID knocked ฿4T off filed revenue in 2020",
                 "Recovery ran at 9.1% a year, triple the pre-COVID pace",
                 "FY2025 sits 41% above the pre-COVID peak"],
      source=SRC_MAXDATA)

b.add("column_comparison", **S("1 · Why now"),
      title="The winners: five industries that took off after COVID",
      categories=["Accommodation\n& food", "Arts &\nentertainment",
                  "Wholesale\n& retail", "Other\nservices", "Mining &\nquarrying"],
      values=[140.9, 140.5, 90.2, 84.9, 74.3],
      focus_index=0,
      description="Industry sections ranked by filed-revenue growth over the recovery window",
      takeaway_header="What it means",
      data_label="Revenue growth FY2020 to FY2025", data_unit="%",
      takeaways=["Tourism reopening powers the top two",
                 "Wholesale and retail added the most baht: ฿33.5T in FY2025",
                 "Inside it, gold-shop trading alone grew ฿1.9T to ฿8.85T",
                 "Health care just missed the list at +72%"],
      source=SRC_MAXDATA)

b.add("column_comparison", **S("1 · Why now"),
      title="The laggards: where value is quietly leaking",
      categories=["Own-account\ninvestment", "Direct-sales\nretail (MLM)",
                  "TV programme\nproduction", "Rubber & plastics\nwholesale",
                  "Residential\ndevelopers"],
      values=[40, 36, 26, 19, 18],
      focus_index=4,
      description="Industries that genuinely contracted, shown as depth of revenue decline",
      takeaway_header="What it means",
      data_label="Revenue DECLINE FY2020 to FY2025", data_unit="%",
      takeaways=["No whole sector shrank in nominal terms",
                 "But these niches genuinely contracted",
                 "Real estate (+4%) and finance (+5%) declined in real terms",
                 "Weak segments are where motivated sellers live"],
      source=SRC_MAXDATA)

b.add("overview_areas", **S("1 · Why now"),
      title="The reshuffle in the four sectors this room plays in",
      subtitle="Operating counts, recovery growth, and margin swings, straight from the registry",
      areas=[
          {"name": "Logistics",
           "bullets": ["44,190 operating companies",
                       "Revenue +66% since 2020",
                       "Margin swing: -14% to +7.5%",
                       "9,399 mid-market targets ฿10M-1B"]},
          {"name": "Manufacturing",
           "bullets": ["109,063 operating companies",
                       "฿21.1T revenue, +30% since 2020",
                       "Steady 4.3% net margin",
                       "Thailand's largest sector"]},
          {"name": "F&B & hospitality",
           "bullets": ["48,564 operating companies",
                       "Revenue +141% off the 2020 trough",
                       "Margin swing: -22% to +7.4%",
                       "The steepest recovery arc"]},
          {"name": "Tech & information",
           "bullets": ["29,027 operating companies",
                       "฿950B revenue, +18% since 2020",
                       "Margin up from 5.9% to 7.5%",
                       "Growth in profit, not just sales"]},
      ],
      call_out="In logistics, fewer than 200 of 44,190 companies exceed ฿1B revenue",
      source=SRC_MAXDATA)

b.add("two_column_compare", **S("1 · Why now"),
      title="Two roads to growth, and only one of them is fast",
      left_label="ORGANIC: build it",
      right_label="INORGANIC: buy it",
      left_items=["Grow sales customer by customer",
                  "Hire and train your own team",
                  "Years to reach scale in a new market",
                  "Low risk per step, slow compounding"],
      right_items=["Acquire revenue, staff, and licences on day one",
                   "A clinic chain grows by buying clinics",
                   "Months, not years, to enter a market",
                   "The skill is picking the right target"],
      right_color="blue",
      source="Max Solutions M&A advisory practice")

b.add("two_column_compare", **S("1 · Why now"),
      title="M&A used to be a big-company game. Not anymore",
      left_label="M&A THEN",
      right_label="M&A NOW",
      left_items=["Reserved for conglomerates and funds",
                  "Deal sizes in the billions",
                  "Returns measured over 10+ years",
                  "Sourced through investment banks"],
      right_items=["Individuals and mid-size corporates buy companies",
                   "SME deals: ฿10M to ฿1B revenue targets",
                   "Payback often 3 to 8 years in our deal experience",
                   "Japan grew SME succession M&A 16x in 8 years"],
      source="Max Solutions deal experience; Japan figure: METI via Kobe University MAREC, 2014-2022")

b.add("column_comparison", **S("1 · Why now"),
      title="Thai M&A just woke up: the strongest quarter in years",
      categories=["Q1 2025", "Q2 2025", "Q3 2025", "Q4 2025", "Q1 2026"],
      values=[1.0, 3.4, 2.5, 1.3, 7.9],
      focus_index=4,
      description="Big-cap headlines mask the SME layer underneath, which is tonight's focus",
      takeaway_header="What it means",
      data_label="Announced deal value", data_unit="USD billions",
      takeaways=["Q1 2026: USD 7.9B across 57 deals, 6x the prior quarter",
                 "One deal, the True Corporation 25% stake, was USD 3.9B",
                 "Inbound buyers took 39% of deal value in Q4 2025",
                 "ASEAN-6 did 3,200+ deals worth USD 130B+ in 2025"],
      source="KPMG Thailand M&A Trends, quarterly reports Q1 2025 to Q1 2026; ASEAN figure: Lyndon Advisory 2026")

b.add("stat_hero", **S("1 · Why now"),
      title="The succession wave is the supply side of this market",
      stat="81%",
      stat_label="of Thailand's 20-year-plus companies have no next-generation director on the board",
      context="111,304 operating companies are more than 20 years old with every "
              "board seat held by the pre-1984 generation. Across all 993k operating "
              "companies, 2 of every 3 directorships sit with that generation. "
              "Companies like these need successors. Increasingly, that means buyers.",
      source="Max Data registry analysis of 4.2M directorships, birth cohorts from national-ID structure, July 2026")

b.add("kpi_dashboard", **S("1 · Why now"),
      title="The demographic clock behind it is public record",
      kpis=[
          {"label": "of Thais are aged 60+, and rising", "value": "20%",
           "delta": "super-aged society by 2033", "delta_dir": "up"},
          {"label": "of Thai businesses are family-owned", "value": "~80%",
           "delta": "฿30T combined value", "delta_dir": "flat"},
          {"label": "of family firms survive to generation 2", "value": "1 in 3",
           "delta": "global benchmark", "delta_dir": "down"},
          {"label": "SME succession deals a year in Japan already", "value": "3,000+",
           "delta": "the preview of Thailand's next decade", "delta_dir": "up"},
      ],
      columns=4,
      source="NESDC via Nation Thailand 2024; Grant Thornton Thailand; METI Japan via ORIX 2025")

b.add("kpi_dashboard", **S("1 · Why now"),
      title="We saw this wave coming three years ago",
      kpis=[
          {"label": "SMEs advised each year", "value": "150+",
           "delta": "across 15 industries", "delta_dir": "flat"},
          {"label": "DealFlow community", "value": "80,000+",
           "delta": "members and growing", "delta_dir": "up"},
          {"label": "Live deals on our platform", "value": "100+",
           "delta": "screened listings", "delta_dir": "up"},
      ],
      columns=3,
      source="Max Solutions, July 2026")

# =====================================================================
# SECTION 2 · WHAT GOOD LOOKS LIKE
# =====================================================================

b.add("section_divider", section_number="02",
      section_title="What a good target looks like",
      subtitle="The lens our deal team applies before we ever talk price")

b.add("poll_slide", **S("2 · The lens"),
      title="Quick pulse check",
      question="You are shown a company for sale. What do you check first?",
      options=["The financial statements",
               "The owner and why they are selling",
               "The customers and recurring revenue",
               "The competition around it"],
      instruction="Vote now. We'll show you our order in a minute")

b.add("five_key_areas", **S("2 · The lens"),
      title="The Green 5: what our analysts look for in every target",
      subtitle="The five signals our deal team scores before we ever talk price",
      areas=[
          {"name": "Synergy",
           "description": "Does it multiply what you already own: customers, licences, capacity, distribution"},
          {"name": "Potential",
           "description": "Is the market growing, and is there obvious headroom the current owner never used"},
          {"name": "Management",
           "description": "Can the business run without the owner, or does the value walk out with them"},
          {"name": "Revenue quality",
           "description": "Recurring customers and repeat contracts beat one-off project income every time"},
          {"name": "Competitive position",
           "description": "A defensible niche with few real rivals, not a price war you are buying into"},
      ],
      source="Max Solutions target evaluation framework")

b.add("three_trends_icons", **S("2 · The lens"),
      title="The Red 3: where deals die in Thailand",
      subtitle="Our deal team's kill criteria, in the order they usually surface",
      trends=[
          {"label": "Business risk", "icon": "📉",
           "bullets": ["Shrinking market or fad demand",
                       "Customer concentration in 1-2 accounts",
                       "Model only works with the founder's relationships"]},
          {"label": "Legal risk", "icon": "⚖",
           "bullets": ["Missing licences and permits",
                       "Land or building not properly owned",
                       "A deal-breaker: you cannot legally buy what isn't legal"]},
          {"label": "Financial risk", "icon": "📒",
           "bullets": ["Two sets of books is common in Thai SMEs",
                       "Hidden liabilities and undeclared debt",
                       "Verify cash, not stories"]},
      ],
      source="Max Solutions deal experience, 150+ SME mandates per year")

b.add("two_column_compare", **S("2 · The lens"),
      title="The two-books problem, and why diligence pays for itself",
      left_label="THE BOOKS THEY SHOW",
      right_label="THE BOOKS THAT MATTER",
      left_items=["Prepared for the tax office",
                  "Understated revenue and profit",
                  "Looks cheap, and is unbankable",
                  "Cannot support your valuation"],
      right_items=["Actual cash flows and contracts",
                   "Real margins, verified in diligence",
                   "Supports financing and a fair price",
                   "This is what you are actually buying"],
      left_color="gray", right_color="navy",
      source="Max Solutions deal experience")

b.add("comparison_table", **S("2 · The lens"),
      title="Take this home: the 60-second target scorecard",
      subtitle="Score any company you are shown. Green 5 up top, Red 3 below",
      options=["Weak", "Average", "Strong"],
      criteria=[
          {"name": "Synergy with what you own", "scores": [1, 2, 4]},
          {"name": "Untapped potential", "scores": [1, 2, 4]},
          {"name": "Runs without the owner", "scores": [0, 2, 4]},
          {"name": "Recurring revenue", "scores": [1, 3, 4]},
          {"name": "Defensible niche", "scores": [1, 2, 4]},
          {"name": "Business risk cleared", "scores": [0, 2, 4]},
          {"name": "Legal risk cleared", "scores": [0, 2, 4]},
          {"name": "Financials verified", "scores": [0, 2, 4]},
      ],
      source="Max Solutions target evaluation framework")

# =====================================================================
# SECTION 3 · WHERE DEALS COME FROM
# =====================================================================

b.add("section_divider", section_number="03",
      section_title="Where deals actually come from",
      subtitle="From two million registered companies to the handful worth your time")

b.add("funnel", **S("3 · The access"),
      title="Finding the needle: Thailand's company universe",
      stages=[
          {"name": "Registered juristic persons", "value": "1,989,709",
           "description": "Every company ever registered in Thailand"},
          {"name": "Operating today", "value": "993,373",
           "description": "Half the registry is already defunct or dissolved"},
          {"name": "Filing real revenue", "value": "648,304",
           "description": "Companies with revenue above zero in FY2025"},
          {"name": "SME sweet spot ฿100M-1B", "value": "37,045",
           "description": "Only 1 in 18 revenue-reporting companies"},
          {"name": "Your shortlist", "value": "10-30",
           "description": "What a focused thesis and good filters produce"},
      ],
      source=SRC_MAXDATA)

b.add("three_trends_icons", **S("3 · The access"),
      title="How everyone finds deals today, and where each channel stops",
      subtitle="Each channel works. Each has a ceiling you will hit",
      trends=[
          {"label": "Word of mouth", "icon": "🤝",
           "bullets": ["Friends, suppliers, golf partners",
                       "High trust, tiny reach",
                       "You see 2-3 deals a year, if lucky"]},
          {"label": "Online listings", "icon": "🌐",
           "bullets": ["Marketplaces and brokers' sites",
                       "Wide reach, uneven quality",
                       "The good ones go fast"]},
          {"label": "Intermediaries", "icon": "💼",
           "bullets": ["Advisors and boutique firms",
                       "Screened deals, real process",
                       "Works best with a clear mandate"]},
      ],
      source="Max Solutions market observation")

b.add("access_ladder", **S("3 · The access"),
      title="Our answer: three doors, one ecosystem",
      steps=[
          {"name": "Community", "stat": "80,000+ members",
           "bullets": ["DealFlow Facebook community",
                       "Owners, buyers, and advisors in one room",
                       "Off-market chatter surfaces here first"]},
          {"name": "Marketplace", "stat": "100+ live deals",
           "bullets": ["DealFlow Market listings",
                       "Screened sellers across 15 industries",
                       "Thailand's leading M&A marketplace"]},
          {"name": "Advisory", "stat": "150 SMEs / year",
           "bullets": ["Full-mandate M&A advisory",
                       "Our deal team runs the process end to end",
                       "For buyers and sellers who are serious"]},
      ],
      source="Max Solutions, July 2026")

b.add("screenshot_slide", **S("3 · The access"),
      title="Inside the community: where off-market deals surface",
      placeholder_label="Screenshot: DealFlow Facebook community",
      placeholder_note="Show member count + 2-3 recent deal posts",
      bullets=["80,000+ members: owners, buyers, advisors",
               "Sellers often test the water here before any listing",
               "Free to join, and the fastest way to see deal flow"],
      stats=[{"value": "80,000+", "label": "community members"}],
      source="DealFlow community, July 2026")

b.add("screenshot_slide", **S("3 · The access"),
      title="Inside DealFlow Market: screened deals, real numbers",
      placeholder_label="Screenshot: DealFlow Market listings page",
      placeholder_note="Show listing grid with sectors + asking prices visible",
      bullets=["100+ live listings across 15 industries",
               "Every listing screened by our team before it goes up",
               "Filter by sector, size, and location"],
      stats=[{"value": "100+", "label": "live deals"},
             {"value": "15", "label": "industries"}],
      source="DealFlow Market, July 2026")

b.add("process_flow_horizontal", **S("3 · The access"),
      title="Max Data: our AI analytics platform, built for every stage",
      steps=[
          {"name": "Research",
           "description": "Market and industry trends from 1.99M companies"},
          {"name": "Source",
           "description": "Filter to targets that fit your thesis"},
          {"name": "Validate",
           "description": "10 years of financials on any company"},
          {"name": "Diligence",
           "description": "Directors, licences, branches, procurement wins"},
          {"name": "Reach",
           "description": "Contact the actual decision maker"},
      ],
      source="Max Data platform, July 2026")

b.add("screenshot_slide", **S("3 · The access"),
      title="Max Data in 90 seconds",
      placeholder_label="Screenshot or live demo: Max Data screener",
      placeholder_note="Suggested demo: Logistics, revenue ฿10M-1B → 9,399 live targets",
      bullets=["Type a thesis, get a shortlist in seconds",
               "Every company: financials, directors, growth, red flags",
               "Built by our deal team for real transactions, not just research"],
      stats=[{"value": "1.99M", "label": "companies tracked"},
             {"value": "9.8M", "label": "financial statements"},
             {"value": "4.2M", "label": "directorships mapped"}],
      source="Max Data platform, July 2026")

# =====================================================================
# SECTION 4 · HOW A DEAL RUNS
# =====================================================================

b.add("section_divider", section_number="04",
      section_title="How a deal actually runs",
      subtitle="The buy-side path in six steps, and the seller's mirror image")

b.add("process_flow_horizontal", **S("4 · The path"),
      title="The buy-side path: six steps from thesis to keys",
      steps=[
          {"name": "Thesis",
           "description": "What you want to own and why. One page, honest"},
          {"name": "Shortlist",
           "description": "Filter 993k operating companies down to 10-30 fits"},
          {"name": "Approach",
           "description": "Reach the owner, build trust, sign an NDA"},
          {"name": "LOI",
           "description": "Letter of intent, sometimes a deposit. Shows you are real"},
          {"name": "Diligence",
           "description": "Clear the Red 3: business, legal, financial"},
          {"name": "Close & handover",
           "description": "Sign, pay, transition the team and customers"},
      ],
      source="Max Solutions buy-side playbook")

b.add("two_column_compare", **S("4 · The path"),
      title="Selling? Same road, driven in reverse",
      left_label="BUYER'S PATH",
      right_label="SELLER'S MIRROR",
      left_items=["Write a thesis",
                  "Shortlist and approach targets",
                  "Sign LOI, run diligence",
                  "Close and take over"],
      right_items=["Prepare your numbers and story",
                   "Go to market quietly, collect offers",
                   "Survive diligence: buyers will check the Green 5 and Red 3 in you",
                   "Close and hand over clean"],
      show_arrow=False,
      source="Max Solutions sell-side playbook. Full process guide on our website")

# =====================================================================
# SECTION 5 · PROOF
# =====================================================================

b.add("section_divider", section_number="05",
      section_title="Proof it works",
      subtitle="Two deals our team closed this year, and the global wave behind them")

b.add("poll_slide", **S("5 · Proof"),
      title="One more pulse check",
      question="Which situation is closer to yours?",
      options=["I want to buy my first business",
               "My company should be acquiring",
               "I own a business that needs a successor",
               "I connect people and want the referral fee"],
      instruction="Vote now. The cases coming up cover all four")

b.add("stat_hero", **S("5 · Proof"),
      title="The search-fund wave: individuals now buy companies, profitably",
      stat="35.1%",
      stat_label="aggregate IRR across search funds tracked by Stanford GSB",
      context="Aggregate returns of 4.5x invested capital, a record 94 funds launched "
              "in 2023, and the model has gone global: 320 international funds by "
              "end-2023, buying real SMEs at a median price of USD 11.7M. The same "
              "playbook, applied to Thailand's succession wave, is what tonight is about.",
      source="Stanford GSB Search Fund Study 2024 (Case E-870); IESE International Search Funds 2024")

b.add("case_slide", **S("5 · Proof"),
      title="Case one: the fire-safety distributor",
      case_name="Project FireGuard · closed Q1 2026",
      sector_chip="FIRE SAFETY / TRADING",
      situation=["Husband-and-wife owners near retirement, no successor",
                 "฿100M revenue, ฿30M EBITDA, decades of relationships",
                 "One of only 112 fire-safety companies tracked in Max Data",
                 "Zero digitization: paper everywhere, no ERP"],
      outcome=["New owner modernized systems within months",
               "Employees stayed, morale improved",
               "Owners exited proud, business on an IPO-track plan"],
      kpis=[{"value": "฿100M", "label": "Revenue at deal"},
            {"value": "฿30M", "label": "EBITDA at deal"},
            {"value": "months", "label": "To visible turnaround"}],
      photo_label=None,
      source="Max Solutions deal team, 2026. Figures approximate to protect the parties")

b.add("case_slide", **S("5 · Proof"),
      title="Case two: the pizza-oven supplier",
      case_name="Project Pizza Oven · closed Q1 2026",
      sector_chip="F&B EQUIPMENT / SUPPLY",
      situation=["Italian owner couple returning home after years in Thailand",
                 "Supplier to major Thai restaurant chains",
                 "One of 94 kitchen-equipment suppliers tracked in Max Data",
                 "Top-3 Google ranking in its niche, loyal recurring customers"],
      outcome=["Foreign buyer acquired end to end through our process",
               "Smooth handover, customers retained",
               "Now on its way to doubling turnover"],
      kpis=[{"value": "End-to-end", "label": "Run by our deal team"},
            {"value": "Recurring", "label": "Customer base"},
            {"value": "2x", "label": "Turnover trajectory"}],
      source="Max Solutions deal team, 2026. Figures approximate to protect the parties")

b.add("three_trends_numbered", **S("5 · Proof"),
      title="Why small deals turn around so fast",
      subtitle="The pattern behind both cases, and most of our closed deals",
      trends=[
          {"label": "Succession creates honest sellers",
           "bullets": ["Owners sell because of age, not because the business is broken",
                       "Price reflects the exit need, not a bidding war"]},
          {"label": "Inefficiency is the upside",
           "bullets": ["No ERP, no marketing, no pricing discipline",
                       "Basic modernization moves margins in months"]},
          {"label": "Capable buyers compound it",
           "bullets": ["Corporate resources or sharp operators unlock value fast",
                       "SME payback in 3-8 years vs 10+ on mega-deals"]},
      ],
      source="Max Solutions deal experience across 150+ SME mandates per year")

b.add("two_column_compare", **S("5 · Proof"),
      title="Sourcing then vs now: what the data layer changes",
      left_label="THE OLD WAY",
      right_label="WITH MAX DATA",
      left_items=["Call owners one by one, hope someone wants to sell",
                  "Guess at financials until diligence",
                  "Months to build a target list",
                  "Your edge = who you happen to know"],
      right_items=["Screen 993k operating companies against your thesis",
                   "10 years of financials before the first call",
                   "A shortlist in an afternoon",
                   "Your edge = seeing what others cannot"],
      right_color="blue",
      source="Max Data platform, July 2026")

# =====================================================================
# SECTION 6 · ACT
# =====================================================================

b.add("section_divider", section_number="06",
      section_title="Your next step",
      subtitle="Three frameworks to keep, one action to take tonight")

b.add("executive_summary_takeaways", **S("6 · Next step"),
      title="What you now have",
      sections=[
          {"takeaway": "A reason to move now",
           "bullets": ["Post-COVID winners and laggards are visible in the data",
                       "The succession wave is bringing good companies to market"]},
          {"takeaway": "A lens to judge any deal",
           "bullets": ["The Green 5 and the Red 3",
                       "The 60-second target scorecard"]},
          {"takeaway": "A way in",
           "bullets": ["Community → Marketplace → Advisory, plus the Max Data layer",
                       "The 6-step path from thesis to keys"]},
      ],
      final_conclusion="The buyers who win the succession wave are the ones who start looking before everyone else.")

b.add("cta_slide", **S("6 · Next step"),
      title="Do one of these before you log off",
      paths=[
          {"who": "BUYERS & INVESTORS", "action": "Book a free opportunity scan",
           "detail": "We map live targets and market data against your thesis in one session."},
          {"who": "OWNERS", "action": "Get a confidential valuation talk",
           "detail": "Know what your business is worth and what buyers would flag, no obligation."},
          {"who": "EVERYONE", "action": "Join the DealFlow community",
           "detail": "80,000+ members. And if you refer a buyer or seller, our referral program pays you."},
      ],
      bottom_actions=["Type 1 in the chat and our team will contact you tomorrow",
                      "Or scan the QR / add our LINE official account now"],
      qr_label="QR: booking page")

b.add("dark_navy_summary",
      body="Q&A: ask us anything. The QR stays on screen, and we stay until your questions run out.",
      eyebrow="Thailand's Hidden M&A Opportunity",
      corner_text="Max Solutions")

b.add("thank_you",
      lines=["Max Solutions · DealFlow Market · Max Data",
             "Thank you for spending your Sunday evening with us"],
      contact_placeholder="Contact block: LINE ID, email, phone, website")

# =====================================================================
# APPENDIX
# =====================================================================

b.add("section_divider", section_number="A",
      section_title="Appendix",
      subtitle="For the replay: process detail, glossary, and methodology")

b.add("process_flow_horizontal", **S("Appendix"),
      title="Sell-side process in full",
      steps=[
          {"name": "Prepare",
           "description": "Clean numbers, story, and a defensible asking range"},
          {"name": "Go to market",
           "description": "Quiet outreach to screened buyers, teaser first"},
          {"name": "Offers",
           "description": "Collect LOIs, compare price and certainty"},
          {"name": "Diligence",
           "description": "Buyer verifies business, legal, financial"},
          {"name": "Close",
           "description": "Sign, settle, announce on your terms"},
          {"name": "Handover",
           "description": "Transition plan for team, customers, suppliers"},
      ],
      source="Max Solutions sell-side playbook")

b.add("five_key_areas", **S("Appendix"),
      title="Glossary: five terms you heard tonight",
      subtitle="Plain-language definitions for the replay",
      areas=[
          {"name": "LOI",
           "description": "Letter of intent: a non-binding offer that states price range and terms before diligence"},
          {"name": "EBITDA",
           "description": "Earnings before interest, tax, depreciation, amortization: the profit buyers actually price on"},
          {"name": "Due diligence",
           "description": "The verification phase: business, legal, and financial checks before closing"},
          {"name": "Search fund",
           "description": "An investor-backed vehicle where one operator searches for, buys, and runs a single SME"},
          {"name": "Succession deal",
           "description": "A sale driven by owner retirement with no family successor, Thailand's fastest-growing deal type"},
      ],
      source="Max Solutions")

b.add("executive_summary_paragraph", **S("Appendix"),
      title="Methodology and sources",
      paragraphs=[
          "Registry statistics in this deck come from Max Data, our AI-driven M&A "
          "analytics platform covering 1.99 million Thai juristic persons, 9.8 million "
          "financial statements, and 4.2 million directorships, snapshot July 2026. "
          "Fiscal years follow Thai filing years. FY2025 filings were about 98% "
          "complete at analysis time.",
          "Industry growth compares revenue filed in FY2020 against FY2025 at TSIC "
          "section and industry level. No whole section declined in nominal terms "
          "over that window. The laggards shown are genuine contracting industries "
          "at TSIC level.",
          "Succession figures use birth cohorts encoded in Thai national-ID "
          "structure: IDs issued before the 1984 system change identify directors "
          "born before 1984, a hard age floor of 42+ today. The registry holds no "
          "birthdates, so we state cohorts, not ages. No published statistic on "
          "Thai owner age exists; this analysis is, to our knowledge, the first.",
          "External statistics are cited on their slides: KPMG Thailand quarterly "
          "M&A reports, NESDC demographic data, Grant Thornton Thailand, METI "
          "Japan, and Stanford GSB search fund research. Case figures are rounded "
          "and lightly disguised to protect client confidentiality.",
      ],
      source="Max Solutions research team, July 2026")

# =====================================================================
# SPEAKER NOTES — timing budgets + delivery cues from the planning session
# =====================================================================

NOTES = {
    1: "Doors open 20:25, greet people by name in chat. Start 20:32 sharp. [T+0:00]",
    2: "Launch poll 1 immediately, read the options aloud. Tease: 'we'll tailor the examples to tonight's mix.' About 2 minutes. [T+0:02]",
    3: "One breath per stop. The promise: 'by the end you'll know where deals come from, how to judge one, and your first step.' [T+0:05]",
    4: "30 seconds each. Max: keep the founder story to ONE line here, the full beat comes at slide 15. [T+0:06]",
    5: "Section 1 · WHY NOW. Budget 12 minutes. [T+0:07]",
    6: "Big line: 'COVID knocked four trillion baht off the market. The recovery ran at triple the old pace.'",
    7: "Pause on the gold factoid. Chat prompt: 'type your industry in the chat.'",
    8: "Frame positively: weak segments are where motivated sellers live. Don't dwell, next slide is the room's own sectors.",
    9: "This room's four sectors. Logistics people: 9,399 mid-market targets is YOUR number. Manufacturing: biggest sector, steady margins.",
    10: "90 seconds max. The room knows this distinction, it's a bridge slide.",
    11: "Key claim: SME deals pay back in 3-8 years vs 10+ for mega-deals. Japan's 16x growth is the evidence the wave is real.",
    12: "Credibility beat: this is KPMG data published last month. Note: headlines are big-cap, our layer is the SME market below them.",
    13: "SLOW DOWN. This is the screenshot slide. Read the stat twice. 'No one else in Thailand can compute this number.'",
    14: "Public-record confirmation of our registry finding. The Japan tile is the preview of Thailand's next decade.",
    15: "Founding story, 45 seconds max: 'we saw the wave three years ago.' [T+0:19]",
    16: "Section 2 · WHAT GOOD LOOKS LIKE. Budget 15 minutes. [T+0:19]",
    17: "Launch poll 2. Callback comes at the Red 3 slide: most people vote financials, and financials is where deals die.",
    18: "One concrete example per flag. Ask chat: 'which of these do you weight most?'",
    19: "One short war story per risk if time allows. Two-books gets its own slide next.",
    20: "Sensitive topic, phrase carefully: 'a common practice, and fixable in diligence' — never accusatory.",
    21: "Tell viewers to screenshot this one. It's also in the replay materials.",
    22: "Section 3 · WHERE DEALS COME FROM. Budget 12 minutes. [T+0:34]",
    23: "Walk the cascade slowly. 'Only 1 in 18 companies sits in the sweet spot' lands well.",
    24: "Two minutes max, this is setup for the ladder.",
    25: "Our ecosystem. Not a pitch: the first two doors are free to open tonight.",
    26: "Show the real Facebook group live if the connection allows, otherwise the screenshot.",
    27: "Pick one live listing that matches the poll-1 audience mix and talk through it for 30 seconds.",
    28: "Position Max Data as the research layer under every stage, not another listing site.",
    29: "DEMO MOMENT. Recorded fallback ready. Filter: Logistics, ฿10M-1B revenue → 9,399 companies on screen. [T+0:44]",
    30: "Section 4 · HOW A DEAL RUNS. Budget 8 minutes. [T+0:46]",
    31: "Walk the six steps in about 4 minutes. Add the LOI + deposit nuance from real deals.",
    32: "Sellers: 60 seconds. Full sell-side detail is in the appendix and on the website.",
    33: "Section 5 · PROOF. Budget 12 minutes. [T+0:54]",
    34: "Launch poll 3. Use the result to decide which case to emphasize.",
    35: "Let 35.1% breathe. Then: 'this playbook is arriving in Thailand.'",
    36: "Tell it as a story: the couple, no successor, the sharp buyer, the New Year visit where everyone was happy.",
    37: "Contrast case: foreign buyer, end-to-end process, now doubling turnover.",
    38: "Generalize the pattern: succession + inefficiency + capable buyer.",
    39: "Callback to the demo: 'your edge is seeing what others cannot.' [T+1:04]",
    40: "Section 6 · YOUR NEXT STEP. Budget 6 minutes. [T+1:06]",
    41: "Three take-homes. Second screenshot moment.",
    42: "THE ask. Type 1 / scan QR / add LINE. Say it twice. Keep the QR visible through Q&A. [T+1:08]",
    43: "About 20 minutes. Seed questions ready: 'what multiples do Thai SMEs sell for?', 'can foreigners buy?', 'how long does a deal take?'",
    44: "Mention the replay and community link land in tomorrow's follow-up email.",
    45: "Appendix: replay material. Skip live unless Q&A pulls them up.",
}

for _i, _slide in enumerate(b.prs.slides, 1):
    note = NOTES.get(_i)
    if note:
        _slide.notes_slide.notes_text_frame.text = note

out = "/home/user/mckinsey-pptx/webinar/output/ma-webinar-deck.pptx"
b.save(out)
print(f"saved {out} ({len(b.prs.slides._sldIdLst)} slides)")
