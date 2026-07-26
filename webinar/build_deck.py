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
import editorial_slides  # noqa: F401  (registers editorial templates)

SRC_MAXDATA = "Max Data registry analysis, July 2026 (1.99M Thai juristic persons, 9.8M financial statements)"

import os
from pptx.util import Inches

ASSETS = "/home/user/mckinsey-pptx/webinar/assets"
GEN = f"{ASSETS}/gen"


def _ic(name):
    """Icon path (transparent PNG); slides skip icons that don't exist yet."""
    return f"{GEN}/icons/{name}.png"


def _pic(slide, path, x, y, w, h):
    """Post-hoc cover-cropped vignette; silently skipped if asset missing."""
    if os.path.exists(path):
        crop = editorial_slides.cover_crop(path, w, h)
        slide.shapes.add_picture(crop, Inches(x), Inches(y),
                                 width=Inches(w), height=Inches(h))

b = PresentationBuilder(theme=MAX_THEME)

# Engine templates default footnote to a "1. xx" placeholder; this deck never
# uses numbered footnotes, so default it off globally. Every content slide in
# sections 1-6 also gets the bottom progress tracker so a viewer who drifted
# off can re-orient instantly.
_engine_add = b.add

SECTION_OF = {
    "1 · Why now": 0, "2 · The access": 1, "3 · The lens": 2,
    "4 · The path": 3, "5 · Proof": 4, "6 · Next step": 5,
}


def _add(slide_type, **kw):
    kw.setdefault("footnote", None)
    slide = _engine_add(slide_type, **kw)
    idx = SECTION_OF.get(kw.get("section_marker"))
    if idx is not None:
        max_slides.add_progress_tracker(slide, idx)
    if slide_type == "section_divider":
        editorial_slides.divider_strip(slide)
    return slide


b.add = _add


def S(marker):
    """Section marker helper."""
    return {"section_marker": marker}


# =====================================================================
# OPEN
# =====================================================================

b.add("max_cover",
      title="Thailand's Hidden M&A Opportunity",
      subtitle="Mergers and acquisitions, in plain English: where deals come from, how to judge them, and the fastest way in",
      event_line="Max Solutions live webinar",
      date="Sunday 26 July 2026 · 20:30 (ICT) · 90 minutes")

b.add("poll_slide", **S("Welcome"),
      title="Before we start",
      question="Which best describes you tonight?",
      options=["Buyer or investor looking for opportunities",
               "Owner thinking about an exit or succession",
               "Advisor, banker, or connector",
               "Exploring M&A for the first time"],
      instruction="Vote now in the poll panel",
      ornament_path=f"{GEN}/backdrops/poll-corner-alpha.png")

b.add("route_map", **S("Agenda"),
      title="Tonight's map: six stops",
      subtitle="Six stops, 90 minutes. The bar at the bottom of every slide shows where we are.",
      stops=[
          {"keyword": "WHY NOW", "descriptor": "The reshuffle and the succession wave"},
          {"keyword": "THE ACCESS", "descriptor": "Community, marketplace, advisory, data"},
          {"keyword": "THE LENS", "descriptor": "5 green flags, 3 red flags"},
          {"keyword": "THE PATH", "descriptor": "Thesis to keys in 6 steps"},
          {"keyword": "PROOF", "descriptor": "Closed deals and the search-fund wave"},
          {"keyword": "NEXT STEP", "descriptor": "What to do tomorrow morning"},
      ])

b.add("audience_map", **S("Welcome"),
      title="Whatever brought you here, tonight has a lane for you",
      takeaway="Stay for all six stops. These are the ones that pay your seat back first.",
      rows=[
          {"icon": _ic("target"), "who": "BD & corporate development",
           "detail": "logistics · manufacturing · F&B",
           "value": "A build-vs-buy case for your growth plan, and Max Data's "
                    "registry intelligence and director network to power your "
                    "pipeline.",
           "stops": [1, 2, 4], "key_stop": 2},
          {"icon": _ic("book"), "who": "IR & listed-company leaders",
           "detail": "market narrative for boards and investors",
           "value": "A data-backed read on the Thai market, and the M&A "
                    "growth story your investors are starting to ask about.",
           "stops": [1, 2, 5], "key_stop": 1},
          {"icon": _ic("storefront"), "who": "Owners & founders",
           "detail": "exit, succession, or just curious",
           "value": "What buyers check first, what your business is worth, "
                    "and how a quiet, well-run exit actually happens.",
           "stops": [3, 4, 6], "key_stop": 3},
          {"icon": _ic("handshake"), "who": "Investors, advisors & explorers",
           "detail": "first deal or fiftieth",
           "value": "Where real deal flow surfaces, proof the model pays, "
                    "and how the referral economy around it works.",
           "stops": [2, 5, 6], "key_stop": 2},
      ],
      offer="Leading a team with one industry and one goal? We run tailored "
            "private sessions for leadership teams. Ask us after Q&A, or book "
            "with the QR at the end.",
      source="Registered attendee mix, July 2026")

b.add("speaker_panels", **S("Welcome"),
      title="Your guides tonight",
      speakers=[
          {"name": "Chatpong Lappitakpong", "role": "MD, Max Solutions",
           "bullets": ["Advises 150+ Thai SMEs per year on M&A",
                       "Founded Max Solutions, a leading SME M&A advisory firm",
                       "Founded DealFlow Market, a leading M&A marketplace with 80k+ members",
                       "Founded Max Data, Thailand's research and opportunity sourcing platform"],
           "photo_path": f"{ASSETS}/people/max-headshot.png"},
          {"name": "Vipin Chugh", "role": "Head of M&A, Max Solutions",
           "bullets": ["16+ years of M&A advisory experience",
                       "Leads deal execution across 15 industries",
                       "Extensive experience managing SME deals, end to end",
                       "Closes 3+ deals a year on average"],
           "photo_path": f"{ASSETS}/people/vipin-headshot-pad.jpg"},
      ])

# =====================================================================
# SECTION 1 · WHY NOW
# =====================================================================

b.add("art_divider", art_path=f"{GEN}/dividers/div-1.png",
      section_number="01",
      section_title="Why Thailand, why now",
      subtitle="The market has re-sorted since COVID, and a generational handover is starting")

b.add("column_split_growth", **S("1 · Why now"),
      title="Thai corporate revenue has fully outgrown COVID: ฿68T and climbing",
      categories=[2016, 2017, 2018, 2019, 2020, 2021, 2022, 2023, 2024, 2025],
      values=[42.5, 44.5, 47.7, 48.1, 44.0, 48.6, 58.6, 60.2, 64.6, 68.0],
      split_index=4,
      growth_pct_first="+4.2%",
      growth_pct_second="+9.1%",
      description="Trough to today the recovery ran 9.1% a year. Peak to peak, FY2019 to FY2025, it still compounded 5.9%",
      takeaway_header=None,
      axis=False,
      bar_colors=["steel", "steel", "steel", "steel", "pale",
                  "navy", "navy", "navy", "navy", "cyan"],
      data_label="Total revenue filed by Thai companies", data_unit="฿ trillion",
      source=SRC_MAXDATA)

b.add("hbar_ranked", **S("1 · Why now"),
      title="The winners: five industries that took off after COVID",
      takeaway="Measured from the 2020 trough, so tourism-linked sectors flatter first.",
      unit_note="Revenue growth FY2020 to FY2025, %",
      items=[
          {"label": "Hotels & hospitality", "value": 141, "display": "+141%"},
          {"label": "Arts & entertainment", "value": 140, "display": "+140%"},
          {"label": "Wholesale & retail", "value": 90, "display": "+90%"},
          {"label": "Repair & personal services", "value": 85, "display": "+85%"},
          {"label": "Mining & quarrying", "value": 74, "display": "+74%"},
      ],
      annotation={"text": "Gold-shop trading alone grew ฿1.9T → ฿8.85T",
                  "row": 2, "x": 6.55, "w": 2.95},
      source=SRC_MAXDATA)

b.add("hbar_ranked", **S("1 · Why now"),
      title="The laggards: where value is quietly leaking",
      takeaway="Every one of these is a hunting ground. Falling revenue makes owners answer the phone.",
      direction="left",
      items=[
{"label": "Residential developers", "value": 18, "display": "-18%",
           "note": "Unsold inventory, high household debt"},
{"label": "Rubber & plastics wholesale", "value": 19, "display": "-19%",
           "note": "Squeezed between factories and end buyers"},
{"label": "TV programme production", "value": 26, "display": "-26%",
           "note": "Ad budgets moved to streaming and social"},
{"label": "Direct-sales retail (MLM)", "value": 36, "display": "-36%",
           "note": "Door-to-door and network selling, displaced by e-commerce"},
{"label": "Investment holding companies", "value": 40, "display": "-40%",
           "note": "Firms whose business is holding assets, not trading"},
                                                        ],
      unit_note="Revenue change FY2020 to FY2025, %. No whole sector shrank in nominal terms, but these niches did",
      source=SRC_MAXDATA)

b.add("sector_matrix", **S("1 · Why now"),
      title="The reshuffle in the four sectors this room plays in",
      subtitle="Operating counts, recovery growth, and margin swings, straight from the registry",
      rows=[
          {"name": "Logistics", "icon": _ic("truck"), "count": "44,190",
           "note": "9,399 mid-market targets ฿10M-1B",
           "growth": 66, "margin_from": "-14%", "margin_to": "+7.5%"},
          {"name": "Manufacturing", "icon": _ic("factory"), "count": "109,063",
           "note": "฿21.1T revenue, the largest of these four",
           "growth": 30, "margin_from": "4.3%", "margin_to": "4.3% steady"},
          {"name": "F&B & Hospitality", "icon": _ic("food"), "count": "48,564",
           "note": "The steepest recovery arc",
           "growth": 141, "margin_from": "-22%", "margin_to": "+7.4%"},
          {"name": "Tech & Information", "icon": _ic("chip"), "count": "29,027",
           "note": "Growth in profit, not just sales",
           "growth": 18, "margin_from": "5.9%", "margin_to": "7.5%"},
      ],
      callout="Fewer than 200 of 44,190 logistics companies exceed ฿1B revenue",
      source=SRC_MAXDATA)

b.add("feature_pick", **S("1 · Why now"),
      title="Drill one level down and the sweet spots appear",
      hero={
          "name": "Pet food manufacturing",
          "headline": "Thailand's quiet export champion, growing with double-digit margins",
          "stats": [{"value": "+93%", "label": "revenue growth since FY2020"},
                    {"value": "10.3%", "label": "net margin FY2025"},
                    {"value": "฿80B", "label": "FY2025 revenue, 258 companies"}],
          "line": "Margins held at 7% or better in every one of the last six years",
          "icon": _ic("food-white"),
      },
      runners=[
          {"name": "Cafés & beverage stands",
           "stat": "+288%",
           "note": "Fastest sub-industry in our screen: ฿8.2B to ฿31.9B, companies nearly doubled"},
          {"name": "Meat & poultry processing",
           "stat": "+126%",
           "note": "Revenue up every single year, and margin turned positive in FY2024"},
          {"name": "Bakery products",
           "stat": "198",
           "note": "Mid-market companies at ฿10M-1B: the deepest acquisition bench in food"},
      ],
      caveat="Our analysts strip reclassification artifacts before ranking: the three raw "
             "growth leaders in manufacturing fail that test and are excluded.",
      source=SRC_MAXDATA)

b.add("fork_road", **S("1 · Why now"),
      title="Two roads to growth, and only one of them is fast",
      source="Max Solutions M&A advisory practice")

b.add("dimension_table", **S("1 · Why now"),
      title="M&A used to be a big-company game. Not anymore",
      left_header="M&A then", right_header="M&A now",
      rows=[
          {"dim": "Who buys",
           "left": "Conglomerates and funds",
           "right": "Listed corporates, mid-size firms, high-net-worth individuals"},
          {"dim": "Deal size",
           "left": "Billions",
           "right": "SME targets: ฿10M to ฿1B revenue"},
          {"dim": "Payback",
           "left": "10+ years",
           "right": "3 to 8 years in our deal experience"},
          {"dim": "Sourcing",
           "left": "Investment banks",
           "right": "Communities, platforms, and data"},
      ],
      chip={"text": "16x: Japan's SME succession M&A growth in 8 years"},
      source="Max Solutions deal experience; Japan figure: METI via Kobe University MAREC, 2014-2022")

b.add("column_comparison", **S("1 · Why now"),
      title="Thai M&A just woke up: the strongest quarter in years",
      categories=["Q1 2025", "Q2 2025", "Q3 2025", "Q4 2025", "Q1 2026"],
      values=[1.0, 3.4, 2.5, 1.3, 7.9],
      focus_index=4,
      description="Q1 2026: 57 deals worth USD 7.9B. Even without True Corporation's USD 3.9B stake sale, the quarter beats every quarter of 2025",
      takeaway_header=None,
      axis=False,
      data_label="Announced Thailand deal value", data_unit="USD billions",
      source="KPMG Thailand M&A Trends quarterlies 2025-Q1 2026; ASEAN: Lyndon Advisory")

b.add("stat_hero_navy", **S("1 · Why now"),
      bg_path=f"{GEN}/backdrops/hero-waffle-glow.png",
      title_eyebrow="The succession wave is the supply side of this market",
      stat="81%",
      stat_label="of Thailand's 20-year-plus companies have no next-generation director on the board",
      waffle_filled=81,
      waffle_caption="111,304 of about 137,000 operating companies over 20 years old, "
                     "with no younger-generation director on the board",
      method=[
          {"text": "Take every company in the registry that has traded for 20 years or more"},
          {"text": "Group the directors on each board into generations, using the era their registry records date from"},
          {"text": "Count the boards where nobody from the younger generation has been added"},
      ],
      closing="These companies need successors. Increasingly, that means buyers.",
      source="Max Data analysis of 4.2M public directorship records, July 2026. Generation is inferred from registry record era, not from any personal data we hold")

b.add("quote_breather", **S("1 · Why now"),
      title="Why we started three years ago",
      quote="We watched a generation of Thai founders build companies their "
            "children didn't want to run. Someone had to build the bridge "
            "between those owners and the people ready to take over.",
      author="Chatpong Lappitakpong",
      author_title="MD, Max Solutions",
      photo_path=f"{ASSETS}/people/max-headshot.png",
      source="Max Solutions")

# =====================================================================
# SECTION 2 · WHERE DEALS COME FROM
# =====================================================================

b.add("art_divider", art_path=f"{GEN}/dividers/div-2.png",
      section_number="02",
      section_title="Where deals actually come from",
      subtitle="From two million registered companies to the handful worth your time")

b.add("taper_funnel", **S("2 · The access"),
      title="Finding the needle: Thailand's company universe",
      stages=[
          {"name": "Registered juristic persons", "value": "1,989,709",
           "description": "Every company ever registered in Thailand"},
          {"name": "Operating today", "value": "993,373",
           "description": "Half the registry is already defunct or dissolved"},
          {"name": "Filing real revenue", "value": "648,304",
           "description": "Revenue above zero in FY2025"},
          {"name": "SME band ฿10M-1B", "value": "176,962",
           "description": "The deal size this webinar is about"},
          {"name": "Core ฿100M-1B", "value": "37,045",
           "description": "Only 1 in 18 revenue-reporting companies"},
          {"name": "Your shortlist", "value": "10-30",
           "description": "What a focused thesis produces"},
      ],
      chip="1 in 18 companies sits in the core band",
      source=SRC_MAXDATA)

b.add("ceilings", **S("2 · The access"),
      title="How everyone finds deals today, and where each channel stops",
      takeaway="Each channel works. Each has a ceiling you will hit.",
      channels=[
          {"name": "Word of mouth", "value": "2-3", "value_label": "DEALS A YEAR", "ceiling": "2-3 deals a year, if lucky",
           "note": "Friends, suppliers, golf partners. High trust, tiny reach"},
          {"name": "Online listings", "value": "100s", "value_label": "LISTINGS, UNEVEN QUALITY", "ceiling": "the good ones go fast",
           "note": "Marketplaces and brokers' sites. Wide reach, uneven quality"},
          {"name": "Intermediaries", "value": "1", "value_label": "MANDATE AT A TIME", "ceiling": "needs a clear mandate",
           "note": "Advisors and boutique firms. Screened deals, real process"},
      ],
      breakout="Breaking all three ceilings at once is what the rest of this section is about",
      source="Max Solutions market observation")

b.add("access_ladder", **S("2 · The access"),
      title="Our answer: three doors, one ecosystem",
      steps=[
          {"kicker": "Door 1 · Community", "icon": _ic("door-white"), "stat": "80,000+ members",
           "bullets": ["DealFlow Facebook community",
                       "Off-market chatter surfaces here first"]},
          {"kicker": "Door 2 · Marketplace", "icon": _ic("door-white"), "stat": "100+ live deals",
           "bullets": ["DealFlow Market listings",
                       "Screened sellers, 15 industries"]},
          {"kicker": "Door 3 · Advisory", "icon": _ic("door-white"), "stat": "150+ SMEs/yr",
           "bullets": ["Full-mandate M&A advisory",
                       "Our deal team runs it end to end"]},
      ],
      source="Max Solutions, July 2026")

b.add("screenshot_slide", **S("2 · The access"),
      title="Inside the community: where off-market deals surface",
      placeholder_label="Screenshot: DealFlow Facebook community",
      image_path=f"{ASSETS}/shots/fb-group-clean.png",
      image_caption="ขายกิจการ ขายโรงงาน หานักลงทุน, our public Facebook group. Search the name to find it",
      kicker="Door 1 · Community",
      claim="Sellers test the water here before any listing exists",
      bullets=["Thailand's largest SME buy-and-sell group, free to join",
               "Owners, buyers, and advisors in one room",
               "Posts are occasional, but this is where sellers surface first"],
      overlap_stat={"value": "80,100", "label": "members and counting"},
      layout_mode="right",
      source="DealFlow community, July 2026")

b.add("screenshot_slide", **S("2 · The access"),
      title="Inside DealFlow Market: screened deals, real numbers",
      placeholder_label="Screenshot: DealFlow Market listings page",
      image_path="/home/user/mckinsey-pptx/webinar/assets/shots/dealflow-market.png",
      image_caption="dealflowmarket.com/listing, live capture July 2026",
      kicker="Door 2 · Marketplace",
      claim="Every listing verified before it goes up",
      bullets=["Real asking prices, real revenue, real provinces",
               "Filter by sector, size, and location"],
      stats=[{"value": "100+", "label": "live deals"},
             {"value": "15", "label": "industries"}],
      layout_mode="left",
      source="DealFlow Market, July 2026")

b.add("metro_line", **S("2 · The access"),
      title="Max Data: our AI analytics platform, built for every stage",
      stops=[
          {"name": "Research", "icon": _ic("chip"),
           "description": "Market and industry trends"},
          {"name": "Source", "icon": _ic("target"),
           "description": "Filter to targets that fit your thesis"},
          {"name": "Validate", "icon": _ic("book"),
           "description": "10 years of financials on any company"},
          {"name": "Diligence", "icon": _ic("magnifier"),
           "description": "Directors, licences, branches, contracts"},
          {"name": "Reach", "icon": _ic("handshake"),
           "description": "Contact the actual decision maker"},
      ],
      input_note="1.99M companies in",
      output_note="the decision maker out",
      source="Max Data platform, July 2026")

b.add("screenshot_slide", **S("2 · The access"),
      title="Max Data in 90 seconds",
      placeholder_label="Screenshot or live demo: Max Data screener",
      placeholder_note="Suggested demo: Logistics, revenue ฿10M-1B → 9,399 live targets",
      image_path=("/home/user/mckinsey-pptx/webinar/assets/shots/maxdata-screener.png"
                  if __import__("os").path.exists(
                      "/home/user/mckinsey-pptx/webinar/assets/shots/maxdata-screener.png")
                  else None),
      image_caption="Max Data screener: Logistics, ฿10M-1B revenue → 9,399 companies",
      claim="Type a thesis. Get a board-ready target list in one afternoon.",
      stats=[{"value": "1.99M", "label": "companies tracked"},
             {"value": "9.8M", "label": "financial statements"},
             {"value": "100k+", "label": "companies with director contacts"}],
      layout_mode="hero",
      source="Max Data platform, July 2026")

b.add("screenshot_slide", **S("2 · The access"),
      title="Step 1: size the niche before you pick a target",
      placeholder_label="Screenshot: Max Data market page",
      image_path=f"{ASSETS}/shots/md-market-top.png",
      image_caption="Max Data market page, TSIC 49331: road transport of refrigerated freight",
      kicker="Cold chain, worked example",
      claim="484 companies, ฿5.9B, and nobody owns it",
      bullets=["Market size, growth and survival rate for any of 1,265 industry codes",
               "Concentration score says perfect competition, so no incumbent blocks a new entrant",
               "The top ten are ranked for you, which is where a buy-side shortlist starts"],
      stats=[{"value": "484", "label": "registered companies"},
             {"value": "฿5.9B", "label": "market revenue"},
             {"value": "83.5%", "label": "survival rate"}],
      layout_mode="hero",
      shot_width=11.9,
      source="Max Data market module, July 2026")

b.add("screenshot_slide", **S("2 · The access"),
      title="Step 2: open the leader and read twenty years in one screen",
      placeholder_label="Screenshot: Max Data company page",
      image_path=f"{ASSETS}/shots/md-company-header.png",
      image_caption="Max Data company page. Every figure traces to a filed statement",
      kicker="Cold chain, worked example",
      claim="What they do, what they own, and who sits on the board",
      bullets=["Revenue, profit, assets and equity with the year-on-year move",
               "Registered capital, company age, exact TSIC code and address",
               "The board, so you already know the name before you reach out"],
      stats=[{"value": "฿995.8M", "label": "FY2025 revenue, +37.1%"},
             {"value": "20.44%", "label": "20-year revenue CAGR"},
             {"value": "16.86%", "label": "share of its TSIC"}],
      layout_mode="hero",
      shot_width=11.9,
      source="Max Data company page, TO.TOO Chiang Mai Seafood, July 2026")

b.add("screenshot_slide", **S("2 · The access"),
      title="Step 3: turn a name on a board into an email you can send",
      placeholder_label="Screenshot: Max Data E-Finding",
      image_path=f"{ASSETS}/shots/maxdata-efinding-blur.png",
      image_caption="Max Data E-Finding. Name and result masked here for privacy, live in the product",
      shot_width=10.8,
      kicker="Cold chain, worked example",
      claim="The last mile most buyers never cross",
      bullets=["Give the engine a director name and the company domain",
               "It generates, cross-references and confirms, then stops when it is sure",
               "You reach the decision maker, not a general enquiries inbox"],
      stats=[{"value": "100k+", "label": "companies with contacts"},
             {"value": "High", "label": "confidence on this result"},
             {"value": "1 click", "label": "single or bulk CSV"}],
      layout_mode="hero",
      source="Max Data E-Finding, July 2026")

b.add("screenshot_slide", **S("2 · The access"),
      title="All of it in one workflow, in one day",
      placeholder_label="Screenshot: Max Data research cockpit",
      image_path="/home/user/mckinsey-pptx/webinar/assets/shots/maxdata-cockpit.png",
      image_caption="Max Data research cockpit. Verification and contact columns are generated per company, then checked against the registry",
      kicker="Thesis to contact, one workflow",
      claim="Thesis, research and contact in a single sheet",
      bullets=["A typical team needs more than a week to build a list and "
               "find the contacts, working across three or four tools",
               "Here it is one question, one sheet, and every column is "
               "labelled with where the answer came from",
               "Built on Max Data and run by AI, with a human approving "
               "each step"],
      stats=[{"value": "1 day", "label": "thesis to call list"},
             {"value": "1 week+", "label": "the manual alternative"},
             {"value": "1 sheet", "label": "instead of four tools"}],
      layout_mode="hero",
      shot_width=11.9,
      source="Max Data research cockpit, July 2026. Director details are public registry records")

# =====================================================================
# SECTION 3 · WHAT GOOD LOOKS LIKE
# =====================================================================

b.add("art_divider", art_path=f"{GEN}/dividers/div-3.png",
      section_number="03",
      section_title="What a good target looks like",
      subtitle="The lens our deal team applies before we ever talk price")

b.add("poll_slide", **S("3 · The lens"),
      title="Quick pulse check",
      question="You are shown a company for sale. What do you check first?",
      options=["The financial statements",
               "The owner and why they are selling",
               "The customers and recurring revenue",
               "The competition around it"],
      instruction="Vote now. We'll show you our order in a minute",
      ornament_path=f"{GEN}/backdrops/poll-corner-alpha.png")

b.add("chevron_flags", **S("3 · The lens"),
      title="The Green 5: what our analysts look for in every target",
      subtitle="Five signals our deal team scores before anyone talks price.",
      flags=[
          {"keyword": "Synergy", "icon": _ic("target-white"),
           "descriptor": "Multiplies what you already own",
           "ask": "What would this business do with my customers?"},
          {"keyword": "Potential", "icon": _ic("flag-white"),
           "descriptor": "Obvious headroom the owner never used",
           "ask": "What have you never had time to try?"},
          {"keyword": "Management", "icon": _ic("people-white"),
           "descriptor": "Runs without the owner in the room",
           "ask": "What breaks if you take a month off?"},
          {"keyword": "Revenue quality", "icon": _ic("book-white"),
           "descriptor": "Recurring customers, repeat contracts",
           "ask": "Which customers bought again last year?"},
          {"keyword": "Position", "icon": _ic("storefront-white"),
           "descriptor": "A defensible niche, not a price war",
           "ask": "Why do customers pick you over the cheaper option?"},
      ],
      rule="Three or more strong is worth a conversation. Fewer, and you are buying a job, not a business.",
      source="Max Solutions target evaluation framework")

b.add("keyline_panels", **S("3 · The lens"),
      title="The Red 3: where deals die in Thailand",
      subtitle="Our deal team's kill criteria, in the order they usually surface",
      panels=[
          {"label": "Business risk",
           "bullets": ["Shrinking market, or 1-2 customers carrying the model",
                       "Value that walks out with the founder"]},
          {"label": "Legal risk",
           "bullets": ["Missing licences, permits, or land title",
                       "A deal-breaker: you cannot legally buy what isn't legal"]},
          {"label": "Financial risk",
           "bullets": ["Two sets of books is common in Thai SMEs",
                       "Verify cash, not stories"]},
      ],
      closing="One red flag unresolved = walk away.",
      source="Max Solutions deal experience, 150+ SME mandates per year")

b.add("ledgers", **S("3 · The lens"),
      title="The two-books problem, and why diligence pays for itself",
      source="Max Solutions deal experience")

b.add("scorecard_slide", **S("3 · The lens"),
      title="Take this home: the 60-second target scorecard",
      subtitle="Screenshot this. Rate every target Weak or Strong per line. One Weak in the Red 3 kills the deal",
      groups=[
          {"name": "THE GREEN 5 · SCORE THE UPSIDE", "color": "green",
           "rows": [
               {"name": "Synergy with what you own",
                "weak": "Standalone business, no overlap",
                "strong": "Multiplies your customers, licences, or capacity"},
               {"name": "Untapped potential",
                "weak": "Owner already squeezed everything",
                "strong": "No marketing, no exports, obvious headroom"},
               {"name": "Runs without the owner",
                "weak": "Founder holds every relationship",
                "strong": "Managers and systems run the week"},
               {"name": "Recurring revenue",
                "weak": "One-off projects, lumpy income",
                "strong": "Repeat contracts and returning customers"},
               {"name": "Defensible niche",
                "weak": "Price war with many rivals",
                "strong": "Few real competitors, switching costs"},
           ]},
          {"name": "THE RED 3 · CLEAR EVERY ONE", "color": "red",
           "rows": [
               {"name": "Business risk",
                "weak": "Fad demand or 1-2 customers",
                "strong": "Durable demand, spread customer base"},
               {"name": "Legal risk",
                "weak": "Licences missing or land unclear",
                "strong": "Every permit verified in diligence"},
               {"name": "Financial risk",
                "weak": "Two books, unverifiable cash",
                "strong": "Statements reconcile to bank accounts"},
           ]},
      ],
      source="Max Solutions target evaluation framework")

# =====================================================================
# SECTION 4 · HOW A DEAL RUNS
# =====================================================================

b.add("art_divider", art_path=f"{GEN}/dividers/div-4.png",
      section_number="04",
      section_title="How a deal actually runs",
      subtitle="The buy-side path in six steps, and the seller's mirror image")

b.add("rising_road", **S("4 · The path"),
      title="The buy-side path: six steps from thesis to keys",
      subtitle="In our experience a focused buyer goes thesis-to-keys in roughly 6 to 12 months",
      steps=[
          {"name": "Thesis",
           "description": "What you want to own and why"},
          {"name": "Shortlist",
           "description": "993k companies down to 10-30 fits"},
          {"name": "Approach",
           "description": "Reach the owner, build trust, sign an NDA"},
          {"name": "LOI",
           "description": "Letter of intent: a written offer, not yet binding"},
          {"name": "Diligence",
           "description": "Clear the Red 3"},
          {"name": "Close & handover",
           "description": "Sign, pay, transition the team"},
      ],
      source="Max Solutions buy-side playbook")

b.add("mirror_spine", **S("4 · The path"),
      title="Selling? Same road, driven in reverse",
      stages=[
          {"stage": "Prepare",
           "left": "Write a thesis",
           "right": "Clean numbers and a defensible story"},
          {"stage": "Go to market",
           "left": "Shortlist and approach targets",
           "right": "Quiet outreach, collect offers"},
          {"stage": "Diligence",
           "left": "Verify the Green 5 and Red 3",
           "right": "Buyers will check them in you"},
          {"stage": "Close",
           "left": "Pay and take over",
           "right": "Hand over clean, on your terms"},
      ],
      source="Max Solutions sell-side playbook. Full process guide on our website")

# =====================================================================
# SECTION 5 · PROOF
# =====================================================================

b.add("art_divider", art_path=f"{GEN}/dividers/div-5.png",
      section_number="05",
      section_title="Proof it works",
      subtitle="Two deals our team closed this year, and the global wave behind them")

b.add("profile_cards", **S("5 · Proof"),
      title="The supply is real: companies like these are in the registry right now",
      headline_stat="14,139",
      headline_label="companies match the succession profile: 20+ years old, "
                     "฿50M-1B revenue, profitable, every director born before 1984",
      section_chips=[
          {"value": "6,185", "label": "WHOLESALE/RETAIL"},
          {"value": "3,547", "label": "MANUFACTURING"},
          {"value": "637", "label": "LOGISTICS"},
          {"value": "354", "label": "FOOD/HOSPITALITY"},
      ],
      profiles=[
          {"sector": "Freight forwarding & customs",
           "province": "SAMUT PRAKAN",
           "facts": ["40 years operating, founded mid-1980s",
                     "~฿260M revenue, ~18% net margin",
                     "2 directors, both born before 1984"]},
          {"sector": "Corrugated packaging manufacturer",
           "province": "BANGKOK",
           "facts": ["31 years operating, founded mid-1990s",
                     "~฿560M revenue, ~15% net margin",
                     "2 directors, both born before 1984"]},
          {"sector": "Hotel & resort operator",
           "province": "CHON BURI",
           "facts": ["49 years operating, founded late 1970s",
                     "~฿260M revenue, ~10% net margin",
                     "2 directors, both born before 1984"]},
      ],
      disclaimer="Real registry rows, anonymized. Revenue rounded to ฿10M, margins to whole %. "
                 "Chips show the four sectors this room plays in, and 3,416 more matches sit in other "
                 "sectors. The screen floors revenue at ฿50M to keep targets bankable. "
                 "Our deal team holds the full list.",
      source=SRC_MAXDATA)

b.add("stat_hero_split", **S("5 · Proof"),
      title="The search-fund wave: individuals now buy companies, profitably",
      stat="35.1%",
      stat_label="average annual return across the search funds Stanford GSB tracks",
      rows=[
          {"value": "4.5x", "label": "money returned per ฿1 invested"},
          {"value": "94", "label": "funds launched in 2023, a record"},
          {"value": "320", "label": "international funds by end-2023"},
          {"value": "$11.7M", "label": "median international deal price"},
      ],
      kicker="Aggregate across funds that reported, so read it as the upside case, not the average outcome. The playbook still travels.",
      source="Stanford GSB Search Fund Study 2024 (Case E-870); IESE International Search Funds 2024")

b.add("case_slide", **S("5 · Proof"),
      title="Case one: the fire-safety distributor",
      case_name="Project FireGuard · closed Q4 2025",
      sector_chip="FIRE SAFETY / TRADING",
      situation=["Husband-and-wife owners near retirement, no successor",
                 "฿100M to ฿200M revenue, ฿30M to ฿50M operating profit",
                 "Decades of trade relationships, run entirely on the owners"],
      outcome=["New owner installed an ERP system and tracked inventory properly",
               "That surfaced the handful of products actually driving revenue",
               "Focused on those, and earnings are now on a double-digit growth track"],
      bridge_stat="< 6 months to visible turnaround",
      kpis=[{"value": "฿100-200M", "label": "Revenue at deal"},
            {"value": "฿30-50M", "label": "Operating profit at deal"},
            {"value": "Double digit", "label": "earnings growth on track"}],
      photo_path=f"{GEN}/scenes/case-fire.png",
      source="Max Solutions deal team, 2026. Figures approximate to protect the parties")

b.add("case_slide", **S("5 · Proof"),
      title="Case two: the pizza-oven supplier",
      case_name="Project Pizza Oven · cross-border buyer",
      sector_chip="F&B EQUIPMENT / SUPPLY",
      situation=["One of Thailand's top pizza-oven manufacturers and distributors",
                 "Already profitable, with 6+ months of pre-orders booked",
                 "Owners returning home, and short of the capital to expand production"],
      outcome=["An overseas buyer acquired it end to end through our process",
               "Fresh capital expanded production to meet the waiting demand",
               "Lead time is coming down from 6 months toward 1 to 2 months"],
      bridge_stat="6 months of backlog, unlocked",
      kpis=[{"value": "6 to 1-2 mo", "label": "lead time after the capital went in"},
            {"value": "Double digit", "label": "revenue growth expected"},
            {"value": "Cross-border", "label": "overseas buyer, Thai target"}],
      photo_path=f"{GEN}/scenes/case-oven.png",
      source="Max Solutions deal team, 2026. Figures approximate to protect the parties")

_sl_caus = b.add("causality_band", **S("5 · Proof"),
      title="Why small deals turn around so fast",
      subtitle="The pattern behind the deals that worked. We see it in the ones that stall too, which is why the Red 3 exists",
      blocks=[
          {"label": "Honest sellers",
           "support": "Price reflects the exit need, not a bidding war"},
          {"label": "Inefficiency is the upside",
           "support": "No ERP, no marketing, no pricing discipline"},
          {"label": "Capable buyers",
           "support": "Corporate resources or sharp operators"},
      ],
      result="Margins move in months, not years. FireGuard did it in under six months",
      source="Max Solutions deal experience across 150+ SME mandates per year")
_pic(_sl_caus, f"{GEN}/scenes/succession.png", 4.97, 4.85, 3.4, 1.35)

# =====================================================================
# SECTION 6 · ACT
# =====================================================================

b.add("art_divider", art_path=f"{GEN}/dividers/div-6.png",
      section_number="06",
      section_title="Your next step",
      subtitle="Three frameworks to keep, one action to take tonight")

b.add("recap_cards", **S("6 · Next step"),
      title="What you now have",
      cards=[
          {"takeaway": "A reason to move now", "icon": _ic("flag"),
           "line": "The succession wave is bringing good companies to market"},
          {"takeaway": "A way in", "icon": _ic("door"),
           "line": "Three doors and the Max Data layer"},
          {"takeaway": "A lens and a path", "icon": _ic("magnifier"),
           "line": "The Green 5, the Red 3, the scorecard, and the 6 steps"},
      ],
      conclusion="The buyers who win the succession wave start looking before everyone else.")

b.add("service_spectrum", **S("6 · Next step"),
      title="Research and sourcing from here: three lanes, one call",
      axis_left="We carry the work",
      axis_right="You run it yourself",
      lanes=[
          {"logo": f"{ASSETS}/maxsolutions-logo.png", "site": "maxsolutions.co.th",
           "kicker": "Full-mandate advisory", "name": "Max Solutions",
           "line": "We do everything with you: research, sourcing, "
                   "negotiation, close.",
           "bullets": ["A dedicated deal team on your mandate end to end",
                       "You decide, we run the machine"],
           "stat": "150+", "stat_label": "SME mandates a year"},
          {"logo": f"{ASSETS}/logos/dealflow-logo.png", "site": "dealflowmarket.com",
           "kicker": "Screened marketplace", "name": "DealFlow Market",
           "line": "You screen the deals yourself, we keep them real.",
           "bullets": ["Live, verified listings with real numbers",
                       "Move at your own shortlist speed"],
           "stat": "100+", "stat_label": "live deals across 15 industries"},
          {"logo": f"{ASSETS}/logos/maxdata-logo.png", "site": "maxdatathailand.com",
           "kicker": "Self-serve platform", "name": "Max Data",
           "line": "Your own research desk for the Thai market.",
           "bullets": ["Research, source, and track targets on your own",
                       "Local intelligence and decision-maker reach built in"],
           "stat": "1.99M", "stat_label": "companies, 10 years of financials"},
      ],
      band_text="Whichever lane fits, schedule a call. The first five "
                "bookings get a customized session: your industry, your "
                "goal, and exactly how we'd research and source for you.",
      band_chip="First 5 bookings",
      source="Max Solutions, July 2026")

b.add("cta_slide", **S("6 · Next step"),
      title="Do one of these before you log off",
      paths=[
          {"num": 1, "icon": _ic("target"), "who": "Buyers & investors", "action": "Book a free opportunity scan",
           "detail": "We map live targets against your thesis, in a session your board can act on."},
          {"num": 2, "icon": _ic("book"), "who": "Owners", "action": "Get a confidential valuation talk",
           "detail": "Know what your business is worth and what buyers would flag, no obligation."},
          {"num": 3, "icon": _ic("people"), "who": "Connectors & everyone", "action": "Join the DealFlow community",
           "detail": "80,000+ members. Refer a buyer or seller and our referral program pays you."},
      ],
      bottom_actions=["Scan the QR to claim one of tonight's five consultation slots",
                      "Or type 1 (buying), 2 (selling), or 3 (referring) in the chat and we will send you the link"],
      qr_label="QR: booking page")

b.add("qa_slide",
      bg_path=f"{GEN}/backdrops/qa-waves.png",
      line="Ask us anything. We stay until your questions run out.",
      qr_caption="Scan to claim one of five slots · or type 1 / 2 / 3 in the chat")

b.add("credits_slide",
      groups=[
          {"to": "To everyone here tonight",
           "line": "Thank you for spending your Sunday evening with us."},
          {"to": "To the Max Solutions team",
           "line": "The moderator, our events crew, Khun Aoy for the client "
                   "insights, and Khun Vipin for the deal expertise behind "
                   "every framework tonight."},
      ],
      finale="And to my partner Pim, who inspired this, helped build every "
             "part of it, and has carried me through the whole journey.",
      finale_name="Thank you, Pim")

b.add("thank_you",
      bg_path=f"{GEN}/backdrops/thanks-skyline.png",
      lines=["Max Solutions · DealFlow Market · Max Data",
             "See you in the deal flow"],
      contact_placeholder="QR + contacts: LINE, email, phone")

# =====================================================================
# APPENDIX
# =====================================================================

b.add("art_divider", art_path=f"{GEN}/dividers/div-A.png",
      section_number="A",
      section_title="Appendix",
      subtitle="For the replay: process detail, glossary, and methodology")

b.add("vertical_ladder", **S("Appendix"),
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

b.add("mandate_grid", **S("Appendix"),
      title="Live mandates we are running right now",
      subtitle="A sample of current sell-side engagements. Project names are anonymized, figures are as presented by the seller",
      groups=[
          {"sector": "Food & beverage", "rows": [
              {"project": "FineDine", "what": "Michelin Guide listed fine dining restaurant",
               "rev": "฿11.6M", "ebitda": "฿3.5M", "ask": "฿27M"},
              {"project": "KBBQ1", "what": "Korean BBQ buffet restaurant chain",
               "rev": "฿52.2M", "ebitda": "฿9.1M", "ask": "฿40M"},
              {"project": "MooPing", "what": "Thai grilled skewer chain",
               "rev": "฿19.5M", "ebitda": "฿3.0M", "ask": "฿25M"},
              {"project": "OceanFreeze", "what": "Seafood processor and distributor",
               "rev": "฿47.4M", "ebitda": "฿3.7M", "ask": "฿45M"}]},
          {"sector": "Healthcare & aesthetics", "rows": [
              {"project": "Oasis", "what": "One-stop aesthetic and wellness clinic",
               "rev": "฿11.1M", "ebitda": "฿1.6M", "ask": "฿15M"},
              {"project": "Aura8", "what": "Aesthetics clinic platform, 8 branches",
               "rev": "฿97.6M", "ebitda": "฿23.3M", "ask": "฿150M"},
              {"project": "Glowshot", "what": "Two-branch aesthetics clinic platform",
               "rev": "฿77.6M", "ebitda": "฿26.3M", "ask": "฿150M"},
              {"project": "Dental2", "what": "Two-branch dental clinics in Bangkok",
               "rev": "฿30.0M", "ebitda": "฿7.0M", "ask": "฿40M"}]},
      ],
      note="Every mandate here came to us through the same three doors you saw tonight.",
      source="Max Solutions live mandate book, July 2026")

b.add("mandate_grid", **S("Appendix"),
      title="Live mandates, continued",
      subtitle="Manufacturing, industrials, consumer and services",
      groups=[
          {"sector": "Manufacturing & industrials", "rows": [
              {"project": "SteelFaceTech", "what": "Hardfacing steel plate manufacturer",
               "rev": "฿102.2M", "ebitda": "฿13.9M", "ask": "฿180M"}]},
          {"sector": "Construction & infrastructure", "rows": [
              {"project": "TwinLift", "what": "Full-service industrial crane and lifting provider",
               "rev": "฿105.0M", "ebitda": "฿13.0M", "ask": "฿120M"},
              {"project": "SkyLift", "what": "Aerial work platform rental operator",
               "rev": "฿48.0M", "ebitda": "฿23.0M", "ask": "On request"}]},
          {"sector": "Consumer & retail", "rows": [
              {"project": "CoreWear", "what": "Mass-market omnichannel fashion brand",
               "rev": "฿76.7M", "ebitda": "฿10.4M", "ask": "฿55M"},
              {"project": "ProSupply", "what": "Premium interior manufacturing and supply",
               "rev": "฿48.3M", "ebitda": "฿9.5M", "ask": "฿80M"}]},
          {"sector": "Logistics & services", "rows": [
              {"project": "ThaiTransport", "what": "Full-service third-party logistics",
               "rev": "฿90.2M", "ebitda": "฿28.6M", "ask": "฿270M"},
              {"project": "2B Laundry", "what": "Laundromat franchisor, 20+ years operating",
               "rev": "฿20.3M", "ebitda": "฿3.7M", "ask": "฿22M"},
              {"project": "FlourTrading", "what": "Tapioca flour and agri commodity trading",
               "rev": "฿365.0M", "ebitda": "฿18.0M", "ask": "฿120M"}]},
      ],
      note="Ask prices are the seller's opening position, not a valuation. Every one is negotiable.",
      source="Max Solutions live mandate book, July 2026")

b.add("stat_band", **S("Appendix"),
      title="We also cover fintech and payments",
      stats=[
          {"glyph": "donut20", "value": "€60M",
           "label": "company valuation on our most recent fintech mandate",
           "context": "cross-border mandate, priced in euros"},
          {"glyph": "bar80", "value": "~€6M",
           "label": "growth equity raised, about a 10% stake",
           "context": "Series A capital raise"},
          {"glyph": "steps", "value": "5",
           "label": "sub-sectors covered end to end",
           "context": "acquiring, processing and switching, SaaS and "
                      "embedded finance, lending and credit, digital wallets"},
          {"glyph": "bar80", "value": "SaaS",
           "label": "recurring-revenue models with bank and FI clients",
           "context": "the segment buyers pay up for"},
      ],
      kicker="Dedicated coverage from transaction-processing infrastructure to SaaS platforms serving banks.",
      source="Max Solutions fintech and payments practice")

b.add("definition_list", **S("Appendix"),
      title="Glossary: five terms you heard tonight",
      subtitle="Plain-language definitions for the replay",
      terms=[
          {"term": "LOI",
           "definition": "Letter of intent: a non-binding offer that states price range and terms before diligence"},
          {"term": "EBITDA",
           "definition": "Earnings before interest, tax, depreciation, amortization: the profit buyers actually price on"},
          {"term": "Due diligence",
           "definition": "The verification phase: business, legal, and financial checks before closing"},
          {"term": "Search fund",
           "definition": "An investor-backed vehicle where one operator searches for, buys, and runs a single SME"},
          {"term": "Succession deal",
           "definition": "A sale driven by owner retirement with no family successor, Thailand's fastest-growing deal type"},
      ],
      source="Max Solutions")

b.add("method_grid", **S("Appendix"),
      title="Methodology and sources",
      blocks=[
          {"header": "The data",
           "body": "Registry statistics come from Max Data, our AI-driven M&A "
                   "analytics platform covering 1.99 million Thai juristic persons, "
                   "9.8 million financial statements, and 4.2 million directorships, "
                   "snapshot July 2026. Fiscal years follow Thai filing years. "
                   "FY2025 filings were about 98% complete at analysis time."},
          {"header": "Growth method",
           "body": "Industry growth compares revenue filed in FY2020 against FY2025 "
                   "at TSIC section and industry level. No whole section declined in "
                   "nominal terms over that window. The laggards shown are genuine "
                   "contracting industries at TSIC level, verified year by year."},
          {"header": "Succession cohorts",
           "body": "Succession figures use birth cohorts encoded in Thai national-ID "
                   "structure: IDs issued before the 1984 system change identify "
                   "directors born before 1984, a hard age floor of 42+ today. The "
                   "registry holds no birthdates, so we state cohorts, not ages. To "
                   "our knowledge this analysis is the first of its kind."},
          {"header": "External sources",
           "body": "Cited on the relevant slides: KPMG Thailand quarterly M&A reports, "
                   "NESDC demographic data, Grant Thornton Thailand, METI Japan, and "
                   "Stanford GSB search fund research. Case figures are rounded and "
                   "lightly disguised to protect client confidentiality."},
      ],
      source="Max Solutions research team, July 2026")

# =====================================================================
# SPEAKER NOTES — timing budgets + delivery cues from the planning session
# =====================================================================

NOTES = {
    1: 'Doors open 20:25. Greet by name in chat. Start 20:32 sharp. [T+0:00 · MC]',
    2: 'MC launches poll 1 and reads the result aloud so we can tailor examples. 2 min. [T+0:03]',
    3: 'One breath per stop. Promise: by the end you know where deals come from, how to judge one, and your first step. [T+0:05]',
    4: "Tie back to the poll: 'here is your row.' One line per seat, under a minute. [T+0:06]",
    5: '30 seconds each. HAND TO VIPIN at the end of this slide. [T+0:07]',
    6: 'SECTION 1 · WHY NOW. VIPIN presents. Budget 13 minutes, hand back by 20:53. [T+0:08]',
    7: 'Big line: COVID knocked four trillion baht off. Say the 5.9% peak-to-peak figure out loud, it buys credibility.',
    8: "Pause on the gold factoid. Chat prompt: 'type your industry in the chat.' 60 seconds.",
    9: 'Weak segments are where motivated sellers live. 45 seconds, do not dwell.',
    10: "This room's four sectors. Logistics people: 9,399 mid-market targets is YOUR number.",
    11: 'Pet food only, 45 seconds. Skip the runners-up if the clock is tight.',
    12: 'THE HINGE. Organic versus inorganic. Slow down, this is the argument for M&A at all. 2 min.',
    13: 'M&A is not a big-company game anymore. The 16x Japan number is the proof point.',
    14: 'Q1 2026 woke up. Say the True caveat yourself before anyone asks.',
    15: "THE BIG ONE. 81%. Walk the three-step method aloud. Never say an age, say 'no next-generation director'. Add: 'Japan already does 3,000 of these a year.' [T+0:18]",
    16: 'Founder beat. 30 seconds, personal, then HAND BACK TO CHATPONG. [T+0:21]',
    17: 'SECTION 2 · WHERE DEALS COME FROM. CHATPONG presents. Budget 17 min. Densest stretch. [T+0:22]',
    18: '1.99M down to a shortlist. Note: a thesis is just a written sentence saying what you want to own. Vipin unpacks it in section 4.',
    19: 'Three channels, three ceilings. Sets up why an ecosystem beats any single channel.',
    20: 'Our answer: three doors. Describe, do not sell.',
    21: 'Real Facebook group, 80,100 members. Say the Thai group name so people can find it.',
    22: 'Real listings, real asking prices. 45 seconds.',
    23: 'Max Data overview. One line per stage. Note: diligence is the verification phase, Vipin has a slide on it.',
    24: 'The screener. 1.99M tracked, 9.8M statements, 100k+ with director contacts.',
    25: 'DEMO STEP 1. Cold chain, TSIC 49331. 484 companies, nobody owns it. HARD CAP the demo at 4 min. [T+0:31]',
    26: 'DEMO STEP 2. Open the leader. Twenty years of financials, and the board is right there.',
    27: 'DEMO STEP 3. Name plus domain equals a verified email. Personal fields blurred on purpose, say so.',
    28: 'The payoff: one day not one week, one sheet instead of four tools. HAND TO VIPIN. [T+0:38]',
    29: 'SECTION 3 · WHAT A GOOD TARGET LOOKS LIKE. VIPIN presents. Budget 11 minutes. [T+0:39]',
    30: 'Poll 2, or just a chat prompt if the clock is tight. 45 seconds either way.',
    31: 'The Green 5, 4 min. Read the ASK THE OWNER questions aloud, the room writes those down.',
    32: 'The Red 3, 3 min. Kill criteria. Firm tone.',
    33: "Two books, 90 seconds. Phrase carefully: 'a common practice, and fixable in diligence'. Never accusatory.",
    34: 'SCREENSHOT MOMENT. Tell the room to screenshot this. Pause 10 seconds. [T+0:48]',
    35: 'SECTION 4 · HOW A DEAL RUNS. VIPIN presents. 6 minutes, runs fast. [T+0:50]',
    36: 'Six steps, thesis to keys, 6 to 12 months. One line per step.',
    37: "The seller's mirror. 45 seconds, then HAND BACK TO CHATPONG. [T+0:55]",
    38: 'SECTION 5 · PROOF. CHATPONG presents. Budget 9 minutes. [T+0:56]',
    39: '14,139 match the profile. Three real anonymized rows. The supply side made concrete.',
    40: "Search funds. Say 'upside case, not the average' yourself.",
    41: 'FireGuard. The ERP story is the point: inefficiency is the upside.',
    42: 'Pizza oven. Cross-border buyer, six months of backlog unlocked by capital.',
    43: 'The pattern behind both. Own the selection bias out loud, then the FireGuard callback. [T+1:03]',
    44: 'SECTION 6 · YOUR NEXT STEP. CHATPONG presents. 5 minutes. [T+1:05]',
    45: 'Three take-homes, in the order the room experienced them. Fast.',
    46: 'The three lanes with logos and websites. Then the five-slot hook, unhurried.',
    47: 'THE ASK. Read the QR instruction aloud. Five slots, first come first served. [T+1:08]',
    48: 'Q&A. MC drives from the question bank. Full 20 minutes available. Close 22:00. [T+1:10]',
    49: 'Credits. Slow down on the last line.',
    50: 'Two QRs: left books a session, right joins the LINE group. Say which is which. [T+1:29]',
    51: 'Appendix. Only if Q&A pulls them up.',
    52: 'Sell-side process in full.',
    53: 'Live mandates. Strong if anyone asks what is actually available right now.',
    54: 'Live mandates continued.',
    55: 'Fintech and payments coverage.',
    56: 'Glossary.',
    57: 'Methodology. Pull this up if anyone challenges the 81% number.',
}

for _i, _slide in enumerate(b.prs.slides, 1):
    note = NOTES.get(_i)
    if note:
        _slide.notes_slide.notes_text_frame.text = note

editorial_slides.fix_baht_spacing(b.prs)
editorial_slides.flatten_all_shadows(b.prs)

out = "/home/user/mckinsey-pptx/webinar/output/ma-webinar-deck.pptx"
b.save(out)
print(f"saved {out} ({len(b.prs.slides._sldIdLst)} slides)")
