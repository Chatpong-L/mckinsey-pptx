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

GEN = "/home/user/mckinsey-pptx/webinar/assets/gen"


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
    "1 · Why now": 0, "2 · The lens": 1, "3 · The access": 2,
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
      instruction="Vote now in the poll panel",
      ornament_path=f"{GEN}/backdrops/poll-corner.png")

b.add("route_map", **S("Agenda"),
      title="Tonight's map: six stops",
      subtitle="Six stops, 90 minutes. The bar at the bottom of every slide shows where we are.",
      stops=[
          {"keyword": "WHY NOW", "descriptor": "The reshuffle and the succession wave"},
          {"keyword": "THE LENS", "descriptor": "5 green flags, 3 red flags"},
          {"keyword": "THE ACCESS", "descriptor": "Community, marketplace, advisory, data"},
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
           "stops": [1, 3, 4], "key_stop": 3},
          {"icon": _ic("book"), "who": "IR & listed-company leaders",
           "detail": "market narrative for boards and investors",
           "value": "A data-backed read on the Thai market, and the M&A "
                    "growth story your investors are starting to ask about.",
           "stops": [1, 3, 5], "key_stop": 1},
          {"icon": _ic("storefront"), "who": "Owners & founders",
           "detail": "exit, succession, or just curious",
           "value": "What buyers check first, what your business is worth, "
                    "and how a quiet, well-run exit actually happens.",
           "stops": [2, 4, 6], "key_stop": 2},
          {"icon": _ic("handshake"), "who": "Investors, advisors & explorers",
           "detail": "first deal or fiftieth",
           "value": "Where real deal flow surfaces, proof the model pays, "
                    "and how the referral economy around it works.",
           "stops": [3, 5, 6], "key_stop": 3},
      ],
      offer="Leading a team with one industry and one goal? We run tailored "
            "private sessions for leadership teams. Ask us after Q&A, or book "
            "with the QR at the end.",
      source="Registered attendee mix, July 2026")

b.add("speaker_panels", **S("Welcome"),
      title="Your guides tonight",
      speakers=[
          {"name": "[Max · full name]", "role": "Founder, Max Solutions",
           "bullets": ["Advises 150+ Thai SMEs per year on M&A",
                       "Built Max Data, our AI analytics platform"],
           "photo_label": "Photo: Max"},
          {"name": "[Khun Vipin · full name]", "role": "Head of M&A, Max Solutions",
           "bullets": ["Leads deal execution across 15 industries",
                       "[Placeholder: deals closed / years experience]"],
           "photo_label": "Photo: Khun Vipin"},
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
      growth_pct_first="+0.9%",
      growth_pct_second="+9.1%",
      description="COVID knocked ฿4T off the market. The recovery ran at 9.1% a year, and FY2025 sits 41% above the pre-COVID peak",
      takeaway_header=None,
      axis=False,
      bar_colors=["steel", "steel", "steel", "steel", "pale",
                  "navy", "navy", "navy", "navy", "cyan"],
      data_label="Total revenue filed by Thai companies", data_unit="฿ trillion",
      source=SRC_MAXDATA)

b.add("hbar_ranked", **S("1 · Why now"),
      title="The winners: five industries that took off after COVID",
      takeaway="Tourism reopening powers the top two. Wholesale & retail added the most baht: ฿33.5T in FY2025.",
      unit_note="Revenue growth FY2020 to FY2025, %",
      items=[
          {"label": "Accommodation & food", "value": 141, "display": "+141%"},
          {"label": "Arts & entertainment", "value": 140, "display": "+140%"},
          {"label": "Wholesale & retail", "value": 90, "display": "+90%"},
          {"label": "Other services", "value": 85, "display": "+85%"},
          {"label": "Mining & quarrying", "value": 74, "display": "+74%"},
      ],
      annotation={"text": "Gold-shop trading alone grew ฿1.9T → ฿8.85T",
                  "row": 2, "x": 6.55, "w": 2.95},
      source=SRC_MAXDATA)

b.add("hbar_ranked", **S("1 · Why now"),
      title="The laggards: where value is quietly leaking",
      takeaway="Weak segments are where motivated sellers live.",
      unit_note="Revenue decline FY2020 to FY2025, %. No whole sector shrank in nominal terms; these niches did",
      direction="left",
      items=[
          {"label": "Own-account investment", "value": 40, "display": "-40%"},
          {"label": "Direct-sales retail (MLM)", "value": 36, "display": "-36%"},
          {"label": "TV programme production", "value": 26, "display": "-26%"},
          {"label": "Rubber & plastics wholesale", "value": 19, "display": "-19%"},
          {"label": "Residential developers", "value": 18, "display": "-18%"},
      ],
      source=SRC_MAXDATA)

b.add("sector_matrix", **S("1 · Why now"),
      title="The reshuffle in the four sectors this room plays in",
      subtitle="Operating counts, recovery growth, and margin swings, straight from the registry",
      rows=[
          {"name": "Logistics", "icon": _ic("truck"), "count": "44,190",
           "note": "9,399 mid-market targets ฿10M-1B",
           "growth": 66, "margin_from": "-14%", "margin_to": "+7.5%"},
          {"name": "Manufacturing", "icon": _ic("factory"), "count": "109,063",
           "note": "฿21.1T revenue, Thailand's largest",
           "growth": 30, "margin_from": "4.3%", "margin_to": "4.3% steady"},
          {"name": "F&B & hospitality", "icon": _ic("food"), "count": "48,564",
           "note": "The steepest recovery arc",
           "growth": 141, "margin_from": "-22%", "margin_to": "+7.4%"},
          {"name": "Tech & information", "icon": _ic("chip"), "count": "29,027",
           "note": "Growth in profit, not just sales",
           "growth": 18, "margin_from": "5.9%", "margin_to": "7.5%"},
      ],
      callout="Fewer than 200 of 44,190 logistics companies exceed ฿1B revenue",
      source=SRC_MAXDATA)

b.add("bubble_chart_takeaways", **S("1 · Why now"),
      title="Inside logistics: growth and profit live in different places",
      bubbles=[
          {"label": "3PL / contract logistics", "x": 139, "y": 3.4,
           "size": 2.6, "group": "blue_light", "label_pos": "top"},
          {"label": "Courier", "x": 134, "y": 0.0, "size": 2.0,
           "group": "navy", "label_pos": "left"},
          {"label": "Air freight", "x": 129, "y": 5.0, "size": 1.0,
           "group": "blue_dark", "label_pos": "left"},
          {"label": "Airports", "x": 107, "y": 25.7, "size": 1.3,
           "group": "blue_royal", "label_pos": "right"},
          {"label": "Freight forwarding", "x": 42, "y": 3.1, "size": 3.5,
           "group": "blue_dark", "label_pos": "top"},
          {"label": "Generic trucking", "x": 16, "y": 3.3, "size": 4.1,
           "group": "navy", "label_pos": "top"},
      ],
      x_max=160, y_max=28,
      x_label="Revenue growth FY2020 to FY2025", x_unit="%",
      y_label="Net margin FY2025", y_unit="%",
      groups=(),
      size_label="FY2025 revenue",
      description="Six logistics sub-industries, verified year-by-year in the registry",
      takeaway_header="What it means",
      takeaways=["3PL more than doubled: the cleanest growth story in Thai logistics",
                 "Couriers grew 134% but ran at a loss for years: a price war, now consolidating",
                 "Generic trucking, the sector's biggest pool: +16% in five years",
                 "401 mid-market 3PL targets (฿10M-1B) sit under the leaders"],
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
           "note": "Fastest sub-industry in our screen: ฿8B to ฿32B, companies nearly doubled"},
          {"name": "Meat & poultry processing",
           "stat": "+126%",
           "note": "Revenue up every single year; margin turned positive in FY2024"},
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
           "right": "Listed corporates, mid-size firms, individuals"},
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
      description="Q1 2026: USD 7.9B across 57 deals, six times the prior quarter. One deal, True Corporation's 25% stake, was USD 3.9B",
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
      waffle_caption="111,304 of ~137,000 operating companies over 20 years old: "
                     "every board seat held by the pre-1984 generation",
      closing="These companies need successors. Increasingly, that means buyers.",
      source="Max Data registry analysis of 4.2M directorships, birth cohorts from national-ID structure, July 2026")

b.add("stat_band", **S("1 · Why now"),
      title="The demographic clock behind it is public record",
      stats=[
          {"glyph": "donut20", "value": "20%",
           "label": "of Thais are aged 60+, and rising",
           "context": "super-aged society by 2033"},
          {"glyph": "bar80", "value": "~80%",
           "label": "of Thai businesses are family-owned",
           "context": "฿30T combined value"},
          {"glyph": "onethree", "value": "1 in 3",
           "label": "family firms survive to generation 2",
           "context": "global benchmark"},
          {"glyph": "steps", "value": "3,000+",
           "label": "SME succession deals a year in Japan",
           "context": "already an industry there"},
      ],
      kicker="Japan is the preview of Thailand's next decade.",
      source="NESDC via Nation Thailand 2024; Grant Thornton Thailand; METI Japan via ORIX 2025")

b.add("quote_breather", **S("1 · Why now"),
      title="Why we started three years ago",
      quote="We watched a generation of Thai founders build companies their "
            "children didn't want to run. Someone had to build the bridge "
            "between those owners and the people ready to take over.",
      author="[Max · full name]",
      author_title="Founder, Max Solutions · placeholder quote, edit to taste",
      photo_label="Photo: Max",
      photo_path=f"{GEN}/scenes/quote-founder.png",
      source="Max Solutions")

# =====================================================================
# SECTION 2 · WHAT GOOD LOOKS LIKE
# =====================================================================

b.add("art_divider", art_path=f"{GEN}/dividers/div-2.png",
      section_number="02",
      section_title="What a good target looks like",
      subtitle="The lens our deal team applies before we ever talk price")

b.add("poll_slide", **S("2 · The lens"),
      title="Quick pulse check",
      question="You are shown a company for sale. What do you check first?",
      options=["The financial statements",
               "The owner and why they are selling",
               "The customers and recurring revenue",
               "The competition around it"],
      instruction="Vote now. We'll show you our order in a minute",
      ornament_path=f"{GEN}/backdrops/poll-corner.png")

b.add("chevron_flags", **S("2 · The lens"),
      title="The Green 5: what our analysts look for in every target",
      subtitle="The five signals our deal team scores before we ever talk price",
      flags=[
          {"keyword": "Synergy",
           "descriptor": "Multiplies what you already own"},
          {"keyword": "Potential",
           "descriptor": "Obvious headroom the owner never used"},
          {"keyword": "Management",
           "descriptor": "Runs without the owner in the room"},
          {"keyword": "Revenue quality",
           "descriptor": "Recurring customers, repeat contracts"},
          {"keyword": "Position",
           "descriptor": "A defensible niche, not a price war"},
      ],
      source="Max Solutions target evaluation framework")

b.add("keyline_panels", **S("2 · The lens"),
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

b.add("ledgers", **S("2 · The lens"),
      title="The two-books problem, and why diligence pays for itself",
      source="Max Solutions deal experience")

b.add("scorecard_slide", **S("2 · The lens"),
      title="Take this home: the 60-second target scorecard",
      subtitle="Screenshot this. Rate every target Weak or Strong per line; one Weak in the Red 3 kills the deal",
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
# SECTION 3 · WHERE DEALS COME FROM
# =====================================================================

b.add("art_divider", art_path=f"{GEN}/dividers/div-3.png",
      section_number="03",
      section_title="Where deals actually come from",
      subtitle="From two million registered companies to the handful worth your time")

b.add("taper_funnel", **S("3 · The access"),
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

b.add("ceilings", **S("3 · The access"),
      title="How everyone finds deals today, and where each channel stops",
      takeaway="Each channel works. Each has a ceiling you will hit.",
      channels=[
          {"name": "Word of mouth", "ceiling": "2-3 deals a year, if lucky",
           "note": "Friends, suppliers, golf partners. High trust, tiny reach"},
          {"name": "Online listings", "ceiling": "the good ones go fast",
           "note": "Marketplaces and brokers' sites. Wide reach, uneven quality"},
          {"name": "Intermediaries", "ceiling": "needs a clear mandate",
           "note": "Advisors and boutique firms. Screened deals, real process"},
      ],
      breakout="Breaking every ceiling at once is what the next four slides are about",
      source="Max Solutions market observation")

b.add("access_ladder", **S("3 · The access"),
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

b.add("screenshot_slide", **S("3 · The access"),
      title="Inside the community: where off-market deals surface",
      placeholder_label="Screenshot: DealFlow Facebook community",
      image_path="/home/user/mckinsey-pptx/webinar/assets/shots/dealflow-community.png",
      image_caption="Illustrative preview. The real group lives on Facebook: DealFlow by Max Solutions",
      kicker="Door 1 · Community",
      claim="Sellers test the water here before any listing exists",
      bullets=["Owners, buyers, and advisors in one room",
               "Free to join, and the fastest way to see deal flow"],
      overlap_stat={"value": "80,000+", "label": "community members"},
      layout_mode="right",
      source="DealFlow community, July 2026")

b.add("screenshot_slide", **S("3 · The access"),
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

b.add("metro_line", **S("3 · The access"),
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

b.add("screenshot_slide", **S("3 · The access"),
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
             {"value": "4.2M", "label": "directorships mapped"}],
      layout_mode="hero",
      source="Max Data platform, July 2026")

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
           "description": "Reach the owner, build trust, NDA"},
          {"name": "LOI",
           "description": "Letter of intent, sometimes a deposit"},
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

b.add("poll_slide", **S("5 · Proof"),
      title="One more pulse check",
      question="Which situation is closer to yours?",
      options=["I want to buy my first business",
               "My company should be acquiring",
               "I own a business that needs a successor",
               "I connect people and want the referral fee"],
      instruction="Vote now. The cases coming up cover all four",
      ornament_path=f"{GEN}/backdrops/poll-corner.png")

b.add("profile_cards", **S("5 · Proof"),
      title="The supply is real: companies like these are in the registry right now",
      headline_stat="14,139",
      headline_label="companies match the succession profile: 20+ years old, "
                     "฿50M-1B revenue, profitable, all-pre-1984 board",
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
                 "Our deal team holds the full list.",
      source=SRC_MAXDATA)

b.add("stat_hero_split", **S("5 · Proof"),
      title="The search-fund wave: individuals now buy companies, profitably",
      stat="35.1%",
      stat_label="aggregate IRR across search funds tracked by Stanford GSB",
      rows=[
          {"value": "4.5x", "label": "aggregate return on invested capital"},
          {"value": "94", "label": "funds launched in 2023, a record"},
          {"value": "320", "label": "international funds by end-2023"},
          {"value": "$11.7M", "label": "median international deal price"},
      ],
      kicker="The same playbook, applied to Thailand's succession wave, is what tonight is about.",
      source="Stanford GSB Search Fund Study 2024 (Case E-870); IESE International Search Funds 2024")

b.add("case_slide", **S("5 · Proof"),
      title="Case one: the fire-safety distributor",
      case_name="Project FireGuard · closed Q1 2026",
      sector_chip="FIRE SAFETY / TRADING",
      situation=["Husband-and-wife owners near retirement, no successor",
                 "฿100M revenue, ฿30M EBITDA, decades of relationships",
                 "One of only 112 fire-safety companies tracked in Max Data"],
      outcome=["New owner modernized systems within months",
               "Employees stayed, morale improved",
               "Owners exited proud, business on an IPO-track plan"],
      bridge_stat="< 6 months to visible turnaround",
      kpis=[{"value": "฿100M", "label": "Revenue at deal"},
            {"value": "฿30M", "label": "EBITDA at deal"},
            {"value": "112", "label": "fire-safety companies tracked"}],
      photo_label=None,
      photo_path=f"{GEN}/scenes/case-fire.png",
      source="Max Solutions deal team, 2026. Figures approximate to protect the parties")

b.add("case_slide", **S("5 · Proof"),
      title="Case two: the pizza-oven supplier",
      case_name="Project Pizza Oven · closed Q1 2026",
      sector_chip="F&B EQUIPMENT / SUPPLY",
      situation=["Italian owner couple returning home after years in Thailand",
                 "Supplier to major Thai restaurant chains",
                 "Top-3 Google ranking in its niche, loyal recurring customers"],
      outcome=["Foreign buyer acquired end to end through our process",
               "Smooth handover, customers retained",
               "Now on its way to doubling turnover"],
      bridge_stat="2x turnover trajectory",
      kpis=[{"value": "Top 3", "label": "Google rank in its niche"},
            {"value": "94", "label": "kitchen-equipment suppliers tracked"}],
      photo_path=f"{GEN}/scenes/case-oven.png",
      source="Max Solutions deal team, 2026. Figures approximate to protect the parties")

_sl_caus = b.add("causality_band", **S("5 · Proof"),
      title="Why small deals turn around so fast",
      subtitle="The pattern behind both cases, and most of our closed deals",
      blocks=[
          {"label": "Honest sellers",
           "support": "Price reflects the exit need, not a bidding war"},
          {"label": "Inefficiency is the upside",
           "support": "No ERP, no marketing, no pricing discipline"},
          {"label": "Capable buyers",
           "support": "Corporate resources or sharp operators"},
      ],
      result="Margins move in months, not years. SME payback: 3-8 years vs 10+ on mega-deals",
      source="Max Solutions deal experience across 150+ SME mandates per year")
_pic(_sl_caus, f"{GEN}/scenes/succession.png", 4.97, 4.85, 3.4, 1.35)

b.add("dimension_table", **S("5 · Proof"),
      title="Sourcing then vs now: what the data layer changes",
      left_header="The old way", right_header="With Max Data",
      check_right=True,
      rows=[
          {"dim": "Sourcing",
           "left": "Call owners one by one, hope",
           "right": "Screen 993k companies against your thesis"},
          {"dim": "Financials",
           "left": "Guesswork until diligence",
           "right": "10 years of statements before the first call"},
          {"dim": "Speed",
           "left": "Months to build a list",
           "right": "A vetted shortlist in an afternoon"},
          {"dim": "Your edge",
           "left": "Who you happen to know",
           "right": "Seeing what others cannot"},
      ],
      source="Max Data platform, July 2026")

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
          {"takeaway": "A lens to judge any deal", "icon": _ic("magnifier"),
           "line": "The Green 5, the Red 3, and the 60-second scorecard"},
          {"takeaway": "A way in", "icon": _ic("door"),
           "line": "Three doors, the Max Data layer, and the 6-step path"},
      ],
      conclusion="The buyers who win the succession wave start looking before everyone else.")

b.add("service_spectrum", **S("6 · Next step"),
      title="Research and sourcing from here: three lanes, one call",
      axis_left="We carry the work",
      axis_right="You run it yourself",
      lanes=[
          {"icon": _ic("handshake-white"), "kicker": "Full-mandate advisory",
           "name": "Max Solutions",
           "line": "We do everything with you: research, sourcing, "
                   "negotiation, close.",
           "bullets": ["A dedicated deal team on your mandate end to end",
                       "You decide, we run the machine"],
           "stat": "150+", "stat_label": "SME mandates a year"},
          {"icon": _ic("storefront-white"), "kicker": "Screened marketplace",
           "name": "DealFlow Market",
           "line": "You screen the deals yourself, we keep them real.",
           "bullets": ["Live, verified listings with real numbers",
                       "Move at your own shortlist speed"],
           "stat": "100+", "stat_label": "live deals across 15 industries"},
          {"icon": _ic("chip"), "kicker": "Self-serve platform",
           "name": "Max Data",
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
      bottom_actions=["Type 1 (buying), 2 (selling), or 3 (referring) in the chat, and our team follows up tomorrow",
                      "Or scan the QR / add our LINE official account now"],
      qr_label="QR: booking page")

b.add("qa_slide",
      bg_path=f"{GEN}/backdrops/qa-waves.png",
      line="Ask us anything. We stay until your questions run out.",
      qr_caption="Scan to book · or type 1 / 2 / 3 in the chat")

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
           "body": "Cited on their slides: KPMG Thailand quarterly M&A reports, "
                   "NESDC demographic data, Grant Thornton Thailand, METI Japan, and "
                   "Stanford GSB search fund research. Case figures are rounded and "
                   "lightly disguised to protect client confidentiality."},
      ],
      source="Max Solutions research team, July 2026")

# =====================================================================
# SPEAKER NOTES — timing budgets + delivery cues from the planning session
# =====================================================================

NOTES = {
    1: "Doors open 20:25, greet people by name in chat. Start 20:32 sharp. [T+0:00]",
    2: "Launch poll 1 immediately, read the options aloud. Tease: 'we'll tailor the examples to tonight's mix.' About 2 minutes. [T+0:02]",
    3: "One breath per stop. The promise: 'by the end you'll know where deals come from, how to judge one, and your first step.' [T+0:05]",
    4: "Tie back to poll 1: 'here's your row.' One line per seat, under a minute total. Plant the tailored-session offer casually, don't sell it. [T+0:06]",
    5: "30 seconds each. Max: keep the founder story to ONE line here, the full beat comes at slide 17. [T+0:06]",
    6: "Section 1 · WHY NOW. Budget 14 minutes. [T+0:07]",
    7: "Big line: 'COVID knocked four trillion baht off the market. The recovery ran at nine percent a year.'",
    8: "Pause on the gold factoid. Chat prompt: 'type your industry in the chat.'",
    9: "Frame positively: weak segments are where motivated sellers live. Don't dwell, next slide is the room's own sectors.",
    10: "This room's four sectors. Logistics people: 9,399 mid-market targets is YOUR number. Move briskly, the drill-down is next.",
    11: "The nuance slide: growth and profit are different places in logistics. 3PL is the clean story; couriers grew broke. 90 seconds.",
    12: "One hero, three runners. Pet food is the 'growth AND margin' proof. Mention the artifact-stripping line: that's analyst rigor.",
    13: "90 seconds max. The room knows this distinction, it's a bridge slide.",
    14: "Key claim: SME deals pay back in 3-8 years vs 10+ for mega-deals. Japan's 16x growth is the evidence the wave is real.",
    15: "Credibility beat: this is KPMG data published last month. Note: headlines are big-cap, our layer is the SME market below them.",
    16: "SLOW DOWN. This is the screenshot slide. Read the stat twice. 'No one else in Thailand can compute this number.'",
    17: "Public-record confirmation of our registry finding. The Japan tile is the preview of Thailand's next decade.",
    18: "Founding story, 45 seconds max. Read the quote, add one personal line, move on. [T+0:21]",
    19: "Section 2 · WHAT GOOD LOOKS LIKE. Budget 14 minutes. [T+0:21]",
    20: "Launch poll 2. Callback comes at the Red 3 slide: most people vote financials, and financials is where deals die.",
    21: "One concrete example per flag. Ask chat: 'which of these do you weight most?'",
    22: "One short war story per risk if time allows. Two-books gets its own slide next.",
    23: "Sensitive topic, phrase carefully: 'a common practice, and fixable in diligence', never accusatory.",
    24: "Tell viewers to screenshot this one. It's also in the replay materials.",
    25: "Section 3 · WHERE DEALS COME FROM. Budget 12 minutes. [T+0:35]",
    26: "Walk the cascade slowly. 'Only 1 in 18 companies sits in the core band' lands well.",
    27: "Two minutes max, this is setup for the ladder.",
    28: "Our ecosystem. Not a pitch: the first two doors are free to open tonight.",
    29: "Screenshot shown is an illustrative preview; show the real Facebook group live if the connection allows.",
    30: "This is the REAL dealflowmarket.com, captured this week. Pick one listing that matches the poll-1 mix, talk 30 seconds.",
    31: "Position Max Data as the research layer under every stage, not another listing site.",
    32: "DEMO MOMENT. Recorded fallback ready. Filter: Logistics, ฿10M-1B revenue → 9,399 companies on screen. [T+0:45]",
    33: "Section 4 · HOW A DEAL RUNS. Budget 8 minutes. [T+0:47]",
    34: "Walk the six steps in about 4 minutes. Add the LOI + deposit nuance from real deals.",
    35: "Sellers: 60 seconds. Full sell-side detail is in the appendix and on the website.",
    36: "Section 5 · PROOF. Budget 13 minutes. [T+0:55]",
    37: "Launch poll 3. Use the result to decide which case to emphasize.",
    38: "The pipeline slide. Read one profile aloud slowly. 'Fourteen thousand of these.' Then: 'so who's buying them?'",
    39: "Let 35.1% breathe. Then: 'this playbook is arriving in Thailand.'",
    40: "Tell it as a story: the couple, no successor, the sharp buyer, the New Year visit where everyone was happy.",
    41: "Contrast case: foreign buyer, end-to-end process, now doubling turnover.",
    42: "Generalize the pattern: succession + inefficiency + capable buyer.",
    43: "Callback to the demo: 'your edge is seeing what others cannot.' [T+1:06]",
    44: "Section 6 · YOUR NEXT STEP. Budget 6 minutes. [T+1:08]",
    45: "Three take-homes. Second screenshot moment.",
    46: "Lanes slide: sweep once left to right, 'from us doing everything to you running it all yourself.' Then the hook, unhurried: 'first five bookings get a customized session.' [T+1:03]",
    47: "THE ask. Type 1 (buy) / 2 (sell) / 3 (refer), scan QR, or add LINE. Say it twice. QR stays up through Q&A. [T+1:10]",
    48: "About 20 minutes. Seed questions ready: 'what multiples do Thai SMEs sell for?', 'can foreigners buy?', 'how long does a deal take?'",
    49: "Personal beat. Read the groups briskly and warmly, then slow right down for the last line. Look at the camera for Pim's line.",
    50: "Mention the replay and community link land in tomorrow's follow-up email.",
    51: "Appendix: replay material. Skip live unless Q&A pulls them up.",
}

for _i, _slide in enumerate(b.prs.slides, 1):
    note = NOTES.get(_i)
    if note:
        _slide.notes_slide.notes_text_frame.text = note

editorial_slides.flatten_all_shadows(b.prs)

out = "/home/user/mckinsey-pptx/webinar/output/ma-webinar-deck.pptx"
b.save(out)
print(f"saved {out} ({len(b.prs.slides._sldIdLst)} slides)")
