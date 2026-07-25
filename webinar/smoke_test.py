"""Quick smoke test of the custom Max slide templates."""
import sys
sys.path.insert(0, "/home/user/mckinsey-pptx")
sys.path.insert(0, "/home/user/mckinsey-pptx/webinar")

from mckinsey_pptx import PresentationBuilder
from max_theme import MAX_THEME
import max_slides  # noqa: F401  (registers the custom templates)

b = PresentationBuilder(theme=MAX_THEME, default_section_marker="M&A Webinar")

b.add("max_cover",
      title="Thailand's Hidden M&A Opportunity",
      subtitle="Where deals come from, how to judge them, and the fastest way in",
      event_line="Max Solutions live webinar",
      date="Sunday 26 July 2026, 20:30 (ICT)")

b.add("speaker_slide", speakers=[
    {"name": "[Max]", "role": "Founder, Max Solutions",
     "bullets": ["150+ SMEs advised per year", "15 industries covered"],
     "photo_label": "Photo: Max"},
    {"name": "[Wipin]", "role": "Head of M&A",
     "bullets": ["Deal execution lead", "100+ transactions"],
     "photo_label": "Photo: Wipin"},
])

b.add("poll_slide",
      question="Which best describes you tonight?",
      options=["Buyer or investor", "Owner thinking about selling",
               "Advisor or connector", "Just exploring M&A"],
      instruction="Vote now in the poll panel")

b.add("screenshot_slide",
      title="Max Data in 90 seconds",
      placeholder_label="Screenshot: Max Data screener",
      placeholder_note="Filter: Logistics, revenue ฿100M-1B",
      bullets=["Filter 890k operating companies in seconds",
               "See 10 years of financials on any company",
               "Reach the actual decision maker"],
      stats=[{"value": "1.99M", "label": "companies tracked"},
             {"value": "9.8M", "label": "financial statements"}],
      source="Max Data platform, July 2026")

b.add("case_slide",
      title="Case: the fire-safety distributor",
      case_name="Project FireGuard",
      sector_chip="FIRE SAFETY / TRADING",
      situation=["Husband-and-wife owners, no successor",
                 "฿100M revenue, ฿30M EBITDA",
                 "30 years of relationships, zero digitization"],
      outcome=["New owner installed ERP within months",
               "Employees retained, morale up",
               "On track to consider IPO within 5 years"],
      kpis=[{"value": "฿100M", "label": "Revenue at deal"},
            {"value": "฿30M", "label": "EBITDA"},
            {"value": "< 6 mo", "label": "To turnaround"}],
      source="Max Solutions deal team, 2026")

b.add("access_ladder",
      title="Where the deals actually are",
      steps=[
          {"name": "Community", "stat": "80,000+ members",
           "bullets": ["DealFlow Facebook community",
                       "Off-market chatter surfaces here first"]},
          {"name": "Marketplace", "stat": "100+ live deals",
           "bullets": ["DealFlow Market listings",
                       "Screened sellers across 15 industries"]},
          {"name": "Advisory", "stat": "150 SMEs / year",
           "bullets": ["Full-mandate M&A advisory",
                       "End-to-end deal execution"]},
      ])

b.add("cta_slide",
      title="Your next step",
      paths=[
          {"who": "BUYERS", "action": "Book an opportunity scan",
           "detail": "We map live targets against your thesis."},
          {"who": "OWNERS", "action": "Get a confidential valuation",
           "detail": "Know what your business is worth today."},
          {"who": "EVERYONE", "action": "Join the community",
           "detail": "80,000+ members trading deal flow daily."},
      ],
      bottom_actions=["Type 1 in the chat and our team will contact you",
                      "Or add our LINE official account today"],
      qr_label="QR: booking page")

b.add("thank_you",
      lines=["Max Solutions · DealFlow Market · Max Data"],
      contact_placeholder="Contact block: LINE ID, email, phone")

out = "/home/user/mckinsey-pptx/webinar/output/smoke.pptx"
b.save(out)
print("saved", out)
