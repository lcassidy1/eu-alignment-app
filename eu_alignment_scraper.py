"""
EU Delegation UN New York – Alignment Statement Scraper
Scrapes all known UN NY statement URLs, detects alignment clauses,
and produces an Excel report for the UN80 session (Sept 8, 2025+).

Usage:
    pip install pandas openpyxl
    python3 eu_alignment_scraper.py

To refresh the URL list from the EEAS sitemap (e.g. at end of each month):
    python3 eu_alignment_scraper.py --refresh-urls
"""

import re
import sys
import time
import urllib.request
from datetime import date, datetime
from pathlib import Path

import pandas as pd

# ── Configuration ─────────────────────────────────────────────────────────────

UN80_START  = date(2025, 9, 8)
OUTPUT_FILE = Path(__file__).parent / "eu_alignment_stats.xlsx"
URL_CACHE   = Path(__file__).parent / "statement_urls.txt"
SITEMAP_PAGES = 21

TRACKED_COUNTRIES = {
    "Turkey":                 r"\bTurkey\b|Türkiye",
    "North Macedonia":        r"North Macedonia",
    "Montenegro":             r"\bMontenegro\b",
    "Serbia":                 r"\bSerbia\b",
    "Albania":                r"\bAlbania\b",
    "Ukraine":                r"\bUkraine\b",
    "Moldova":                r"Republic of Moldova|\bMoldova\b",
    "Bosnia and Herzegovina": r"Bosnia and Herzegovina",
    "Georgia":                r"\bGeorgia\b",
    "Iceland":                r"\bIceland\b",
    "Liechtenstein":          r"\bLiechtenstein\b",
    "Norway":                 r"\bNorway\b",
    "Armenia":                r"\bArmenia\b",
    "Azerbaijan":             r"\bAzerbaijan\b",
    "Andorra":                r"\bAndorra\b",
    "Monaco":                 r"\bMonaco\b",
    "San Marino":             r"San Marino",
    "UK":                     r"United Kingdom|\bUK\b",
}

# Matches any sentence where countries "align/associate themselves"
# Catches: "align themselves with this statement"
#          "associate themselves with this statement"
#          "continue to align themselves with the above statement"
#          "align themselves with the position expressed"
ALIGNMENT_PATTERN = re.compile(
    r"\b(align|associate)\b[^.!?]{0,60}\bthemselves\b",
    re.IGNORECASE,
)

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "en-US,en;q=0.9",
}

SKIP_PATTERNS = [
    "/global-gateway", "/about-ambassador", "/who-we-are",
    "/european-union-and-united-nations", "/meet-eu-youth",
    "/newsletter", "/vacancy", "/job-", "/who-rules-world",
    "/reforming-impact", "/time-accountability", "/call-action",
]

# URLs known to exist but not indexed in the EEAS sitemap
EXTRA_URLS = [
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-security-council-future-peace-operations_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-security-council-maintenance-international-peace-and-security-poland_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-ad-hoc-working-group-un80-first-informal-meeting_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-explanation-vote-%E2%80%93-un-general-assembly-1st-committee-vote-cluster-iii-outer-space-no-first_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-explanation-votes-%E2%80%93-un-general-assembly-3rd-committee-vote-amendment-l-56-resolution_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-briefing-secretary-general-un80-initiative_en",
]

# ── URL list (from sitemap scan, June 2026) ───────────────────────────────────

STATEMENT_URLS = [
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-interactive-multi-stakeholder-hearing-part-preparatory-process-high-level-meeting_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/un-general-assembly-agenda-item-31b-strengthening-role-mediation-peaceful-settlement-disputes_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-high-level-debate-illicit-trafficking-wildlife_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-security-council-ukraine-13_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-joint-meeting-executive-board-undpunfpaunops-unicef-un-women-wfp_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-world-environment-day-2026_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-ecosoc-operational-activities-segment-dialogue-unds-executive-heads_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-ecosoc-operational-activities-segment-dialogue-un-development-system-funding_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-ecosoc-operational-activities-segment-dialogue-host-governments-rcs-and-uncts_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-ecosoc-operational-activities-segment-improving-unds-governance-oversight-and_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-9th-biennial-meeting-states-small-arms-and-light-weapons-implementation-programme_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-9th-biennial-meeting-states-small-arms-and-light-weapons-implementation-programme-0_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-protracted-conflicts-guam-area-and-their-implications-0_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-meeting-open-ended-technical-expert-group-developments-small-arms-and-light-weapons_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-explanation-vote-un-general-assembly-draft-resolution-global-health-and-foreign-policy_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-security-council-arria-formula-meeting-west-bank-including-east-jerusalem_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/meet-eu-youth-delegates-un-81st-general-assembly-session_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-ecosoc-operational-activities-segment-high-level-dialogue-secretary-general-his-report_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-ecosoc-operational-activities-segment-dialogue-deputy-secretary-general-unsdg-chair%E2%80%99s_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-criteria-and-modalities-review-existing-stock-ga-mandates_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-undp-annual-meeting-rule-law-and-human-rights_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-un-security-council-briefing-threats-international-peace-and-security-0_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-un-security-council-maintenance-peace-and-security-ukraine-5_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-un-security-council-high-level-open-debate-upholding-purposes-and-principles-un-charter_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/children-crisis-spotlight-underfunded-humanitarian-emergencies_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/un-information-briefing-un-secretary-general_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-11th-review-conference-parties-treaty-non-proliferation-nuclear-weapons-npt-closing_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-un-security-council-open-debate-protection-civilians_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/joint-statement-41-members-inter-regional-task-force-moratorium-use-death-penalty_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-un-general-assembly-resolution-advisory-opinion-international-court-justice-obligations_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-un-general-assembly-ahwg-draft-templates-and-clauses-mandate-implementation-review_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-un-general-assembly-high-level-panel-theme-%E2%80%9C-importance-complying-venice-principles_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-un-security-council-ukraine_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-general-assembly-fifth-committee-second-resumed-session-improving-financial-situation_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-un-general-assembly-informal-meeting-general-committee_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-ecosoc-special-meeting-safeguarding-energy-and-supply-flows-supporting-global_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/interactive-multi-stakeholder-hearing-part-preparatory-process-2026-high-level-meeting-hivaids_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-interactive-dialogues-candidates-position-president-81st-session_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-global-dialogue-ai-governance-informal-thematic-deep-dive-session_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-security-council-debate-bosnia-and-herzegovina-1_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-committee-information-closing-plenary_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-explanation-position-%E2%80%93-un-general-assembly-adoption-draft-resolution-organization-2026-high-level_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-joint-un-general-assembly-and-economic-and-social-council-special-meeting-advancing_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-security-council-arria-formula-meeting-protecting-medical-care-conflict-amid_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-11th-review-conference-parties-treaty-non-proliferation-nuclear-weapons-npt-main-0_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-5th-committee-second-resumed-session-organisation-work-1_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-11th-review-conference-parties-treaty-non-proliferation-nuclear-weapons-npt-main-1_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-11th-review-conference-parties-treaty-non-proliferation-nuclear-weapons-npt-main_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-monthly-information-briefing-un80-initiative-ws3_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-security-council-open-debate-situation-middle-east-including-palestinian-question-2_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-peacebuilding-commission-ambassadorial-level-meeting-gambia_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-25th-session-permanent-forum-indigenous-issues-agenda-item-4-discussion-six-mandated_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-25th-permanent-forum-indigenous-issues-human-rights-dialogue-special-rapporteur_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-explanation-position-un-financing-development-forum-outcome-document_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-security-council-safety-and-protection-maritime-waterways_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-48th-committee-information-opening-session_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-11th-review-conference-parties-treaty-non-proliferation-nuclear-weapons-npt-general_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-expert-mechanism-right-development-general-debate_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-un-general-assembly-chornobyl-commemoration_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-un-general-assembly-interactive-dialogue-pact-future-and-un80-initiative-commitments_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-25th-permanent-forum-indigenous-issues-ensuring-indigenous-peoples%E2%80%99-health-including_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-interactive-dialogues-candidates-position-un-secretary-general_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-permanent-forum-indigenous-issues-indigenous-peoples-and-climate-change_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-security-council-ukraine-12_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-59th-commission-population-development-closing-statement_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-59th-commission-population-development-opening-session_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-second-informal-consultation-hlpf-ministerial-declaration_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/who-rules-world-human-rights-who-needs-them-spoiler-everyone_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-use-veto-maritime-security-strait-hormuz_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-adoption-draft-resolution-role-diamonds-fuelling-conflict-1_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-48th-session-un-committee-information-informal-briefing-usg-department-global_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-prevention-armed-conflict_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/who-rules-world-ep-46-art-war-and-advocacy-ramikas-journey-and-life-women-afghanistan_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-security-concil-briefing-unmik-kosovo_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-further-intergovernmental-consultation-declaration-sea-level-rise_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-informal-consultation-global-dialogue-ai-governance_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-disarmament-commission-general-exchange-views_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-briefing-un80-initiative-workstream-3_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-ecosoc-recalibrating-resident-coordinator-system_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-informal-ad-hoc-working-group-un80-initiative-final-meeting_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-ecosoc-special-meeting-credit-ratings_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-event-women-breaking-barriers_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-5th-committee-first-resumed-session-closing-statement_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-explanation-vote-%E2%80%93-un-general-assembly-action-a80l48-declaration-trafficking-enslaved-africans_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-high-level-event-commemorate-international-day-combat_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-international-day-remembrance-victims-slavery-and-transatlantic_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-international-day-forests-2026_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-international-day-elimination-racial-discrimination-0_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-high-level-event-world-water-day-and-world-day-glaciers_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-security-council-ukraine-11_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/statement-23-eu-member-states-%E2%80%93-un-general-assembly-zero-draft-imrf-progress-declaration_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-informal-consultation-financing-development-forum-outcome_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-briefing-unep-executive-director-andersen_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-peacebuilding-commission-2nd-annual-interactive-dialogue-peacebuilding-fund_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-5th-committee-first-resumed-session-special-political-missions_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-briefing-un80-reform-and-humanitarian-compact_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-un80-meeting-rev1-resolution_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-first-informal-consultation-2026-high-level-political-forum_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/european-union-70th-un-commission-status-women-rights-justice-action_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/arbitrary-detentions-and-enforced-disappearances-ukrainian-women-journalists-russia_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/csw70-side-event-automating-justice-can-artificial-intelligence-increase-women%E2%80%99s-and-girl%E2%80%99s-access_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/csw70-side-event-strengthening-access-justice-all-women-and-girls-what-more-can-we-do-protect-women_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/csw-70-high-level-event-commemorating-5th-anniversary-group-friends-elimination-violence-against_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/csw70-side-event-intergenerational-engagement-advancing-young-womens-leadership_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/csw70-side-event-preventing-and-combating-all-forms-cyber-violence-against-girls_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/csw70-side-event-afghan-women-economic-actors-amid-crisis-and-restrictions_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/statement-behalf-25-eu-member-states-%E2%80%93-un-general-assembly-informal-meeting-imrf-progress_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-revitalization-ga_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/european-union-69th-un-commission-status-women-beijing30_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-informal-meeting-general-committee_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-9th-review-un-global-counter-terrorism-strategy-briefing-member-states_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/statement-behalf-25-eu-member-states-%E2%80%93-un-general-assembly-secretary-general%E2%80%99s-report-global-compact_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-pact-future-implementation_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-security-council-arria-formula-meeting-red-hand-day-2026-safe-education-prevent_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-2nd-committee-first-informal-dialogue-revitalization-second_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-5th-committee-first-resumed-session-human-resources-management_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-11th-emergency-special-session-ukraine_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-security-council-ukraine-10_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/80th-session-general-assembly-fifth-committee-first-resumed-session_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-un-security-council-meeting-sudan_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/joint-press-statement-latest-israeli-decisions-regarding-occupied-west-bank_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-special-committee-charter-and-strengthening-role-organization_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-special-committee-peacekeeping-operations-c34-general-debate_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-executive-board-un-women-opening-regular-session_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-appointment-independent-international-scientific-panel-artificial_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-briefing-2026-financing-development-forum-outcome_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-unicef-executive-board-update-unicef-humanitarian-action_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/united-nations-human-rights-fora-council-approves-eu-priorities-2026_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-ecosoc-first-informal-consultation-ecosoc-%E2%80%93-hlpf-review_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-unicef-executive-board-plan-global-evaluations-2026-29_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-working-group-ga-revitalization-role-and-authority-general-assembly_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-unicef-executive-board-opening-statement-3_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-executive-board-undpunfpaunops-regular-session-statement-unops-executive-director_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-executive-board-undp-unfpa-and-unops-unfpa-segment_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-security-council-threats-international-peace-and-security-caused-terrorist-acts_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-joint-briefing-presidents-general-assembly-and-economic-and_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-security-council-arria-formula-upholding-sanctity-treaties-maintenance_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-64th-session-commission-social-development-general-discussion_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-executive-board-undpunfpaunops-regular-session-interactive-dialogue-undp-administrator_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-executive-board-undpunfpaunops-regular-session-un80-initiative_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-peacebuilding-commission-20th-session-organizational-committee_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-security-council-open-debate-middle-east-including-palestinian-questions_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-ecosoc-2026-ecosoc-coordination-segment_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-ecosoc-2026-ecosoc-partnership-forum_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-security-council-open-debate-rule-law-pathways-reinvigorating-peace-justice-and_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-peacebuilding-commission-ambassadorial-meeting-national-efforts-prevention-and_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-ecosoc-commemoration-ecosoc-80_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-un-general-assembly-briefing-president-ga-priorities-resumed-part-80th-session_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-working-group-un-conference-plenipotentiaries-prevention-and-punishment-crimes_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-arria-formula-meeting-advancing-new-paradigms-peacebuilding-fortifying-inclusive-and_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/european-commission-announces-%E2%82%AC19-billion-humanitarian-aid-budget-2026_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-informal-ad-hoc-working-group-un80-response-zero-draft_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-un-security-council-maintenance-peace-and-security-ukraine-4_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-un-general-assembly-briefing-secretary-general-his-priorities-2026_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-un-general-assembly-fifth-committee-main-session-formal-closing-meeting_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/joint-statement-high-representative-kallas-and-commissioners-lahbib-and-%C5%A1uica-registration_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-special-committee-charter-united-nations-and-strengthening-role-organization-0_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-adoption-programme-action-landlocked-developing-countries-decade_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-un-general-assembly-special-committee-peacekeeping-operations-c34-general-debate_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-special-committee-charter-peaceful-settlement-disputes_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-informal-briefing-unccc-cop-30_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-explanation-vote-%E2%80%93-un-general-assembly-adoption-draft-resolution-un-decade-afforestation-and_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-international-impartial-and-independent-mechanism-syria_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-special-solemn-meeting-commemoration-all-victims-second-world-war_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-high-level-meeting-recommendations-open-ended-working-group-ageing_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-informal-interactive-dialogue-eliminating-child-labour_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-high-level-committee-south-south-cooperation_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-first-preparatory-meeting-wsis20-review_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-joint-meeting-ecosoc-and-peacebuilding-commission-building-and-sustaining-peace-haiti_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-un-security-council-high-level-open-debate-%E2%80%9Cpoverty-underdevelopment-and-conflict_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-ambassadorial-level-consultation-un80-initiative-context-reconfiguration-un%E2%80%99s-counter_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-security-council-open-debate-maritime-security_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-general-assembly-debate-global-health-and-foreign-policy_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-united-nations-fourteenth-article-xiv-conference-support-entry-force-comprehensive_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-5th-committee-organisation-work-0_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-6th-committee-rule-law-0_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-6th-committee-criminal-accountability-un-officials-and-experts-1_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-6th-committee-crimes-against-humanity-2_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-6th-committee-protection-persons-event-disasters-0_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-report-secretary-general-un80-initiative_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-plenary-new-partnership-africa%E2%80%99s-development-agenda-2063_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-6th-committee-report-special-committee-un-charter-and_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-5th-committee-programme-plan-and-programme-budget-2026_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-6th-committee-report-un-commission-international-trade-law-1_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-6th-committee-international-residual-mechanism-criminal-tribunals-2_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assemby-5th-committee-improving-financial-situation-united-nations_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-6th-committee-universal-jurisdiction_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-6th-committee-administration-justice-un-0_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-5th-committee-special-political-missions-1_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-open-briefing-un-security-council-1540-committee_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-report-international-court-justice-1_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-informal-ad-hoc-working-group-un80-mandate-implementation_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-general-statement-%E2%80%93-un-general-assembly-2nd-committee-operational-activities-development-un_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-2nd-committee-vote-care-economy-resolution_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-explanation-position-%E2%80%93-un-general-assembly-vote-amendment-resolution-care-economy_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-un-general-assembly-fifth-committee-items-136-and-148_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-oceans-and-law-sea-2_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-informal-meeting-un80-initiative-workstream-2_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-security-council-briefing-non-proliferation_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-un-general-assembly-second-committee-agenda-item-18f-implementation-convention_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-un-general-assembly-second-committee-agenda-18-sustainable-mountain-development_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-un-general-assembly-second-committee-agenda-item-18a-promoting-sustainable-consumption_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-remarks-ambassador-samson-roundtable-information-integrity-evolving-ai-landscape_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-high-level-meeting-implementation-outcomes-world-summit_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-explanation-position-%E2%80%93-un-general-assembly-political-declaration-non-communicable-diseases-and_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-security-council-leadership-peace_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-peacebuilding-commission-ambassadorial-level-meeting-peaceful-settlement-border_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/roundtable-information-integrity-evolving-ai-landscape_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-humanitarian-general-debate_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/human-rights-day-statement-high-representative-behalf-european-union_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-high-level-meeting-10th-anniversary-international-day_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-security-council-maintenance-peace-and-security-ukraine-15_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-un-general-assembly-informal-meeting-plenary-commemorate-and-promote-international-day_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-un-general-assembly-second-committee-agenda-item-18f-sustainable-mountain-development_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-un-general-assembly-11th-emergency-special-session-ukraine-0_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-ecosoc-meeting-between-united-nations-resident-coordinators-and-member-states_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-2nd-committee-information-and-communications-technologies_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-general-statement-%E2%80%93-un-general-assembly-2nd-committee-development-cooperation-middle-income_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-explanation-position-%E2%80%93-un-general-assembly-2nd-committee-science-technology-and-innovation_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-explanation-vote-%E2%80%93-un-general-assembly-2nd-committee-promotion-inclusive-and-effective_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-explanation-vote-%E2%80%93-un-general-assembly-2nd-committee-cooperation-combat-illicit-financial-flows_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-explanation-vote-%E2%80%93-un-general-assembly-2nd-committee-eradication-rural-poverty-implement-2030_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statements-%E2%80%93-un-general-assembly-2nd-committee-international-trade-and-development_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-general-statement-%E2%80%93-un-general-assembly-2nd-committee-culture-and-sustainable-development_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-un-global-plan-action-combat-trafficking-persons_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-2nd-committee-south-south-cooperation-development_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-general-statement-%E2%80%93-un-general-assembly-2nd-committee-un-convention-combat-desertification-1_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-2nd-committee-implementation-third-un-decade-eradication-poverty_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-explanation-vote-%E2%80%93-un-general-assembly-2nd-committee-after-action-resolution-protection-global_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-general-statement-%E2%80%93-un-general-assembly-2nd-committee-follow-antigua-and-barbuda-agenda-small_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-building-peaceful-and-better-world-through-sport-and-olympic_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-security-council-arria-formula-meeting-protection-seafarers_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/joint-statement-international-day-elimination-violence-against-women_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-explanation-vote-%E2%80%93-un-general-assembly-3rd-committee-amendments-rights-child-resolution_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-explanation-vote-%E2%80%93-un-general-assembly-3rd-committee-amendment-resolution-strengthening-un-crime_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-explanation-vote-%E2%80%93-un-general-assembly-3rd-committee-amendments-rights-child-resolution-0_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-explanation-vote-%E2%80%93-un-general-assembly-3rd-committee-after-vote-draft-resolution-human-rights-and_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-explanation-vote-%E2%80%93-un-general-assembly-vote-resolution-torture_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-explanation-position-%E2%80%93-un-general-assembly-3rd-committee-after-adoption-resolution-30th_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-explanation-votes-%E2%80%93-un-general-assembly-rights-persons-disabilities_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-explanation-vote-%E2%80%93-un-general-assembly-after-adoption-resolution-implementation-outcome-world_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-explanation-votes-%E2%80%93-un-general-assembly-3rd-committee-vote-amendment-l-56-resolution_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-explanation-vote-%E2%80%93-un-general-assembly-3rd-committee-after-vote-draft-resolution-right_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-explanation-vote-%E2%80%93-un-general-assembly-vote-resolution-human-rights-temporarily-occupied_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-explanation-vote-%E2%80%93-un-general-assembly-vote-human-rights-iran_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-explanation-vote-%E2%80%93-un-general-assembly-after-adoption-resolution-human-rights-syria_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-security-council-ukraine-9_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/un-general-assembly-third-committee-eu-eov-draft-resolution-l24-report-human-rights-council_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-general-statement-%E2%80%93-un-general-assembly-after-adoption-resolution-human-rights-rohingya-muslims_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-explanation-vote-%E2%80%93-un-general-assembly-resolution-human-rights-safe-drinking-water-and-sanitation_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-explanation-vote-%E2%80%93-un-general-assembly-right-peoples-self-determination_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-6th-committee-revitalisation-work-ga_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-un-security-council-open-debate-conflict-related-hunger_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-un-general-assembly-statement-favour-confirmation-alexander-de-croo-undp-administrator_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/un-general-assembly-3c-combating-glorification-nazism-neo-nazism-and-other-practices-contribute_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-report-iaea-2_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-un-general-assembly-4th-committee-unrwa-general-debate_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-un-general-assembly-4th-committee-interactive-dialogue-unrwa-commissioner-general_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-un-general-assembly-joint-debate-item-120-121-ga-revitalization_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-report-international-criminal-court-1_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-5th-committee-working-methods_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-4th-committee-questions-related-information_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-explanation-vote-%E2%80%93-un-general-assembly-4th-committee-questions-relating-information-draft_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-security-council-open-debate-small-arms-and-light-weapons_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-4th-committee-comprehensive-review-peacekeeping-operations-1_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-general-statement-%E2%80%93-un-general-assembly-1st-committee-vote-cluster-vi-regional-disarmament-and_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-explanation-vote-%E2%80%93-un-general-assembly-1st-committee-vote-report-conference-disarmament_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-6th-committee-report-committee-relations-host-country_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-explanation-position-world-social-summit-political-declaration_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-3rd-committee-interactive-dialogue-people-african-descent-0_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-3rd-committee-interactive-dialogue-sse-mercenaries_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-3rd-committee-report-un-high-commissioner-refugees-0_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-3rd-committee-interactive-dialogue-durban-declaration-and-0_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-1st-committee-cluster-vi-disarmament-machinery_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-explanation-vote-%E2%80%93-un-general-assembly-1st-committee-vote-cluster-iii-outer-space-no-first_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-3rd-committee-interactive-dialogue-discrimination-based-sexual_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-3rd-committee-interactive-dialogue-human-rights-persons-albinism-1_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-3rd-committee-interactive-dialogue-right-education-0_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-world-social-summit-plenary_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-3rd-committee-interactive-dialogue-trafficking-persons-especially-0_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-security-council-debate-bosnia-and-herzegovina-0_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-report-human-rights-council-1_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statements-%E2%80%93-un-general-assembly-6th-committee-cluster-i-report-ilc-work-its-76th-session_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statements-%E2%80%93-un-general-assembly-6th-committee-cluster-ii-report-ilc-work-its-76th-session_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-3rd-committee-interactive-dialogue-cultural-rights-0_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-3rd-committee-interactive-dialogue-contemporary-forms-slavery-0_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-security-council-briefing-unscrs-unmik-un-mission-kosovo_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-3rd-committee-interactive-dialogue-human-rights-iran-0_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-3rd-committee-interactive-dialogue-independent-international-0_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-3rd-committee-interactive-dialogue-human-rights-afghanistan-0_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-3rd-committee-interactive-dialogue-human-rights-experts-nicaragua_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-3rd-committee-interactive-dialogue-international-fact-finding_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-3rd-committee-interactive-dialogue-human-rights-burundi-0_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-3rd-committee-interactive-dialogue-human-rights-eritrea-1_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-3rd-committee-interactive-dialogue-human-rights-south-sudan-0_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-4th-committee-peaceful-uses-outer-space-1_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/d%C3%A9claration-de-l%E2%80%99ue-assembl%C3%A9e-g%C3%A9n%C3%A9rale-des-nations-unies-3e-commission-dialogue-interactif-sur-la_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-3rd-committee-interactive-dialogue-human-rights-somalia-0_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-3rd-committee-interactive-dialogue-human-rights-dprk_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-3rd-committee-interactive-dialogue-freedom-religion-or-belief-0_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-3rd-committee-interactive-dialoue-human-rights-and-transnational_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-3rd-committee-interactive-dialogue-human-rights-myanmar_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-1st-committee-regional-disarmament-and-security-2_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-un-security-council-open-debate-situation-middle-east-including-palestinian-question-0_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/un-security-council-open-debate-maintenance-international-peace-and-security-%E2%80%93-united-nations_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-1st-committee-other-weapons-mass-destruction-2_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-1st-committee-outer-space-1_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-3rd-committee-interactive-dialogue-right-development-1_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-1st-committee-other-disarmament-measures-and-international_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-4th-committee-assistance-mine-action-0_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-6th-committee-un-programme-assistance-international-law-0_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-3rd-committee-interactive-dialogue-human-rights-belarus-0_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-3rd-committee-interactive-dialogue-human-rights-russian-0_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-3rd-committee-interactive-dialogue-commission-inquiry-ukraine-0_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-un-general-assembly-6th-committee-agenda-item-79_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-consultations-informal-ad-hoc-working-group-un80-%E2%80%93-mandate-creation_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-un-general-assembly-1st-committee-statement-conventional-weapons_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-3rd-committee-interactive-dialogue-human-right-clean-healthy-and-0_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-3rd-committee-interactive-dialogue-adequate-housing-0_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-3rd-committee-interactive-dialogue-unilateral-coercive-measures_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-4th-committee-revitalization-work-ga-working-methods_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-3rd-committee-interactive-dialogue-freedom-expression-and-opinion_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-3rd-committee-interactive-dialogue-extreme-poverty-and-human_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-3rd-committee-interactive-dialogue-persons-affected-leprosy_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-3rd-committee-interactive-dialogue-right-health_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-3rd-committee-interactive-dialogue-extrajudicial-summary-or-0_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-3rd-committee-interactive-dialogue-independence-judges-and-0_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-4th-committee-effects-atomic-radiation-1_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-un-general-assembly-report-united-nations-youth-office_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-un-general-assembly-3rd-committee-interactive-dialogue-director-sustainable-development_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-1st-committee-nuclear-weapons-1_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-3rd-committee-interactive-dialogue-high-commissioner-human-rights-0_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-3rd-committee-interactive-dialogue-rights-peasants-and-other-0_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-3rd-committee-interactive-dialogue-chair-human-rights-committee_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-3rd-committee-interactive-dialogue-economic-social-and-cultural_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-security-council-role-young-persons-addressing-security-challenges-mediterranean_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-3rd-committee-interactive-dialogue-torture_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-wsis20-review_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/child-centred-response-sexual-exploitation-children-street-situations_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-un-general-assembly-3rd-committee-interactive-dialogue-rights-child_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-3rd-committee-interactive-dialogue-un-office-drugs-and-crime_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-3rd-committee-interactive-dialogue-indigenous-peoples_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-and-ecosoc-reimagining-public%E2%80%93investor-partnerships_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-2nd-committee-general-debate-2_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-3rd-committee-interactive-dialogue-discrimination-and-violence_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/remarks-ambassador-lambrinidis-un80-bridging-generations-%E2%80%93-youth-and-legacy-multilateralism-event_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-3rd-committee-interactive-dialogue-older-persons-1_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-3rd-committee-interactive-dialogue-un-desa-0_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-1st-committee-general-statement-0_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/un80-bridging-generations-youth-and-legacy-multilateralism_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-security-council-open-debate-women-peace-and-security-0_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-6th-committee-measures-eliminate-international-terrorism-1_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-peacebuilding-commission-meeting-central-african-republic_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-3rd-committee-general-debate_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-use-veto-%E2%80%93-item-64-special-report-security-council-debate_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-high-level-plenary-meeting-30th-anniversary-world-programme_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-high-level-meeting-human-rights-situation-rohingya-muslims-and_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-peacebuilding-committee-post-transition-peace-efforts-chad_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/shaping-water-resilient-world-water-security-all_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/speech-president-european-council-ant%C3%B3nio-costa-80th-united-nations-general-assembly_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-united-nations-annual-ministerial-meeting-foreign-ministers-landlocked-developing_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-high-level-multi-stakeholder-informal-meeting-launch-global-dialogue-artificial_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/speech-president-european-council-ant%C3%B3nio-costa-second-meeting-defense-democracy-fighting-against_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-high-level-meeting-non-communicable-diseases-and-promotion-mental_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/humanitarian-diplomacy-realising-hope-middle-east-crises_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/speech-commission-evp-ribera-high-level-solutions-dialogue-accelerating-early-warning-and-extreme_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/speech-president-european-council-ant%C3%B3nio-costa-united-nations-security-council-ukraine_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/speech-commission-president-von-der-leyen-unga80-side-event-protecting-children-digital-age_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/speech-european-commission-president-von-der-leyen-united-nations-climate-summit-2025_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/full-speed-ahead-global-partnership-eliminate-violence-against-women-girls_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/discours-du-pr%C3%A9sident-ant%C3%B3nio-costa-%C3%A0-la-conf%C3%A9rence-de-haut-niveau-sur-la-solution-%C3%A0-deux-%C3%A9tats_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/speech-commission-president-von-der-leyen-unga-side-event-%E2%80%98restoring-childhood-and-humanity-%E2%80%93_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/ukraine-press-stake-out-high-representative-kaja-kallas-following-un-security-council-meeting_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-commissioner-lahbib-%E2%80%93-un-general-assembly-high-level-meeting-30th-anniversary-fourth_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/speech-commission-president-von-der-leyen-high-level-international-conference-peaceful-settlement_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/hunger-crossfire-no-time-spare-end-famine-now_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/call-action-palestinian-children-west-bank-and-gaza_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/reforming-impact-delivering-better-outcomes-people-through-fit-purpose-aid-system_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/time-accountability-and-justice-rule-law-under-threat-belarus_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-ad-hoc-working-group-un80-first-informal-meeting_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/who-rules-world-ep40-un-usg-melissa-fleming-un80-and-importance-multilateralism_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-security-council-maintenance-international-peace-and-security-poland_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-security-council-future-peace-operations_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-explanation-vote-%E2%80%93-un-general-assembly-adoption-resolution-cooperation-between-un-and-african_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-explanation-vote-%E2%80%93-un-general-assembly-adoption-resolution-cooperation-between-un-and-economic_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-explanation-vote-%E2%80%93-un-general-assembly-adoption-resolution-cooperation-between-un-and-shanghai_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-international-day-clean-air-blue-skies-0_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-%E2%80%93-un-general-assembly-international-day-against-nuclear-tests-0_en",
    "https://www.eeas.europa.eu/delegations/un-new-york/eu-statement-briefing-secretary-general-un80-initiative_en",
]


# ── HTTP helpers ──────────────────────────────────────────────────────────────

def fetch_page(url):
    req = urllib.request.Request(url, headers=HEADERS)
    try:
        with urllib.request.urlopen(req, timeout=15) as resp:
            return resp.read().decode("utf-8", errors="replace")
    except Exception:
        return None


def parse_date(html):
    m = re.search(r'article:published_time"\s+content="(\d{4}-\d{2}-\d{2})', html)
    if m:
        try:
            return datetime.strptime(m.group(1), "%Y-%m-%d").date()
        except ValueError:
            pass
    m = re.search(r"\b(\d{2}\.\d{2}\.(202[0-9]))\b", html)
    if m:
        try:
            return datetime.strptime(m.group(1), "%d.%m.%Y").date()
        except ValueError:
            pass
    return None


def extract_title(html):
    m = re.search(r"<h1[^>]*>(.*?)</h1>", html, re.DOTALL | re.IGNORECASE)
    if m:
        return re.sub(r"<[^>]+>", "", m.group(1)).strip()
    return ""


def extract_body_text(html):
    html = re.sub(r"<(script|style)[^>]*>.*?</\1>", " ", html, flags=re.DOTALL | re.IGNORECASE)
    text = re.sub(r"<[^>]+>", " ", html)
    return re.sub(r"\s+", " ", text).strip()


def extract_alignment_info(text):
    sentences = re.split(r"(?<=[.!?])\s+", text)
    alignment_sentences = [s for s in sentences if ALIGNMENT_PATTERN.search(s)]
    if not alignment_sentences:
        return False, {c: False for c in TRACKED_COUNTRIES}
    combined = " ".join(alignment_sentences)
    country_flags = {
        country: bool(re.search(pattern, combined, re.IGNORECASE))
        for country, pattern in TRACKED_COUNTRIES.items()
    }
    return True, country_flags


# ── Sitemap refresh (optional) ────────────────────────────────────────────────

def refresh_urls_from_sitemap():
    """Scan EEAS sitemap and save fresh URL list to statement_urls.txt."""
    entry_pat = re.compile(r"<url>(.*?)</url>", re.DOTALL)
    loc_pat = re.compile(r"<loc>([^<]+)</loc>")
    lastmod_pat = re.compile(r"<lastmod>([^<]+)</lastmod>")

    seen = set()
    urls = []
    for pg in range(1, SITEMAP_PAGES + 1):
        url = f"https://www.eeas.europa.eu/sitemap.xml?page={pg}"
        try:
            req = urllib.request.Request(url, headers=HEADERS)
            text = urllib.request.urlopen(req, timeout=15).read().decode("utf-8", errors="replace")
        except Exception as e:
            print(f"  Page {pg}: {e}")
            time.sleep(2)
            continue

        for entry in entry_pat.findall(text):
            loc_m = loc_pat.search(entry)
            if not loc_m:
                continue
            loc = loc_m.group(1).split("?")[0]
            if "/delegations/un-new-york/" not in loc:
                continue
            if not loc.endswith("_en"):
                continue
            if any(s in loc for s in SKIP_PATTERNS):
                continue
            lastmod_m = lastmod_pat.search(entry)
            try:
                lastmod = datetime.strptime(lastmod_m.group(1)[:10], "%Y-%m-%d").date() if lastmod_m else None
            except Exception:
                lastmod = None
            if lastmod and lastmod >= UN80_START and loc not in seen:
                seen.add(loc)
                urls.append(loc)

        print(f"  Sitemap page {pg:2d}: {len(urls)} total UN NY URLs so far")
        time.sleep(1.5)

    # Merge extra URLs
    for u in EXTRA_URLS:
        if u not in seen:
            seen.add(u)
            urls.append(u)

    URL_CACHE.write_text("\n".join(urls))
    print(f"\nSaved {len(urls)} URLs to {URL_CACHE}")
    return urls


# ── Main scraping loop ────────────────────────────────────────────────────────

RESULTS_CACHE = Path(__file__).parent / "scraped_results.csv"

# Pages that returned a date before UN80 or no date — no need to re-fetch them
SKIP_CACHE = Path(__file__).parent / "skipped_urls.txt"


def load_cache():
    """Return (rows, done_urls) from previous runs."""
    rows, done = [], set()
    if RESULTS_CACHE.exists():
        cached = pd.read_csv(RESULTS_CACHE)
        rows = cached.to_dict("records")
        done = set(cached["URL"])
    if SKIP_CACHE.exists():
        done |= set(SKIP_CACHE.read_text().splitlines())
    return rows, done


def save_cache(rows):
    pd.DataFrame(rows).to_csv(RESULTS_CACHE, index=False)


def mark_skipped(url):
    with open(SKIP_CACHE, "a") as f:
        f.write(url + "\n")


def scrape_all(urls):
    urls = list(dict.fromkeys(urls))  # deduplicate
    rows, done = load_cache()
    todo = [u for u in urls if u not in done]
    total = len(todo)
    print(f"  Cached from previous runs: {len(rows)} statements "
          f"({len(done)} URLs done, {total} still to fetch)\n")

    failed = 0
    for i, url in enumerate(todo, 1):
        slug = url.rstrip("/").split("/")[-1][:55]
        print(f"  [{i:3d}/{total}] {slug}", end=" ... ", flush=True)

        html = fetch_page(url)
        if not html:
            print("FAILED")
            failed += 1
            if failed >= 10:
                print("\n  Too many consecutive failures — likely rate-limited.")
                print("  Progress is saved; re-run later to continue where we left off.")
                break
            time.sleep(2)
            continue
        failed = 0

        stmt_date = parse_date(html)
        if stmt_date is None:
            print("no date")
            mark_skipped(url)
            continue
        if stmt_date < UN80_START:
            print(f"skip ({stmt_date})")
            mark_skipped(url)
            continue

        title = extract_title(html)
        text = extract_body_text(html)
        has_alignment, country_flags = extract_alignment_info(text)

        row = {
            "Date": str(stmt_date),
            "Title": title or slug,
            "URL": url,
            "Has Alignment": has_alignment,
        }
        row.update(country_flags)
        rows.append(row)
        save_cache(rows)

        print(f"OK [{'ALIGNMENT' if has_alignment else 'no alignment'}] ({stmt_date})")
        time.sleep(0.5)

    print(f"\nTotal statements (cached + new): {len(rows)}")
    return rows


# ── Excel output ──────────────────────────────────────────────────────────────

def write_excel(rows):
    df = pd.DataFrame(rows)
    df["_sort"] = pd.to_datetime(df["Date"], errors="coerce")
    df = df.sort_values("_sort", ascending=False).reset_index(drop=True)
    df["Date"] = df["_sort"].dt.strftime("%d/%m/%Y")
    df = df.drop(columns=["_sort"])

    country_cols = list(TRACKED_COUNTRIES.keys())
    df_aligned = df[df["Has Alignment"] == True].copy()
    total_with_alignment = len(df_aligned)

    summary_rows = []
    for country in country_cols:
        if country in df_aligned.columns:
            count = int(df_aligned[country].sum())
            pct = round(count / total_with_alignment * 100, 1) if total_with_alignment else 0.0
            summary_rows.append({
                "Country": country,
                "Times Aligned": count,
                "Statements with Alignment Clause": total_with_alignment,
                "Alignment %": pct,
            })

    df_summary = pd.DataFrame(summary_rows).sort_values("Alignment %", ascending=False)
    raw_cols = (
        ["Date", "Title", "URL", "Has Alignment"]
        + [c for c in country_cols if c in df.columns]
    )

    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        df_summary.to_excel(writer, sheet_name="Summary", index=False)
        ws = writer.sheets["Summary"]
        for col, w in [("A", 22), ("B", 16), ("C", 38), ("D", 14)]:
            ws.column_dimensions[col].width = w

        df[raw_cols].to_excel(writer, sheet_name="All Statements", index=False)
        ws2 = writer.sheets["All Statements"]
        ws2.column_dimensions["A"].width = 13
        ws2.column_dimensions["B"].width = 65
        ws2.column_dimensions["C"].width = 75
        ws2.column_dimensions["D"].width = 15

        if total_with_alignment:
            df_aligned[raw_cols].to_excel(writer, sheet_name="Alignment Statements Only", index=False)
            ws3 = writer.sheets["Alignment Statements Only"]
            ws3.column_dimensions["A"].width = 13
            ws3.column_dimensions["B"].width = 65
            ws3.column_dimensions["C"].width = 75

    print(f"\nSaved: {OUTPUT_FILE}")
    print(f"\n{'='*60}")
    print(f"Total statements scraped (UN80+):  {len(df)}")
    print(f"Statements WITH alignment clause:  {total_with_alignment}")
    if total_with_alignment:
        print(f"\nTop aligners:")
        print(df_summary[df_summary["Times Aligned"] > 0].to_string(index=False))


# ── Entry point ───────────────────────────────────────────────────────────────

if __name__ == "__main__":
    if "--refresh-urls" in sys.argv:
        print("Refreshing URL list from EEAS sitemap (takes ~2 min)...\n")
        urls = refresh_urls_from_sitemap()
    else:
        urls = list(STATEMENT_URLS)

    print(f"\nScraping {len(dict.fromkeys(urls))} unique statement URLs...\n")
    rows = scrape_all(urls)
    if rows:
        write_excel(rows)
    else:
        print("No statements scraped.")
