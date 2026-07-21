# Electronic Renewal Review

A single-page Streamlit app for making renewal decisions on electronic
subscriptions. Compares what a database *uniquely* provides against usage
signals and produces a per-title recommendation (renew, negotiate, cancel)
along with a full audit trail as a downloadable XLSX brief.

Extracted from the [Howard-Tilton Memorial Library / Tulane
Collection Analysis Suite](https://collection-analysis-suite-tul.streamlit.app/)
for teams that only need the renewal workflow.

---

## What it does

Given an Alma electronic-journal coverage export plus (optionally) a
usage file, the tool:

- Classifies every title in the target subscription as **sole source**,
  **unique coverage**, or **redundant** against the rest of your
  electronic collection
- Computes what you'd actually lose if you cancelled — at day resolution,
  capped at a date you can vouch for so Alma's stale "ongoing" claims
  don't inflate the picture
- Applies a per-title decision matrix with T/L/R protection and
  a 2.5-year threshold rule
- Exports a multi-sheet XLSX brief you can share with liaisons and
  vendors

---

## Quick start

```bash
pip install -r requirements.txt
streamlit run app.py
```

Or deploy to Streamlit Cloud, Heroku, or any Python host that runs
Streamlit apps. The app is a single file with no database or filesystem
dependencies — session state only.

**Python 3.10 or newer.** Tested through 3.14.

---

## The workflow

The page walks through nine steps in sequence:

1. **Setup** — vendor name, renewal deadline, and T/L/R protection mode
   (whole subscription / specific titles / not applicable)
2. **Coverage upload** — Alma electronic-journal coverage export
   (CSV or XLSX)
3. **Flag databases already cancelled** — multiselect to exclude
   phantom-active subscriptions from the redundancy math
4. **Cap ongoing coverage** — date input that caps "Available from X"
   claims at a verified end date (default: today)
5. **Pick the focus database** — the subscription under review
6. **Attach usage** *(optional)* — three modes: skip, direct upload
   (COUNTER TR_J3 or similar), or extract inline from a stacked
   multi-database titles usage report (ProQuest, EBSCO, or similar)
7. **External availability check** — quick lookups you do outside the
   tool: WorldCat, DOAJ, Ulrichsweb, DOAB, SCImago, JCR
8. **Mark T/L/R relevant titles** — interactive editor (when in
   specific-titles mode) or upload a title list
9. **Review the decision matrix** — automated per-title recommendation
   with reasoning, exported as XLSX

---

## The decision matrix

Two overrides fire first, then the standard matrix runs:

| Rule tier | Condition | Result |
|---|---|---|
| **Override 1** | T/L/R relevant | Renew (protected) |
| **Override 2** | Sole source / unique coverage with **< 2.5 yrs** unique loss | Cancel candidate |
| Standard | Sole source · used · ≥ 2.5 yrs unique | **Renew or find equivalent subscription** |
| Standard | Sole source · unused · not T/L/R | Cancel candidate |
| Standard | Unique coverage · used · ≥ 2.5 yrs unique | **Renew/Negotiate or get quote for unique coverage from vendor we already use** |
| Standard | Unique coverage · unused · not T/L/R | Cancel candidate |
| Standard | Redundant · high use | Negotiate / restructure |
| Standard | Redundant · low use | Cancel candidate |

Each decision includes a reasoning string with the specific unique-year
count, so the threshold call is auditable in the exported brief.

---

## Input file formats

### Coverage export (required)

Alma's electronic-journal coverage export or an A-to-Z list export. The
tool auto-detects columns; expected fields:

- **Title** and **Title (Normalized)** — for matching across sources
- **Electronic Collection Public Name** — the database each title lives in
- **Coverage Information Combined** — Alma coverage strings like
  `Available from 2013-08-01 until 2020-10-31` or `Available from 2018`
  (ongoing)
- **Interface Name**, **ISBN (Normalized)** — optional

Multi-clause coverage strings with `volume: X issue: Y` metadata between
`from` and `until` parse correctly.

### Usage file (optional)

Any title-level usage export works: COUNTER TR_J3, a Zero-Use master, a
non-COUNTER vendor report, or the CSV output of the tool's inline
Multi-Database Usage Extractor. Per-year usage columns are summed
automatically.

### Multi-database titles usage report (optional)

For vendors that ship stacked reports (ProQuest, EBSCO, Gale, and
others) with a section per database in one `.xls`. The tool detects
section headers using a keyword library (BriefCitation, Citation, Full
Text, Total, Requests, Searches, and 30+ others) and extracts the focus
database's titles.

### T/L/R title list (optional)

CSV or XLSX with a Title column. Any title matching (by normalized
title) is protected from Cancel-candidate verdicts regardless of the
2.5-year threshold or usage signal.

---

## The exported brief

Multi-sheet XLSX with everything from the review:

| Sheet | Contents |
|---|---|
| Setup | Vendor, deadline, T/L/R mode, coverage-as-of date, excluded databases, usage source, materiality threshold |
| Uniqueness | Per-title status (sole / unique / redundant), unique-year count, unique coverage span, other holders |
| Usage | Per-title recent-period usage (if attached) |
| Decisions | Per-title recommendation with reasoning |
| T/L/R protected | Titles flagged for protection (if any) |
| Cross-tab | Uniqueness × decision matrix as a summary |

Filename convention: `renewal_brief_[VendorName].xlsx`.

---

## External resources referenced

The tool doesn't call any external APIs, but the External availability
check step (Step 7) points liaisons to these for context-of-collection
decisions:

- **[WorldCat](https://www.worldcat.org/)** — broadly held elsewhere?
- **[DOAJ](https://doaj.org/)** — open-access journal equivalent?
- **[DOAB](https://www.doabooks.org/)** — open-access book equivalent?
- **[Ulrichsweb](https://ulrichsweb.serialssolutions.com/)** — still
  active and peer-reviewed?
- **[SCImago Journal Rank](https://www.scimagojr.com/)** — subject
  quartile and impact
- **[Journal Citation Reports](https://jcr.clarivate.com/)** — impact
  factor and category rank

---

## Context

Developed at Howard-Tilton Memorial Library, Tulane University, as part
of the Collection Oversight Committee's annual review cycle. This app
implements the subscription-renewal portion of that cycle as a
standalone tool.

For teams that need the full pipeline — annual print snapshot, ILL
review, ebook snapshot, vendor mapping, watchlist/wishlist lifecycle —
see the [full Collection Analysis Suite](https://collection-analysis-suite-tul.streamlit.app/).

---

## Deployment notes

**Streamlit Cloud** — free tier is enough. Point the app at `app.py`
with the accompanying `requirements.txt`.

**Local** — `streamlit run app.py` from any directory containing the
file. Data never leaves the local session; the tool has no telemetry.

**Data privacy** — coverage exports and usage files stay in Streamlit's
in-memory session state and are discarded when the session ends. No
persistence, no external API calls, no logging beyond Streamlit's default.

**Optional dependencies** — `xlrd` for legacy `.xls` support (needed for
the multi-database extractor); `openpyxl` for XLSX export. Both listed
in `requirements.txt`. If either is missing, the corresponding feature
degrades gracefully.

---

## Credits

Built by **Kay P. Maye**, Scholarly Engagement Librarian & Resource /
Data Analyst, Howard-Tilton Memorial Library, Tulane University.

Guide: *Print & Electronic Resource Analysis Guide* (Collection
Oversight Committee, rev. July 2026).
