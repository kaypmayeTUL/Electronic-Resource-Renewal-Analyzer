"""Electronic Renewal Review — Standalone Streamlit App.

Uniqueness analysis + optional inline multi-database usage extraction +
per-title renewal decisions (renew / negotiate / cancel) for electronic
subscriptions, all on one page.

Features:
- Coverage math at day resolution (with a user-set "coverage as-of" cap)
- Phantom-cancelled-DB exclusion picker
- Per-title T/L/R protection with three modes (whole-subscription /
  specific-titles-uploaded / not-applicable)
- Decision matrix with 2.5-year threshold + T/L/R override
- Multi-vendor usage extraction (ProQuest, EBSCO, similar stacked .xls)
- Multi-sheet XLSX brief export

Deploy: `pip install streamlit pandas openpyxl xlrd plotly` and
`streamlit run renewal_review_app.py`.
"""

import streamlit as st
import pandas as pd
import numpy as np
import re
import warnings
import unicodedata
from io import BytesIO
from datetime import date, timedelta
from collections import defaultdict
import hashlib
import csv

warnings.filterwarnings('ignore')

# xlrd is needed for the multi-database usage extractor (.xls). openpyxl is
# needed for the brief export. Both are optional at runtime — the tool
# degrades gracefully if either is missing.
try:
    import xlrd  # noqa: F401
    XLS_AVAILABLE = True
except ImportError:
    XLS_AVAILABLE = False

try:
    import openpyxl  # noqa: F401
    XLSX_AVAILABLE = True
except ImportError:
    XLSX_AVAILABLE = False

st.set_page_config(
    page_title="Electronic Renewal Review",
    page_icon="🔄",
    layout="wide",
    initial_sidebar_state="collapsed",
)


# ============================================================
# COLUMN-DETECTION CONSTANTS
# ============================================================

TITLE_ALIASES = ['Title', 'title', 'TITLE', 'Book Title', 'Item Title',
                 'Title (Normalized)', 'File Name', 'Filename',
                 'Item Name', 'Resource Title', 'Object Title']

# Weight/usage aliases — ORDER MATTERS. More specific and meaningful usage
# columns appear first so they win the alias contest before generic ones.
# 'Uses' (bare) deliberately omitted — too generic; matches "Remaining CAM Uses"
# which is *available* capacity, not actual usage.

NORM_TITLE_ALIASES = ['Title (Normalized)', 'Normalized Title', 'Title Normalized',
                      'Title (normalized)']

# Metric labels observed in ProQuest "Database Titles Usage Report" section headers.
# Used to auto-detect the header rows that separate database sections in the
# multi-DB .xls exports. Any column-1 value matching one of these signals a
# section header row (which carries the metric column labels for that section).

COVERAGE_ALIASES = ['Coverage Information Combined', 'Coverage Information',
                    'Coverage Statement', 'Coverage Combined', 'Available Coverage',
                    'Date Coverage', 'Coverage Dates', 'Coverage']

COLLECTION_ALIASES = ['Electronic Collection Public Name', 'Electronic Collection',
                      'Collection Public Name', 'Public Name', 'Collection Name',
                      'Package Name', 'Database Name', 'Resource Name',
                      'Collection', 'Package', 'Database']

INTERFACE_ALIASES = ['Interface Name', 'Interface', 'Provider Name', 'Provider',
                     'Platform', 'Vendor', 'Service Provider']

WEIGHT_ALIASES = [
    # COUNTER 5 metrics (most specific — formal e-resource usage)
    'Total_Item_Requests', 'Unique_Item_Requests',
    'Total_Item_Investigations', 'Unique_Item_Investigations',
    'Total_Title_Requests', 'Unique_Title_Requests',
    # EBSCO Detailed Report — actual usage (in preference order)
    'Total Accesses', 'Full Downloads', 'Chapter Downloads', 'Online Views',
    # Print circulation
    'Loans (Total)', 'Loans (In House + Not In House)',
    'Loans', 'Checkouts', 'Circulation', 'checkouts',
    # Digital views / downloads
    'Digital File Downloads', 'Digital File Views',
    'Total Book Downloads', 'Book Downloads', 'Downloads',
    'Read Online (post Trigger) Sessions',
    'Read Online Sessions', 'Sessions',
    # Generic — last resort
    'Views', 'Requests', 'Hits', 'Usage', 'Total Uses', 'Count',
]

# Identifier columns for cross-file matching (used by Zero-Use Identifier).
# ISBN/ISSN/DOI/OCLC are reliable join keys; title+author is the fallback
# when identifiers are absent.

_PERYEAR_RE = re.compile(r'^\s*use\s*(?:fy)?\s*(\d{4})\s*$', re.IGNORECASE)

PROQUEST_METRIC_KEYWORDS = frozenset([
    'BriefCitation', 'Citation', 'Citation/Abstract', 'Full Text', 'Full Text Scanned',
    'Full Text  PDF', 'Full Text PDF', 'PDFLink', 'PDFPage', 'Page View',
    'Preview - PDF', 'Figures and Tables', 'Video', 'Audio', 'Supplemental File',
    'Presentation', 'Spreadsheet', 'Webpage', 'Zip File', 'BriefCitationAbstract',
    'PatentFamily', 'Kwic', 'Dataset', 'FullSizeImage', 'Images', 'Total',
    'Requests', 'Searches', 'Sessions', 'Turnaways', 'Result Clicks',
    'Regular Searches', 'Federated Searches', 'Investigations', 'Unique Items',
])


# ============================================================
# HELPERS — utilities, coverage math, decision matrix
# ============================================================

def normalize_text(text):
    """Lowercase → strip accents → clean punctuation → collapse whitespace."""
    if pd.isna(text) or not isinstance(text, str):
        return ""
    text = text.lower()
    text = unicodedata.normalize("NFD", text)
    text = "".join(ch for ch in text if unicodedata.category(ch) != "Mn")
    text = re.sub(r"[^\w\s]", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def find_column(df_or_cols, aliases, partial=True):
    """Find a column matching any alias. Accepts a DataFrame or list of column names."""
    cols = list(df_or_cols.columns) if isinstance(df_or_cols, pd.DataFrame) else list(df_or_cols)
    for alias in aliases:
        if alias in cols:
            return alias
        if partial:
            for col in cols:
                if alias.lower() in col.lower():
                    return col
    return None


def _make_file_key(uploaded_file):
    """Build a stable cache key from an uploaded file object."""
    if uploaded_file is None:
        return None
    try:
        return (uploaded_file.name, uploaded_file.size)
    except AttributeError:
        # Fallback for file-like objects without .size
        return (uploaded_file.name, None)


def _count_leading_comment_lines(file_bytes, comment='#'):
    """Count contiguous leading '#'-comment lines (a provenance header).

    Our exports — e.g. the Zero-Use explicit-zero master — prepend a '#' metadata
    block via _annotate_csv. Counting only *leading* comment lines lets the loaders
    skip that block when the file is re-ingested, without disturbing a '#' that
    appears inside a real data field. The blank separator after the block is
    handled by pandas' skip_blank_lines.
    """
    head = file_bytes[:65536]
    try:
        text = head.decode('utf-8-sig', errors='replace')
    except Exception:
        text = head.decode('latin-1', errors='replace')
    n = 0
    for line in text.splitlines():
        if line.lstrip('\ufeff').lstrip().startswith(comment):
            n += 1
        else:
            break
    return n


def _load_csv_chunked(file_bytes, filename, cols_to_keep=None):
    """Load CSV or Excel file with minimal memory footprint.

    Despite the name (kept for backward compatibility with cache keys), this now
    dispatches based on filename extension: .xls/.xlsx use pandas.read_excel,
    everything else uses read_csv with utf-8-sig → latin-1 fallback.
    """
    if filename and filename.lower().endswith(('.xls', '.xlsx')):
        # Excel path — usecols works the same way as CSV
        try:
            df = pd.read_excel(BytesIO(file_bytes), usecols=cols_to_keep)
        except Exception:
            # If cols_to_keep failed (e.g., mismatch), try without it
            df = pd.read_excel(BytesIO(file_bytes))
        df.columns = [c.strip() if isinstance(c, str) else c for c in df.columns]
        return df

    # CSV path (original behavior). Our own exports (e.g. the Zero-Use
    # explicit-zero master) carry a leading '#'-comment provenance block from
    # _annotate_csv; skip those lines so the file round-trips cleanly. Only
    # *leading* comment lines are skipped, so a '#' inside a data field survives.
    skip = _count_leading_comment_lines(file_bytes)
    try:
        df = pd.read_csv(BytesIO(file_bytes), encoding='utf-8-sig', low_memory=False,
                         skiprows=skip, usecols=cols_to_keep)
    except Exception:
        try:
            df = pd.read_csv(BytesIO(file_bytes), encoding='latin-1', low_memory=False,
                             skiprows=skip, usecols=cols_to_keep)
        except Exception:
            try:
                df = pd.read_csv(BytesIO(file_bytes), encoding='utf-8-sig', low_memory=False,
                                 skiprows=skip)
            except Exception:
                df = pd.read_csv(BytesIO(file_bytes), encoding='latin-1', low_memory=False,
                                 skiprows=skip)
    df.columns = df.columns.str.strip()
    return df


# =====================================================================
# SHARED: Footer & decision-box helper
# =====================================================================


def _detect_per_year_usage_columns(df):
    """Find per-year usage columns like 'Use FY2023' / 'Use 2023'.

    Returns {column_name: year_int} for columns that match the pattern and hold
    numeric values. These are produced by the Zero-Use Identifier so the master
    can plot true usage volumes per (fiscal) year.
    """
    out = {}
    for c in df.columns:
        if not isinstance(c, str):
            continue
        m = _PERYEAR_RE.match(c)
        if m and pd.api.types.is_numeric_dtype(df[c]):
            yr = int(m.group(1))
            if 1500 <= yr <= 2100:
                out[c] = yr
    return out


def _parse_proquest_usage_report(file_bytes, filename=""):
    """Parse a ProQuest 'Database Titles Usage Report' .xls into per-DB sections.

    ProQuest exports these reports as a single .xls with many database sections
    stacked vertically: row 0 = time frame, row 1 = account info, then repeating
    (section header row = DB name in col 0 + metric labels in cols 1..N, followed
    by N title rows with numeric usage per metric). Metrics vary by DB but usually
    include Total, Full Text, and platform-specific counters.

    Returns:
        dict with:
          time_frame: str from the top row (e.g. "JAN-2026 to JUN-2026") or None
          sections: list of {database, metrics: [str], titles: [{Title, <metric>...}]}
          filename: passed-through for tracking

    Never raises — returns an empty result on unparseable files.
    """
    result = {'time_frame': None, 'sections': [], 'filename': filename,
              'error': None}
    try:
        df = pd.read_excel(BytesIO(file_bytes), engine='xlrd', header=None)
    except Exception as e:
        result['error'] = f"Could not read as .xls: {e}"
        return result

    # Time frame from first cell (best-effort)
    if len(df) and isinstance(df.iloc[0, 0], str) and 'Time Frame' in df.iloc[0, 0]:
        result['time_frame'] = df.iloc[0, 0].replace('Time Frame:', '').strip()

    # Section header rows: col 1 holds a recognized metric-label keyword AND
    # col 0 holds a non-empty string. Title rows have numeric col 1.
    header_rows = []
    for i in range(len(df)):
        v1 = df.iloc[i, 1]
        v0 = df.iloc[i, 0]
        if (isinstance(v0, str) and v0.strip()
                and isinstance(v1, str) and v1.strip() in PROQUEST_METRIC_KEYWORDS):
            header_rows.append(i)

    n_cols = len(df.columns)
    for idx, h_row in enumerate(header_rows):
        db_name = str(df.iloc[h_row, 0]).strip()

        # Metric labels: cols 1 → first NaN
        metric_labels = []
        for col in range(1, n_cols):
            v = df.iloc[h_row, col]
            if pd.notna(v) and isinstance(v, str) and v.strip():
                metric_labels.append((col, v.strip()))
            else:
                break

        # Title rows: from h_row+1 up to next header row
        next_h = header_rows[idx + 1] if idx + 1 < len(header_rows) else len(df)
        titles = []
        for r in range(h_row + 1, next_h):
            title = df.iloc[r, 0]
            if pd.isna(title):
                continue
            title = str(title).strip()
            if not title:
                continue
            row_data = {'Title': title}
            for col, metric in metric_labels:
                v = df.iloc[r, col]
                if pd.isna(v):
                    row_data[metric] = 0
                else:
                    # Prefer int; fall back to float; strings stay strings
                    try:
                        row_data[metric] = int(v)
                    except (TypeError, ValueError):
                        try:
                            row_data[metric] = float(v)
                        except (TypeError, ValueError):
                            row_data[metric] = v
            titles.append(row_data)

        result['sections'].append({
            'database': db_name,
            'metrics': [m for _, m in metric_labels],
            'titles': titles,
        })

    return result


def _ovl_parse_one_date(s, is_end):
    """Parse a single coverage date token (YYYY, YYYY-M, or YYYY-M-D) into a
    date. Year-only / month-only tokens expand to the start of the span for a
    start date and the end of the span for an end date, so "1847 until 1847"
    means the whole of 1847."""
    from datetime import date, timedelta
    parts = str(s).strip().split('-')
    try:
        y = int(parts[0])
    except (ValueError, IndexError):
        return None
    if y < 1 or y > 2999:
        return None
    if len(parts) == 1:
        return date(y, 12, 31) if is_end else date(y, 1, 1)
    try:
        m = int(parts[1])
    except ValueError:
        return date(y, 12, 31) if is_end else date(y, 1, 1)
    m = min(max(m, 1), 12)
    if len(parts) == 2:
        if is_end:
            nxt = date(y + (m == 12), (m % 12) + 1, 1)
            return nxt - timedelta(days=1)
        return date(y, m, 1)
    try:
        d = int(parts[2])
        return date(y, m, d)
    except (ValueError, IndexError):
        # Bad day-of-month — fall back to month bounds
        if is_end:
            nxt = date(y + (m == 12), (m % 12) + 1, 1)
            return nxt - timedelta(days=1)
        return date(y, m, 1)


def _ovl_parse_coverage(text, present):
    """Parse a coverage statement into a list of (start_date, end_date, ongoing)
    intervals. Handles multiple "Available from X until Y;" clauses in one cell.
    An open-ended clause (no "until") ends at `present` and is flagged ongoing."""
    if text is None or (isinstance(text, float)):
        return []
    s = str(text)
    if not s.strip() or s.lower() == 'nan':
        return []
    intervals = []
    # The pattern allows "volume: N issue: M" style metadata between the from-date
    # and the optional "until" clause — Alma coverage often reads like
    #   "Available from 2013-08-01 volume: 1 issue: 1 until 2020-10-31 volume: 8;"
    # so we can't require "until" to sit right after the from-date. The gap
    # matcher [^;]*? stops at the semicolon that separates coverage clauses.
    for mm in re.finditer(r'from\s+([\d\-]+)(?:[^;]*?\s+until\s+([\d\-]+))?', s, re.I):
        start = _ovl_parse_one_date(mm.group(1), is_end=False)
        if start is None:
            continue
        if mm.group(2):
            end = _ovl_parse_one_date(mm.group(2), is_end=True)
            ongoing = False
        else:
            end = present
            ongoing = True
        if end is not None and end >= start:
            intervals.append((start, end, ongoing))
    return intervals


def _ovl_merge(intervals):
    """Merge a list of (start, end) intervals into non-overlapping, sorted pieces."""
    from datetime import timedelta
    segs = sorted(intervals)
    merged = []
    for s, e in segs:
        if merged and s <= merged[-1][1] + timedelta(days=1):
            merged[-1] = (merged[-1][0], max(merged[-1][1], e))
        else:
            merged.append((s, e))
    return merged


def _ovl_subtract(target, others):
    """Return the pieces of `target` not covered by the union of `others`.
    Both are lists of (start, end) tuples. Target is merged first so a title
    listed more than once in the same database isn't double-counted. Result is
    a list of non-overlapping (start, end) pieces."""
    from datetime import timedelta
    merged = _ovl_merge(others)
    target = _ovl_merge(target)
    result = []
    for s, e in target:
        cur = s
        for ms, me in merged:
            if me < cur or ms > e:
                continue
            if ms > cur:
                result.append((cur, min(ms - timedelta(days=1), e)))
            cur = max(cur, me + timedelta(days=1))
            if cur > e:
                break
        if cur <= e:
            result.append((cur, e))
    return [(s, e) for s, e in result if e >= s]


def _ovl_span_years(ivs):
    """Total span of a list of (start, end) intervals, in fractional years."""
    return sum((e - s).days + 1 for s, e in ivs) / 365.25


def _ovl_endpoint_precision(d, is_end):
    """Determine what precision this endpoint needs:
      0 = year (start = Jan 1 / end = Dec 31 — no info lost)
      1 = month (start = day 1 / end = last day of month — day info not needed)
      2 = day (specific date within a month)
    """
    if not is_end:
        if d.month == 1 and d.day == 1:
            return 0
        if d.day == 1:
            return 1
        return 2
    from datetime import timedelta
    next_day = d + timedelta(days=1)
    if next_day.month == 1 and next_day.day == 1:
        return 0
    if next_day.day == 1:
        return 1
    return 2


def _ovl_fmt_at_precision(d, precision):
    """Format a date at a fixed precision level (0=year, 1=month, 2=day-ISO)."""
    if precision == 0:
        return f"{d.year}"
    if precision == 1:
        return d.strftime("%b %Y")
    return d.strftime("%Y-%m-%d")


def _ovl_fmt_ranges(ivs):
    """Render intervals with consistent precision inside each interval — both
    endpoints use the same granularity (the finer of what either endpoint
    needs). Whole-year intervals compress to "YYYY"; month-boundary intervals
    read as "Mon YYYY"; anything with day precision uses ISO "YYYY-MM-DD" so
    the values are unambiguous and math-friendly (sortable, easy to copy into
    Excel or a date-difference calc).

    Examples:
      Jan 1, 2013 – Dec 31, 2020       → "2013–2020"
      Aug 1, 2013 – Oct 31, 2020       → "Aug 2013–Oct 2020"
      Jan 2, 1997 – Jan 31, 1997       → "1997-01-02–1997-01-31"
      Jan 1, 2021 – Jul 10, 2026       → "2021-01-01–2026-07-10"
    """
    out = []
    for s, e in ivs:
        precision = max(_ovl_endpoint_precision(s, is_end=False),
                        _ovl_endpoint_precision(e, is_end=True))
        start_str = _ovl_fmt_at_precision(s, precision)
        end_str = _ovl_fmt_at_precision(e, precision)
        if start_str == end_str:
            out.append(start_str)
        else:
            out.append(f"{start_str}\u2013{end_str}")
    return "; ".join(out)


def _ovl_classify(df, group_col, title_key_col, title_disp_col, coverage_col,
                  min_years, excluded_databases=None, coverage_as_of_date=None):
    """Classify every (database, title) pair. Returns a long DataFrame with one
    row per database-title combination.

    excluded_databases: optional iterable of database names that have already
    been cancelled but still appear in the coverage export. These are silently
    dropped before classification runs, so titles previously marked "redundant"
    because a phantom-cancelled DB also held them get correctly reclassified as
    "sole source" or "unique coverage" against the databases we actually still
    have.

    coverage_as_of_date: date at which ongoing coverage claims are considered
    trustworthy. Defaults to today. Set earlier to conservatively cap "loss"
    computations — Alma coverage often stays flagged as ongoing after a vendor
    has actually stopped providing a title, so extending ongoing to today can
    overstate what we'd actually lose by cancelling.

    Per-title safety cap: after computing unique loss for a database, the
    interval is capped at the max end date observed across all subscriptions
    for that title. This ensures we never claim to lose coverage for dates
    when no subscription in the file ever claimed to hold the title.

    Columns: database, title, status, unique_years, unique_ranges,
             other_count, also_in.
    """
    from datetime import date

    present = coverage_as_of_date or date.today()
    excluded = set(excluded_databases or ())

    # Drop rows for already-cancelled databases before parsing
    if excluded:
        df = df[~df[group_col].astype(str).str.strip().isin(excluded)]

    # Parse coverage once per row.
    parsed = df[coverage_col].apply(lambda t: _ovl_parse_coverage(t, present))

    # Build title -> list of (group, [intervals]) so each title is touched once.
    title_recs = defaultdict(list)
    title_disp = {}
    for grp, tkey, tdisp, ivs in zip(df[group_col], df[title_key_col],
                                     df[title_disp_col], parsed):
        if tkey is None or (isinstance(tkey, float) and pd.isna(tkey)):
            continue
        tkey = str(tkey).strip()
        if not tkey:
            continue
        if pd.isna(grp) or not str(grp).strip():
            continue
        title_recs[tkey].append((str(grp).strip(), ivs))
        title_disp.setdefault(tkey, tdisp)

    rows = []
    for tkey, recs in title_recs.items():
        groups_here = {g for g, _ in recs}
        disp = title_disp.get(tkey, tkey)

        # Safety cap: max end date across all subscriptions for this title.
        # Loss for any single database can't extend past what any subscription
        # claimed to provide — this handles cases where the focus's own
        # ongoing extension would otherwise inflate loss beyond the observable
        # coverage window.
        all_ends = [e for _, ivs in recs for _, e, _ in ivs]
        title_max_end = max(all_ends) if all_ends else None

        for g in groups_here:
            target = [(s, e) for rg, ivs in recs if rg == g for (s, e, _) in ivs]
            other_groups = sorted({rg for rg, _ in recs if rg != g})
            other = [(s, e) for rg, ivs in recs if rg != g for (s, e, _) in ivs]
            if not other_groups:
                status = "Sole source"
                unique = _ovl_merge(target)
            else:
                unique = _ovl_subtract(target, other)
                yrs = _ovl_span_years(unique)
                status = "Unique coverage" if yrs > min_years else "Redundant"
            # Apply per-title max-end safety cap (no-op if focus's own end IS
            # the max, which is the normal case). It kicks in when the focus's
            # coverage extends past all OTHER subscriptions AND was truncated
            # by a coverage-as-of date being applied to others differently.
            if title_max_end is not None:
                unique = [(s, min(e, title_max_end)) for s, e in unique
                          if s <= title_max_end]
                unique = [(s, e) for s, e in unique if e >= s]
            rows.append({
                "database": g,
                "title": disp,
                "status": status,
                "unique_years": round(_ovl_span_years(unique), 2),
                "unique_ranges": _ovl_fmt_ranges(unique),
                "other_count": len(other_groups),
                "also_in": ", ".join(other_groups),
            })
    return pd.DataFrame(rows)


# Status display order + colors (Tulane palette + neutral).
_OVL_STATUS_ORDER = ["Sole source", "Unique coverage", "Redundant"]
_OVL_STATUS_COLORS = {
    "Sole source": "#285C4D",      # Tulane green — most irreplaceable
    "Unique coverage": "#71C5E8",  # Tulane blue — partial loss if cancelled
    "Redundant": "#C9CCCE",        # neutral gray — safe to cancel
}


def _ovl_cached_classification(tool_key, uploaded_file, group_col,
                               title_key_col, title_disp_col, coverage_col,
                               min_years, df, excluded_databases=None,
                               coverage_as_of_date=None):
    """Memoize the (somewhat expensive) classification in session_state, keyed
    on file + the settings that affect the result. Returns the long DataFrame.
    Cache key includes the excluded-databases set and coverage-as-of date so
    toggling either reruns the classification correctly."""
    file_key = _make_file_key(uploaded_file)
    ex_key = tuple(sorted(excluded_databases)) if excluded_databases else ()
    as_of_key = coverage_as_of_date.isoformat() if coverage_as_of_date else None
    sig = (file_key, group_col, title_key_col, coverage_col, min_years,
           ex_key, as_of_key)
    slot = st.session_state.get(f"_ovl_cache_{tool_key}")
    if slot and slot.get("sig") == sig:
        return slot["result"]
    result = _ovl_classify(df, group_col, title_key_col, title_disp_col,
                           coverage_col, min_years,
                           excluded_databases=excluded_databases,
                           coverage_as_of_date=coverage_as_of_date)
    st.session_state[f"_ovl_cache_{tool_key}"] = {"sig": sig, "result": result}
    return result


def _wfe_build_usage_map(usage_df):
    """Return {normalized_title: total_uses} from an arbitrary usage DataFrame.

    Handles files with a single weight column and files with per-year usage
    columns (from Zero-Use Identifier / Multi-Database Usage Extractor output) — sums the
    per-year columns when no single-total column is present.
    """
    usage_map = {}
    if usage_df is None or usage_df.empty:
        return usage_map, None, None
    u_title_col = find_column(usage_df, TITLE_ALIASES)
    u_weight_col = find_column(usage_df, WEIGHT_ALIASES)
    u_peryear = _detect_per_year_usage_columns(usage_df)
    if not u_weight_col and u_peryear:
        peryear_cols = list(u_peryear.keys())
        usage_df = usage_df.copy()
        usage_df['_total_uses'] = (
            usage_df[peryear_cols].apply(pd.to_numeric, errors='coerce')
            .fillna(0).sum(axis=1))
        u_weight_col = '_total_uses'
    if not (u_title_col and u_weight_col):
        return usage_map, u_title_col, u_weight_col
    for _, r in usage_df.iterrows():
        raw_t = r[u_title_col]
        if pd.isna(raw_t):
            continue
        k = normalize_text(raw_t)
        if not k:
            continue
        v = pd.to_numeric(r[u_weight_col], errors='coerce')
        usage_map[k] = usage_map.get(k, 0) + (int(v) if pd.notna(v) else 0)
    return usage_map, u_title_col, u_weight_col


def _wfe_classify_uniqueness(df, coverage_col, group_col, title_disp_col,
                              title_norm_col, min_years, coverage_file,
                              excluded_databases=None,
                              coverage_as_of_date=None):
    """Build the _ovl_key column and run the cached overlap classification.

    Wraps the same normalization + call used by page_overlap_analyzer so the
    workflow page produces identical results to the standalone tool.
    excluded_databases lets the caller treat "phantom-active" subscriptions
    (already cancelled but still in the Alma coverage export) as gone before
    the redundancy math runs. coverage_as_of_date caps ongoing coverage
    claims so loss doesn't extend past a date the analyst can vouch for.
    """
    df = df.copy()
    if title_norm_col and title_norm_col in df.columns:
        df["_ovl_key"] = df[title_norm_col].apply(
            lambda v: normalize_text(v) if pd.notna(v) and str(v).strip() else None)
        blank = df["_ovl_key"].isna() | (df["_ovl_key"] == "")
        df.loc[blank, "_ovl_key"] = df.loc[blank, title_disp_col].apply(
            lambda v: normalize_text(v) if pd.notna(v) else None)
    else:
        df["_ovl_key"] = df[title_disp_col].apply(
            lambda v: normalize_text(v) if pd.notna(v) else None)
    return _ovl_cached_classification(
        "wfe", coverage_file, group_col, "_ovl_key",
        title_disp_col, coverage_col, min_years, df,
        excluded_databases=excluded_databases,
        coverage_as_of_date=coverage_as_of_date)


def _wfe_apply_decision_matrix(long_df, focus_db, tlr_keys, low_use_threshold):
    """Apply the per-title renew / negotiate / cancel decision matrix.

    T/L/R relevance is per-title: `tlr_keys` is a set of normalized title keys
    that the librarian has flagged as teaching/learning/research relevant. A
    row is protected only if its own key is in the set — not because the
    whole subscription is flagged. Pass an empty set for "no titles protected"
    or the full title-key set for "everything protected."

    Rule order (per-title against the focus DB's placements):
      1. T/L/R flag protects sole-source and unique-coverage titles — always
         wins (whether used or unused, whether coverage is thin or thick).
         The librarian's explicit institutional signal beats the automated
         rules.
      2. Below-threshold rule: if unique-loss coverage is < 2.5 years, the
         title falls to Cancel candidate even when sole-source or used.
         Rationale: a fraction of a subscription's worth of unique material
         doesn't justify the subscription cost.
      3. Otherwise, follow the standard matrix:
         - Sole source + used            → Renew or find equivalent subscription
         - Sole source + unused          → Cancel candidate
         - Unique coverage + used        → Renew/Negotiate or get quote
                                            for unique coverage from vendor
                                            we already use
         - Unique coverage + unused      → Cancel candidate
         - Redundant + high use          → Negotiate / restructure
         - Redundant + low use           → Cancel candidate
    """
    sub = long_df[long_df["database"] == focus_db].copy()
    has_usage = "uses" in sub.columns
    if not has_usage:
        sub["uses"] = 0
    # Compute each row's T/L/R flag from its own title
    sub["_tlr_row"] = sub["title"].apply(
        lambda t: normalize_text(t) in tlr_keys if pd.notna(t) else False)

    MIN_UNIQUE_YEARS = 2.5

    def _row_decision(r):
        status = r["status"]
        used = r["uses"] > 0
        heavy = r["uses"] >= low_use_threshold
        tlr = bool(r["_tlr_row"])
        unique_yrs = r["unique_years"]

        if status == "Sole source":
            if tlr:
                use_note = "with use, " if used else "unused, "
                return ("Renew (protected)",
                        f"Sole source ({use_note}{unique_yrs} yrs) — "
                        "flagged as T/L/R relevant; protection outweighs other signals.")
            if unique_yrs < MIN_UNIQUE_YEARS:
                return ("Cancel candidate",
                        f"Sole source but only {unique_yrs} yrs of unique coverage "
                        f"(threshold {MIN_UNIQUE_YEARS} yrs) — not enough material "
                        "to justify the subscription for this title.")
            if used:
                return ("Renew or find equivalent subscription",
                        f"Sole source with recorded use ({unique_yrs} yrs unique) — "
                        "irreplaceable AND earning its keep, but check whether an "
                        "equivalent subscription exists at a lower cost.")
            return ("Cancel candidate",
                    "Sole source but unused and not T/L/R relevant — "
                    "unused, not T/L/R relevant, sole-source exception applies.")

        if status == "Unique coverage":
            if tlr:
                use_note = "with use, " if used else "unused, "
                return ("Renew (protected)",
                        f"Unique coverage ({use_note}{unique_yrs} yrs) — "
                        "flagged as T/L/R relevant; protection outweighs other signals.")
            if unique_yrs < MIN_UNIQUE_YEARS:
                return ("Cancel candidate",
                        f"Unique coverage but only {unique_yrs} yrs "
                        f"(threshold {MIN_UNIQUE_YEARS} yrs) — not enough to "
                        "justify the subscription for this title.")
            if used:
                return ("Renew/Negotiate or get quote for unique coverage from vendor we already use",
                        f"Unique coverage with use ({unique_yrs} yrs unique) — "
                        "renew, negotiate for the gap years, or ask an existing "
                        "vendor for a quote to cover the unique span.")
            return ("Cancel candidate",
                    "Unique coverage but unused and not T/L/R relevant.")

        # Redundant
        if heavy:
            return ("Negotiate / restructure",
                    "Fully covered elsewhere but used — worth negotiating.")
        return ("Cancel candidate",
                "Redundant coverage with low/no use — the cleanest cut.")

    decisions = sub.apply(_row_decision, axis=1)
    sub["Decision"] = decisions.apply(lambda x: x[0])
    sub["Reasoning"] = decisions.apply(lambda x: x[1])
    return sub.drop(columns=["_tlr_row"])



# ============================================================
# THE RENEWAL REVIEW PAGE
# ============================================================

def page_workflow_e():
    """Renewal-Driven Resource Review as a single-page walkthrough.

    Steps: setup → uniqueness (Step 1) → usage (Step 2, branch-agnostic) →
    decision matrix (renew / negotiate / cancel-candidate). Coverage and usage
    files uploaded here are held in session only.
    """
    st.header("🔄 Renewal-Driven Resource Review")
    st.markdown(
        "**Everything for one subscription's renewal on one page.** Setup, "
        "uniqueness classification, usage triage, and the renew / "
        "negotiate / cancel decision matrix — no tool-hopping."
    )
    with st.expander("ℹ️ How this workflow works", expanded=False):
        st.markdown(
            "This page walks through the review in four steps:\n\n"
            "1. **Setup** — vendor, renewal deadline, and any T/L/R relevance "
            "note the data can't see.\n"
            "2. **Uniqueness (Step 1)** — upload the Alma coverage / A-to-Z "
            "export; the tool classifies every title as sole source, unique "
            "coverage, or redundant.\n"
            "3. **Usage (Step 2)** — upload a title-level usage file (COUNTER, "
            "non-COUNTER vendor report, Zero-Use master, or multi-database extract). "
            "Usage attaches to the uniqueness classification.\n"
            "4. **Decision matrix** — the tool applies renew / negotiate / "
            "cancel-candidate rules to each title of the focus database and "
            "produces a downloadable brief."
        )

    st.markdown("---")

    # =============================================================
    # 1. SETUP
    # =============================================================
    st.subheader("1️⃣ Setup")
    c1, c2 = st.columns(2)
    with c1:
        vendor_name = st.text_input(
            "Vendor / package name:", key="wfe_vendor",
            placeholder="e.g., ABI/INFORM Global (ProQuest)"
        )
    with c2:
        from datetime import date
        deadline = st.date_input(
            "Renewal deadline:", value=None, key="wfe_deadline",
            help="Approximate is fine — this just anchors the review."
        )
    tlr_mode = st.radio(
        "Teaching / Learning / Research (T/L/R) protection:",
        ["Not applicable to this review",
         "Whole subscription is T/L/R relevant",
         "Specific titles are T/L/R relevant (pick in Step 4)"],
        key="wfe_tlr_mode",
        help="**Not applicable** — no titles are protected from cancellation "
             "on T/L/R grounds; the decision matrix runs on uniqueness × "
             "usage alone.  \n"
             "**Whole subscription** — every unused sole-source or "
             "unique-coverage title is protected. Use for specialized "
             "discipline databases that are wall-to-wall course-relevant.  \n"
             "**Specific titles** — the common case. Optionally upload a "
             "T/L/R title list below (course readings, faculty citations); "
             "you can also tick additional titles interactively in Step 4."
    )
    tlr_note = st.text_area(
        "T/L/R justification (optional, appears in the exported brief):",
        key="wfe_tlr_note",
        placeholder="Course dependencies, faculty citations, grant work…",
        height=68
    )
    tlr_list_file = None
    if tlr_mode.startswith("Specific"):
        tlr_list_file = st.file_uploader(
            "Optional: T/L/R title list (CSV, XLS, XLSX) — "
            "titles matched by normalized title",
            type=['csv', 'xls', 'xlsx'], key="wfe_tlr_list",
            help="A one-column file of titles known to be T/L/R relevant "
                 "(course reserves, syllabi, faculty citation lists). "
                 "Titles that match the coverage export get pre-checked in "
                 "Step 4; unmatched titles are reported so you can verify "
                 "spelling or add them separately."
        )

    st.markdown("---")

    # =============================================================
    # 2. UNIQUENESS (STEP 1)
    # =============================================================
    st.subheader("2️⃣ Uniqueness — Step 1")
    st.caption(
        "Upload the Alma electronic-journal coverage / A-to-Z export "
        "(one row per title × database, with a coverage statement)."
    )
    coverage_file = st.file_uploader(
        "Coverage / A-Z export (CSV, XLS, XLSX)",
        type=['csv', 'xls', 'xlsx'], key="wfe_coverage_file"
    )
    if not coverage_file:
        st.info("Upload the coverage export to run the uniqueness classification.")
        return

    try:
        if coverage_file.name.lower().endswith(('.xls', '.xlsx')):
            engine = 'xlrd' if coverage_file.name.lower().endswith('.xls') else 'openpyxl'
            df = pd.read_excel(BytesIO(coverage_file.getvalue()), engine=engine)
        else:
            df = _load_csv_chunked(coverage_file.getvalue(), coverage_file.name)
    except Exception as e:
        st.error(f"❌ Couldn't read coverage export: {e}")
        return

    st.success(f"✅ Loaded {len(df):,} coverage rows.")

    coverage_col = find_column(df, COVERAGE_ALIASES)
    collection_col = find_column(df, COLLECTION_ALIASES)
    interface_col = find_column(df, INTERFACE_ALIASES)
    title_disp_col = find_column(df, TITLE_ALIASES)
    title_norm_col = find_column(df, NORM_TITLE_ALIASES)

    group_col = collection_col or interface_col

    if not (title_disp_col and coverage_col and group_col):
        st.error(
            f"❌ Need title, coverage, and database columns. Detected: "
            f"title=`{title_disp_col}`, coverage=`{coverage_col}`, "
            f"database=`{collection_col}`, interface=`{interface_col}`. "
            f"Use the Overlap & Uniqueness tool (under Individual tools) if "
            f"the columns need manual overrides."
        )
        return

    min_years = st.slider(
        "Materiality threshold — minimum unique span to count as unique coverage (years):",
        0.0, 5.0, 0.0, 0.25, key="wfe_min_years",
        help="Raise to ignore small gaps that come from year-level metadata rounding."
    )

    # ---- Already-cancelled databases (phantom-active in Alma) ----
    # Some databases stay in the Alma coverage export even after they've been
    # cancelled. If we don't tell the tool to ignore them, titles show up as
    # "redundant" because the phantom DB "still" holds them — masking a real
    # loss. This multiselect lets the analyst nominate any such subscriptions
    # to exclude from the redundancy math before classification runs.
    all_dbs_in_file = sorted(df[group_col].dropna().astype(str).str.strip().unique())
    all_dbs_in_file = [d for d in all_dbs_in_file if d]
    excluded_dbs = st.multiselect(
        "Databases already cancelled but still in the coverage export "
        "(exclude from redundancy math):",
        all_dbs_in_file,
        key="wfe_excluded_dbs",
        help="Any database selected here is treated as if it doesn't hold "
             "any of its listed titles. Use this when Alma still shows a "
             "cancelled subscription as active — otherwise the tool thinks "
             "titles held by that phantom database are 'redundant' and won't "
             "flag the true loss you'd take by cancelling the focus database."
    )
    if excluded_dbs:
        st.caption(f"⚠️ Excluding **{len(excluded_dbs)}** database(s) from "
                   f"redundancy math: {', '.join(excluded_dbs)}")

    # ---- Coverage-as-of date ----
    # Alma's coverage claims often extend as "ongoing" past the point where
    # a vendor actually stopped providing a title (nobody updates the record).
    # If ongoing coverage is extended all the way to today, the tool overstates
    # loss for the cancellation review — showing years of "loss" for coverage
    # we never actually had. This date input lets the analyst cap ongoing
    # claims at a date they can vouch for. Default is today; set earlier to
    # be conservative.
    from datetime import date as _date_cls, timedelta
    _today = _date_cls.today()
    coverage_as_of = st.date_input(
        "Cap ongoing coverage claims at (\"coverage as-of\" date):",
        value=_today,
        min_value=_date_cls(2000, 1, 1),
        max_value=_today,
        key="wfe_coverage_as_of",
        help="Any subscription with 'Available from X' (no end date) is treated "
             "as extending only up to this date. If you know a subscription's "
             "vendor actually stopped providing titles before today, set this "
             "to a date you can verify — otherwise the tool will claim loss "
             "for years the vendor wasn't actually delivering. Default is today. "
             "A common conservative choice is the end of the most recent fiscal "
             "year where you can verify all subscriptions were current."
    )
    if coverage_as_of < _today:
        _days_back = (_today - coverage_as_of).days
        st.caption(f"📅 Ongoing coverage capped at "
                   f"**{coverage_as_of.strftime('%b %-d, %Y')}** "
                   f"({_days_back:,} days before today). Loss beyond this date "
                   f"won't be attributed to any subscription.")

    with st.spinner("Classifying uniqueness…"):
        long_df = _wfe_classify_uniqueness(
            df, coverage_col, group_col, title_disp_col,
            title_norm_col, min_years, coverage_file,
            excluded_databases=excluded_dbs or None,
            coverage_as_of_date=coverage_as_of)

    if long_df.empty:
        st.warning("No title/database pairs could be built.")
        return

    # ---- Focus database picker ----
    db_list = sorted(long_df["database"].unique())
    # Prefill from vendor_name if it matches a database exactly (or nearly)
    default_idx = 0
    if vendor_name:
        for i, db in enumerate(db_list):
            if vendor_name.lower() in db.lower():
                default_idx = i
                break
    focus_db = st.selectbox(
        "Which database is under review?",
        db_list, index=default_idx, key="wfe_focus_db"
    )

    focus_sub = long_df[long_df["database"] == focus_db]
    n_sole = int((focus_sub["status"] == "Sole source").sum())
    n_uniq = int((focus_sub["status"] == "Unique coverage").sum())
    n_red = int((focus_sub["status"] == "Redundant").sum())
    n_tot = len(focus_sub)

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Titles in focus DB", f"{n_tot:,}")
    k2.metric("Sole source", f"{n_sole:,}",
              help="Lost entirely if cancelled.")
    k3.metric("Unique coverage", f"{n_uniq:,}",
              help="Kept, but with a date gap if cancelled.")
    k4.metric("Redundant", f"{n_red:,}",
              help="Fully duplicated elsewhere.")

    from datetime import date as _wfe_date
    _cap_label = (coverage_as_of.strftime('%b %-d, %Y')
                  if coverage_as_of < _wfe_date.today()
                  else "today")
    st.caption(
        f"📅 **Loss shown as of {_cap_label}**. Coverage extending into the "
        "future is clipped to that date; ongoing subscriptions display "
        "specific end dates (e.g. \"Jul 2026\") rather than year rollups so "
        "partial-year loss reads accurately. Loss for any title is also "
        "capped at the max end date across all subscriptions for that title, "
        "so we never claim to lose coverage for dates when no subscription "
        "actually held the title."
    )

    st.markdown("---")

    # =============================================================
    # 3. USAGE (STEP 2)
    # =============================================================
    st.subheader("3️⃣ Usage — Step 2")
    st.caption(
        "Attach usage data so the decision matrix can weigh irreplaceability "
        "against actual use. Skip if you're doing a uniqueness-only review "
        "(the matrix will assume 0 uses for every title)."
    )

    usage_source = st.radio(
        "Usage source:",
        ["Skip (uniqueness-only)",
         "Upload usage file directly",
         "Extract from multi-database titles usage report"],
        key="wfe_usage_source", horizontal=False,
        help="The 'Extract' option works with ProQuest, EBSCO, and other "
             "vendors that pack many databases into one stacked `.xls`. See "
             "the Multi-Database Usage Extractor (Individual tools) for "
             "standalone use."
    )

    usage_map = {}
    usage_source_desc = None

    if usage_source == "Upload usage file directly":
        usage_file = st.file_uploader(
            "Usage file (CSV, XLS, XLSX)",
            type=['csv', 'xls', 'xlsx'], key="wfe_usage_upload",
            help="Any title-level usage file: COUNTER TR_J3, a Zero-Use master, "
                 "a non-COUNTER vendor report, or the CSV from the "
                 "Multi-Database Usage Extractor. Per-year usage columns are "
                 "summed automatically."
        )
        if usage_file:
            try:
                if usage_file.name.lower().endswith(('.xls', '.xlsx')):
                    engine = 'xlrd' if usage_file.name.lower().endswith('.xls') else 'openpyxl'
                    usage_df_raw = pd.read_excel(BytesIO(usage_file.getvalue()), engine=engine)
                else:
                    usage_df_raw = _load_csv_chunked(usage_file.getvalue(), usage_file.name)
                usage_map, u_title_col, u_weight_col = _wfe_build_usage_map(usage_df_raw)
                if u_title_col and u_weight_col:
                    usage_source_desc = (
                        f"`{usage_file.name}` (title=`{u_title_col}`, "
                        f"uses=`{u_weight_col}`; {len(usage_map):,} distinct titles)"
                    )
                else:
                    st.warning(
                        f"Couldn't detect title + usage columns in **{usage_file.name}**. "
                        f"Title col: `{u_title_col}`, usage col: `{u_weight_col}`."
                    )
            except Exception as e:
                st.warning(f"Couldn't parse usage file: {e}")

    elif usage_source == "Extract from multi-database titles usage report":
        if not XLS_AVAILABLE:
            st.error("Reading legacy `.xls` files needs the `xlrd` package.")
        else:
            pq_files = st.file_uploader(
                "Multi-database titles usage report `.xls` file(s):",
                type=['xls'], accept_multiple_files=True, key="wfe_pq_files",
                help="ProQuest, EBSCO, or other vendors that ship a stacked "
                     "multi-DB .xls. The tool will extract usage for the focus "
                     "database and combine periods."
            )
            if pq_files:
                pq_metric_default = ['Total']
                pq_metric_choice = st.multiselect(
                    "Metrics to sum:", pq_metric_default + ['Full Text', 'Full Text  PDF', 'Page View'],
                    default=pq_metric_default, key="wfe_pq_metrics",
                    help="Defaults to the vendor's precomputed 'Total' when present. Add "
                         "Full Text / PDF for a stricter measure."
                )
                # Try to auto-detect a section that matches the focus_db
                aggregated = {}  # title → total uses
                matched_sections = 0
                for pqf in pq_files:
                    pqf.seek(0)
                    parsed = _parse_proquest_usage_report(pqf.read(), pqf.name)
                    if parsed.get('error'):
                        st.warning(f"{pqf.name}: {parsed['error']}")
                        continue
                    for sec in parsed['sections']:
                        # Fuzzy match: focus_db in section name or vice versa
                        if (focus_db.lower() in sec['database'].lower()
                                or sec['database'].lower() in focus_db.lower()):
                            matched_sections += 1
                            for tr in sec['titles']:
                                title = str(tr.get('Title', '')).strip()
                                if not title:
                                    continue
                                use_total = 0
                                for m in pq_metric_choice:
                                    v = tr.get(m, 0)
                                    if isinstance(v, (int, float)) and not pd.isna(v):
                                        use_total += int(v)
                                aggregated[title] = aggregated.get(title, 0) + use_total
                if matched_sections:
                    for title, use_total in aggregated.items():
                        k = normalize_text(title)
                        if k:
                            usage_map[k] = usage_map.get(k, 0) + use_total
                    usage_source_desc = (
                        f"Multi-database extract — {len(pq_files)} file(s), "
                        f"{matched_sections} matching section(s) for '{focus_db}', "
                        f"{len(usage_map):,} distinct titles."
                    )
                    st.success(
                        f"✅ Extracted usage for '{focus_db}' from {matched_sections} "
                        f"section(s) across {len(pq_files)} file(s)."
                    )
                else:
                    st.warning(
                        f"No sections matching '{focus_db}' in the uploaded files. "
                        f"Check that the focus database name aligns with a "
                        f"section name in the file."
                    )

    # Attach usage to long_df
    has_usage = bool(usage_map)
    if has_usage:
        long_df = long_df.copy()
        long_df["_k"] = long_df["title"].apply(
            lambda t: normalize_text(t) if pd.notna(t) else None)
        long_df["uses"] = long_df["_k"].map(usage_map).fillna(0).astype(int)
        long_df = long_df.drop(columns=["_k"])
        focus_sub = long_df[long_df["database"] == focus_db]
        sole_used = int(((focus_sub["status"] == "Sole source") & (focus_sub["uses"] > 0)).sum())
        uniq_used = int(((focus_sub["status"] == "Unique coverage") & (focus_sub["uses"] > 0)).sum())
        red_used = int(((focus_sub["status"] == "Redundant") & (focus_sub["uses"] > 0)).sum())
        st.info(
            f"**Usage attached.** In '{focus_db}': "
            f"**{sole_used:,}** of {n_sole:,} sole-source titles are used · "
            f"**{uniq_used:,}** of {n_uniq:,} unique-coverage titles are used · "
            f"**{red_used:,}** of {n_red:,} redundant titles are used."
        )

    st.markdown("---")

    # =============================================================
    # 4. DECISION MATRIX
    # =============================================================
    st.subheader("4️⃣ Decision matrix")

    low_use_threshold = st.number_input(
        "Low-use threshold (uses per title to treat as 'used' for redundant titles):",
        min_value=1, max_value=1000, value=5, step=1, key="wfe_low_use_thresh",
        help="Redundant titles above this threshold get 'Negotiate / restructure' "
             "(worth keeping the good stuff, dropping the tail). Below → cancel "
             "candidate."
    ) if has_usage else 5

    # ---- Build the T/L/R key set based on setup mode ----
    # Modes:
    #   "Not applicable"       → empty set (no titles protected)
    #   "Whole subscription"   → all title keys in the focus DB (universal protect)
    #   "Specific titles"      → keys from upload ∪ keys ticked interactively
    focus_placements = long_df[long_df["database"] == focus_db].copy()
    focus_placements["_key"] = focus_placements["title"].apply(
        lambda t: normalize_text(t) if pd.notna(t) else None)

    tlr_keys = set()
    tlr_upload_keys = set()
    tlr_upload_summary = None
    if tlr_list_file is not None and tlr_mode.startswith("Specific"):
        try:
            if tlr_list_file.name.lower().endswith(('.xls', '.xlsx')):
                engine = 'xlrd' if tlr_list_file.name.lower().endswith('.xls') else 'openpyxl'
                tlr_df_raw = pd.read_excel(BytesIO(tlr_list_file.getvalue()), engine=engine)
            else:
                tlr_df_raw = _load_csv_chunked(tlr_list_file.getvalue(), tlr_list_file.name)
            t_col = find_column(tlr_df_raw, TITLE_ALIASES) or tlr_df_raw.columns[0]
            for raw_t in tlr_df_raw[t_col].dropna():
                k = normalize_text(raw_t)
                if k:
                    tlr_upload_keys.add(k)
            matched = tlr_upload_keys & set(focus_placements["_key"].dropna())
            unmatched = tlr_upload_keys - matched
            tlr_upload_summary = (
                f"Uploaded T/L/R list `{tlr_list_file.name}` "
                f"(title column: `{t_col}`) — "
                f"**{len(matched):,}** of {len(tlr_upload_keys):,} titles matched "
                f"the focus database."
            )
            if unmatched:
                with st.expander(
                    f"⚠️ {len(unmatched):,} T/L/R titles not found in '{focus_db}'"
                ):
                    st.caption(
                        "These titles didn't match anything in the coverage export. "
                        "Check spelling, or add them separately via the interactive "
                        "editor below."
                    )
                    st.dataframe(
                        pd.DataFrame({'Unmatched T/L/R title': sorted(unmatched)}),
                        use_container_width=True, hide_index=True
                    )
        except Exception as e:
            st.warning(f"Couldn't parse T/L/R list: {e}")

    if tlr_mode == "Whole subscription is T/L/R relevant":
        tlr_keys = set(focus_placements["_key"].dropna())
        st.info(f"**T/L/R protection: whole subscription** — all "
                f"**{len(tlr_keys):,}** titles in '{focus_db}' are treated as "
                f"T/L/R relevant.")
    elif tlr_mode.startswith("Specific"):
        # Show interactive editor over the titles that would benefit from
        # T/L/R protection (unused sole-source + unique-coverage). Pre-check
        # from the upload; librarian can add or remove.
        protectable = focus_placements[
            focus_placements["status"].isin(["Sole source", "Unique coverage"])
        ].copy()
        if has_usage:
            protectable = protectable[protectable["uses"] == 0]
        protectable = protectable.sort_values(["status", "title"]).reset_index(drop=True)

        st.markdown("**Mark titles as T/L/R relevant**")
        if tlr_upload_summary:
            st.caption(tlr_upload_summary)
        st.caption(
            "Only unused sole-source and unique-coverage titles are shown — "
            "those are the ones that would be cancel candidates without T/L/R "
            "protection. Used titles are already renew-recommended and don't "
            "need the flag. Tick a title to protect it from cancellation."
        )
        if len(protectable) == 0:
            st.info("No unused sole-source or unique-coverage titles in the "
                    "focus database. T/L/R protection has nothing to attach to.")
        else:
            editor_df = pd.DataFrame({
                'T/L/R': protectable["_key"].apply(lambda k: k in tlr_upload_keys),
                'Title': protectable["title"],
                'Status': protectable["status"],
                'Uses': protectable["uses"] if has_usage else 0,
                'Unique coverage (years)': protectable["unique_ranges"],
            })
            edited = st.data_editor(
                editor_df,
                key="wfe_tlr_editor",
                disabled=["Title", "Status", "Uses", "Unique coverage (years)"],
                hide_index=True,
                use_container_width=True,
                column_config={
                    'T/L/R': st.column_config.CheckboxColumn(
                        "T/L/R relevant?", help="Tick to protect from cancellation."),
                    'Uses': st.column_config.NumberColumn(format="%d"),
                },
                height=min(600, 40 + 35 * len(protectable)),
            )
            # Extract ticked keys
            ticked_titles = edited[edited['T/L/R']]['Title'].tolist()
            tlr_keys = set(normalize_text(t) for t in ticked_titles if pd.notna(t))
            n_ticked = len(tlr_keys)
            n_from_upload = len(tlr_keys & tlr_upload_keys)
            n_added = n_ticked - n_from_upload
            st.caption(
                f"**T/L/R protection:** {n_ticked:,} title(s) selected "
                f"({n_from_upload:,} from upload, {n_added:,} added interactively)."
            )
    else:
        st.info("**T/L/R protection: none** — the decision matrix runs on "
                "uniqueness × usage alone.")

    decision_df = _wfe_apply_decision_matrix(long_df, focus_db, tlr_keys, low_use_threshold)

    # Summary counts
    dcounts = decision_df["Decision"].value_counts()
    _renew_total = int(dcounts.get("Renew or find equivalent subscription", 0)
                       + dcounts.get("Renew (protected)", 0)
                       + dcounts.get("Renew/Negotiate or get quote for unique coverage from vendor we already use", 0))
    _negot_total = int(dcounts.get("Negotiate / restructure", 0))
    _cancel_total = int(dcounts.get("Cancel candidate", 0))

    d1, d2, d3, d4 = st.columns(4)
    d1.metric("Titles reviewed", f"{len(decision_df):,}")
    d2.metric("Recommend RENEW", f"{_renew_total:,}",
              help="Renew, Renew (protected), or Renew/Negotiate combined.")
    d3.metric("Recommend NEGOTIATE", f"{_negot_total:,}",
              help="Redundant-but-used titles worth restructuring.")
    d4.metric("Recommend CANCEL", f"{_cancel_total:,}",
              help="Candidates for dropping — check each against T/L/R notes.")

    # Recommendation callout
    if _cancel_total > 0.5 * len(decision_df) and _renew_total < 0.2 * len(decision_df):
        st.warning(
            f"⚠️ **Cancellation trend**: {_cancel_total:,} of {len(decision_df):,} "
            f"titles ({_cancel_total/len(decision_df)*100:.0f}%) meet the "
            f"cancel-candidate criteria. Consider a full cancellation review — "
            f"is this subscription still earning its cost?"
        )
    elif _renew_total > 0.5 * len(decision_df):
        st.success(
            f"✅ **Retention signal**: {_renew_total:,} of {len(decision_df):,} "
            f"titles ({_renew_total/len(decision_df)*100:.0f}%) are recommended "
            f"for renewal. The subscription is earning its keep."
        )
    else:
        st.info(
            f"**Mixed signal**: {_renew_total:,} renew, {_negot_total:,} "
            f"negotiate, {_cancel_total:,} cancel-candidate. A restructure "
            f"negotiation likely gets the best outcome."
        )

    # Decision breakdown table
    with st.expander("📊 Decision breakdown by category", expanded=True):
        # Cross-tab: status × decision
        crosstab = pd.crosstab(decision_df["status"], decision_df["Decision"], margins=True, margins_name="Total")
        st.dataframe(crosstab, use_container_width=True)

    # Per-title recommendations
    with st.expander("📋 Per-title recommendations", expanded=False):
        show_only = st.multiselect(
            "Show decisions:",
            sorted(decision_df["Decision"].unique()),
            default=sorted(decision_df["Decision"].unique()),
            key="wfe_decision_filter"
        )
        display = decision_df[decision_df["Decision"].isin(show_only)].copy()
        display = display.rename(columns={
            "title": "Title", "status": "Uniqueness", "uses": "Uses",
            "unique_years": "Unique years",
            "unique_ranges": "Unique coverage (years)",
            "also_in": "Also available in"
        })
        cols = ["Title", "Uniqueness", "Uses", "Decision", "Reasoning",
                "Unique years", "Unique coverage (years)", "Also available in"]
        st.dataframe(
            display[cols].sort_values(["Decision", "Uses"], ascending=[True, False]),
            use_container_width=True, hide_index=True
        )

    # =============================================================
    # 5. EXPORT BRIEF
    # =============================================================
    st.markdown("---")
    st.subheader("📥 Export renewal review brief")

    # Build a small brief header + the decision table
    # Human-readable T/L/R summary for the brief
    if tlr_mode == "Not applicable to this review":
        tlr_summary = "Not applicable"
    elif tlr_mode == "Whole subscription is T/L/R relevant":
        tlr_summary = f"Whole subscription — {len(tlr_keys):,} titles protected"
    else:
        tlr_summary = f"Per-title — {len(tlr_keys):,} title(s) flagged"

    brief_header = pd.DataFrame([
        {'Field': 'Vendor / package', 'Value': vendor_name or focus_db},
        {'Field': 'Focus database (from coverage)', 'Value': focus_db},
        {'Field': 'Renewal deadline', 'Value': str(deadline) if deadline else '—'},
        {'Field': 'T/L/R protection mode', 'Value': tlr_summary},
        {'Field': 'T/L/R justification', 'Value': tlr_note or '—'},
        {'Field': 'Coverage file', 'Value': coverage_file.name},
        {'Field': 'Coverage as-of date',
         'Value': coverage_as_of.strftime('%b %-d, %Y') +
                  (' (today)' if coverage_as_of >= _wfe_date.today() else '')},
        {'Field': 'Excluded (already-cancelled) databases',
         'Value': ', '.join(excluded_dbs) if excluded_dbs else '(none)'},
        {'Field': 'Usage source', 'Value': usage_source_desc or '(none — uniqueness only)'},
        {'Field': 'Materiality threshold (yrs)', 'Value': str(min_years)},
        {'Field': 'Low-use threshold', 'Value': str(low_use_threshold) if has_usage else '—'},
        {'Field': 'Titles reviewed', 'Value': f"{len(decision_df):,}"},
        {'Field': 'Recommend RENEW (all types)', 'Value': f"{_renew_total:,}"},
        {'Field': 'Recommend NEGOTIATE', 'Value': f"{_negot_total:,}"},
        {'Field': 'Recommend CANCEL', 'Value': f"{_cancel_total:,}"},
    ])

    # Build the T/L/R-protected titles list for the brief
    tlr_titles_list = []
    if tlr_keys:
        tlr_titles_list = sorted(
            focus_placements[focus_placements["_key"].isin(tlr_keys)]["title"].unique()
        )

    csv_buf = BytesIO()
    csv_buf.write(b"# Renewal Review Brief\n")
    csv_buf.write(f"# Generated: {pd.Timestamp.now()}\n\n".encode('utf-8'))
    csv_buf.write(b"# ---- Setup ----\n")
    brief_header.to_csv(csv_buf, index=False)
    if tlr_titles_list:
        csv_buf.write(b"\n# ---- T/L/R protected titles ----\n")
        pd.DataFrame({'T/L/R protected title': tlr_titles_list}).to_csv(csv_buf, index=False)
    csv_buf.write(b"\n# ---- Per-title decisions ----\n")
    dcsv = decision_df.rename(columns={
        "title": "Title", "status": "Uniqueness", "uses": "Uses",
        "unique_years": "Unique years", "unique_ranges": "Unique coverage (years)",
        "also_in": "Also available in"
    })[["Title", "Uniqueness", "Uses", "Decision", "Reasoning",
        "Unique years", "Unique coverage (years)", "Also available in"]]
    dcsv.to_csv(csv_buf, index=False)

    safe_name = re.sub(r'[^\w\-]+', '_', focus_db)[:60].strip('_') or "renewal"
    st.download_button(
        "📥 Renewal review brief (CSV)",
        csv_buf.getvalue(),
        f"renewal_brief_{safe_name}.csv",
        "text/csv",
        key="wfe_dl_brief"
    )

    # XLSX version with separate sheets
    if XLSX_AVAILABLE:
        xbuf = BytesIO()
        with pd.ExcelWriter(xbuf, engine='openpyxl') as writer:
            brief_header.to_excel(writer, sheet_name='Setup', index=False)
            dcsv.to_excel(writer, sheet_name='Decisions', index=False)
            crosstab.to_excel(writer, sheet_name='Summary')
            if tlr_titles_list:
                pd.DataFrame({'T/L/R protected title': tlr_titles_list}).to_excel(
                    writer, sheet_name='T_L_R titles', index=False)
        st.download_button(
            "📥 Renewal review brief (XLSX, multi-sheet)",
            xbuf.getvalue(),
            f"renewal_brief_{safe_name}.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="wfe_dl_brief_xlsx"
        )



# ============================================================
# MAIN
# ============================================================

def main():
    """Standalone entry point — no sidebar nav, no other tools."""
    # A slim top strip that echoes the parent-suite look without needing the
    # full dashboard chrome.
    st.markdown(
        """
        <div style="background:linear-gradient(135deg,#285C4D 0%,#1a3d32 100%);
                    color:white;padding:14px 20px;margin:-1rem -1rem 20px;
                    border-radius:0 0 8px 8px;">
          <div style="font-family:'Source Serif 4',Georgia,serif;
                      font-weight:600;font-size:1.3rem;">
            🔄 Electronic Renewal Review
          </div>
          <div style="font-size:0.9rem;opacity:0.9;margin-top:2px;">
            Uniqueness + usage + decision matrix on one page
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    page_workflow_e()


if __name__ == "__main__":
    main()
