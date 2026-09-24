"""
generate_ocm_net_report.py
Updates OCM Network KPI files (Daily, Weekly, Monthly) from vendor raw files.

Supported vendors/generations:
  Huawei  2G, 3G, 4G  + dedicated packet-loss files (H_3G_PL, H_4G_PL)
  Nokia   2G, 3G, 4G  + dedicated packet-loss files (N_3G_PL, N_4G_PL)
  ZTE     2G, 3G, 4G  + dedicated packet-loss files (Z_3G_PL, Z_4G_PL)

Sheets updated in each OCM file:
  Avail2G, Avail3G, Avail4G,
  DailyCombinedCSTrafic (Kerl), DailyCombinedPSTrafic (TB),
  CSSR2G, CSSR3G, DCR2G, DCR3G,
  3G PacketLoss, 4G PacketLoss, Throughput 4G Mbps
"""

from __future__ import annotations

import os
import re
import logging
from datetime import datetime, date
from collections import defaultdict

import openpyxl
from openpyxl.styles import PatternFill
from openpyxl.formatting.rule import FormulaRule

logger = logging.getLogger(__name__)

# ─── OCM SHEET NAMES ──────────────────────────────────────────────────────────

SHEETS = [
    'Avail2G',
    'Avail3G',
    'Avail4G',
    'DailyCombinedCSTrafic (Kerl)',
    'DailyCombinedPSTrafic (TB)',
    'CSSR2G',
    'CSSR3G',
    'DCR2G',
    'DCR3G',
    '3G PacketLoss',
    '4G PacketLoss',
    'Throughput 4G Mbps',
]

# Sheets for dedicated packet-loss-only update mode
PL_SHEETS = {'3G PacketLoss', '4G PacketLoss'}
PL_VENDOR_KEYS = {'H_3G_PL', 'H_4G_PL', 'N_3G_PL', 'N_4G_PL', 'Z_3G_PL', 'Z_4G_PL'}

# Sheets where values from multiple vendor files are SUMMED (not averaged)
ADDITIVE_SHEETS = {
    'DailyCombinedCSTrafic (Kerl)',
    'DailyCombinedPSTrafic (TB)',
}

# Sheets where values are rounded to 2 decimal places on output
TRAFFIC_SHEETS = {
    'DailyCombinedCSTrafic (Kerl)',
    'DailyCombinedPSTrafic (TB)',
}

# ─── KPI MAPPINGS ─────────────────────────────────────────────────────────────
# {sheet_name: {vendor_gen: (column_name_in_source, multiplier)}}
#
# multiplier converts the raw value to the target unit:
#   availability  → % (0-100)
#   CSSR / DCR    → % (0-100)
#   CS Traffic    → Kilo-Erlang  (raw Erl × 0.001)
#   PS Traffic    → Tera-Bytes   (raw MB × 1e-6, raw GB × 0.001)
#   Throughput    → Mbps         (raw kbps × 0.001)
#
# UNIT NOTES per vendor:
#   Huawei  2G/3G/4G  : all rates already in %; traffic in Erl/MB/kB·s⁻¹
#   Nokia   2G/3G/4G  : rates in % (rows start at index 2, index 1 is KPI codes)
#   ZTE     2G        : availability is ratio (0-1); CSSR/DCR already in %
#   ZTE     3G        : availability, CSSR, DCR all ratios (0-1) → ×100
#   ZTE     4G        : availability and rates already in %

KPI_MAP: dict[str, dict[str, tuple[str, float]]] = {

    'Avail2G': {
        'H_2G': ('RR307:TCH Availability(%)',               1.0),
        'N_2G': ('TCH Availability Normal TRXs',            1.0),    # already %
        'Z_2G': ('ORA_2G_TCH Availability Normal TRXs',    100.0),   # ratio → %
    },

    'Avail3G': {
        'H_3G': ('3G Availability (Group)',                                         1.0),
        'N_3G': ('Cell Availability, excluding blocked by user state (BLU)',        1.0),   # %
        'Z_3G': ('ORA_3G_Cell Availability, excluding BLU (%)',                   100.0),   # ratio → %
    },

    'Avail4G': {
        'H_4G': ('4G Cell Availability (Excluding manual)',         1.0),
        'N_4G': ('Cell Avail',                                      1.0),
        'Z_4G': ('ORA_4G_Cell Availability, excluding BLU_ZTE(%)', 1.0),
    },

    'DailyCombinedCSTrafic (Kerl)': {
        # Raw values in Erlang; output in Kilo-Erlang
        'H_2G': ('K3014:Traffic Volume on TCH(Erl)',     0.001),
        'H_3G': ('Grp_ 3G TRAFFIC – SPEECH(Erl)',         0.001),  # Erl → Kerl
        'N_2G': ('Erlang_Traffic_Carried_2G',            0.001),
        'N_3G': ('grp_traffic_speech_erlangs',           0.001),
        'Z_2G': ('ORA_2G_CS_TRAFFIC',                   0.001),
        'Z_3G': ('ORA_3G_CS Voice Traffic (Erl)',        0.001),
    },

    'DailyCombinedPSTrafic (TB)': {
        # Formula: (2GPS_GB + 3GPS_GB + 4GPS_GB) / 1000 → TB
        # MB sources are converted MB→GB (×0.001) then GB→TB (×0.001) = ×1e-6
        'H_2G': ('2G PS Traffic(MB)',                                  1e-6),  # MB → TB
        'H_3G': ('Grp_3G&3G+ TOTAL DATA TRAFFIC VOLUME DL&UL(GB)',    0.001), # GB → TB
        'H_4G': ('4G_Grp_4G/LTE TRAFFIC VOLUME(GB)',                  0.001), # GB → TB
        'N_2G': ('Grp_2G_Traffic_Data_Mbytes',                        1e-6),  # MB → TB
        # N_3G: confirm units for Grp_Total_Data_Traffic_Vol_DlUL
        'N_4G': ('Grp_4G_Traffic_New_Gbytes',                         0.001), # GB → TB
        'Z_2G': ('ORA_2G_Traffic_Data_GB',                            0.001), # GB → TB
        'Z_3G': ('ORA_3G_Traffic Total Data_MAC (GB)',                 0.001), # GB → TB
        'Z_4G': ('ORA_4G_Total TRAFFIC(DL+UL)(GB)',                   0.001), # GB → TB
    },

    'CSSR2G': {
        'H_2G': ('FT_Call Setup Success Rate-Speech',   1.0),
        'N_2G': ('ORA_2G_CSSR_CS_new',                  1.0),
        'Z_2G': ('ORA_2G_CSSR_CS_New(%)',               1.0),
    },

    'CSSR3G': {
        'H_3G': ('GRP_3G Call Setup Success Rate (CS)(%)',          1.0),
        'N_3G': ('ORA_CSSR_CS_CellPCH_URAPCHnew',                  1.0),
        'Z_3G': ('ORA_3G_CSSR CS with Cell PCH/URA PCH (%)',      100.0),  # ratio → %
    },

    'DCR2G': {
        'H_2G': ('FT_Drop Call Rate-Speech',            1.0),
        'N_2G': ('ORA_2G_Call_Drop_CS_new',             1.0),
        'Z_2G': ('ORA_2G_Call_Drop_CS_New(%)',          1.0),
    },

    'DCR3G': {
        'H_3G': ('Grp_3G Drop Call Rate (CS)%_Update',     1.0),
        'N_3G': ('grp_3G_Drop_Call_CS',                    1.0),
        'Z_3G': ('ORA_3G_Drop Call Rate CS(%)',           100.0),  # ratio → %
    },

    'PacketLoss_REMOVED': {},  # replaced by 3G PacketLoss and 4G PacketLoss below

    '3G PacketLoss': {
        # Huawei 3G: IP-path forward packet drop mean (%)
        'H_3G_PL': ('VS.IPPM.Forword.DropMeans(%)',    1.0),
        # Nokia 3G: HS-DSCH frame-loss credit reductions (count, not %) – included as-is
        'N_3G_PL': ('HS-DSCH CREDIT REDUCTIONS DUE TO FRAME LOSS', 1.0),
        # ZTE 3G: Packet loss enhance (%)
        'Z_3G_PL': ('Packet loss Enhance(%)',           1.0),
    },

    '4G PacketLoss': {
        # Huawei 4G: DL packet loss rate of data service (%)
        'H_4G_PL': ('DL Packet Loss Rate of Data Service(%)', 1.0),
        # Nokia 4G: computed via formula 100-(100*SCTP_S1CHSNT(M8035C0))/(SCTP_S1CHRSNT(M8035C2)+SCTP_S1CHSNT(M8035C0))
        # handled in COMPUTED_KPIS below – no direct column entry here
        # ZTE 4G: IP-path packet loss rate (%)
        'Z_4G_PL': ('Packet Loss Rate(%)',              1.0),
    },

    'Throughput 4G Mbps': {
        # Both Huawei and Nokia report in kbps despite label differences
        'H_4G': ('4G_Grp_AVG 4G/LTE DL USER THRPUT (ALL) (KBPS)(kB/s)', 0.001),  # kbps → Mbps
        'N_4G': ('LTE_DL_USER_THRPUT_ALL_MBPS_NEW',                       0.001),  # ÷1000 → Mbps
        'Z_4G': ('ORA_4G_DL_User_Throughput_New(kbps)',                   0.001),  # kbps → Mbps
    },
}

# ─── COMPUTED KPIs ────────────────────────────────────────────────────────────
# For KPIs that cannot be read directly from a single column but need a formula.
# {vendor_gen: [(sheet_name, col_a_name, col_b_name, compute_fn)]}
# compute_fn(a, b) → float | None

def _n4g_pl_formula(snt, rsnt):
    """Nokia 4G packet loss: 100 - (100 * SCTP_S1CHSNT(M8035C0)) / (SCTP_S1CHRSNT(M8035C2) + SCTP_S1CHSNT(M8035C0))"""
    a = _to_float(snt)    # SCTP_S1CHSNT  (M8035C0) — sent chunks
    b = _to_float(rsnt)   # SCTP_S1CHRSNT (M8035C2) — retransmitted chunks
    if a is None or b is None:
        return None
    denom = b + a
    if denom == 0:
        return None
    return round(100.0 - (100.0 * a) / denom, 6)

COMPUTED_KPIS: dict[str, list[tuple]] = {
    # (sheet_name, col_a_header, col_b_header, fn(a_val, b_val) → float|None)
    'N_4G_PL': [
        ('4G PacketLoss',
         'SCTP_S1CHSNT (M8035C0)',
         'SCTP_S1CHRSNT (M8035C2)',
         _n4g_pl_formula),
    ],
}

# ─── VENDOR FILE READ CONFIGS ─────────────────────────────────────────────────
# date_col       : column header containing the date
# site_col       : column header containing the site/node name
# skip_extra_hdr : Nokia files have a second header row (KPI codes) → skip it
# skip_site_re   : regex to skip aggregate/invalid site name rows

FILE_CONFIGS: dict[str, dict] = {
    'H_2G': {
        'date_col': 'Date',
        'site_col': 'Site Name',
        'skip_extra_hdr': False,
        'skip_site_re': None,
    },
    'H_3G': {
        'date_col': 'Date',
        'site_col': 'NODEBNAME',
        'skip_extra_hdr': False,
        'skip_site_re': None,
    },
    'H_4G': {
        'date_col': 'Date',
        'site_col': 'eNodeB Name',
        'skip_extra_hdr': False,
        'skip_site_re': None,
    },
    'N_2G': {
        'date_col': 'Period start time',
        'site_col': 'BCF name',
        'skip_extra_hdr': True,   # row 1 = KPI codes
        'skip_site_re': r'^\d+$', # purely numeric BCF names are non-site rows
    },
    'N_3G': {
        'date_col': 'Period start time',
        'site_col': 'WBTS name',
        'skip_extra_hdr': True,
        'skip_site_re': r'^\d+$', # e.g. '300000000' aggregate rows
    },
    'N_4G': {
        'date_col': 'Period start time',
        'site_col': 'LNBTS name',
        'skip_extra_hdr': True,
        'skip_site_re': None,
    },
    'Z_2G': {
        'date_col': 'Begin Time',
        'site_col': 'SITE Name',
        'skip_extra_hdr': False,
        'skip_site_re': None,
    },
    'Z_3G': {
        'date_col': 'Begin Time',
        'site_col': 'NodeB Name',
        'skip_extra_hdr': False,
        'skip_site_re': None,
    },
    'Z_4G': {
        # ZTE 4G uses a non-breaking space (\xa0) in the column name;
        # headers are normalised before comparison so use a regular space here.
        'date_col': 'Begin Time',
        'site_col': 'Managed Element',
        'skip_extra_hdr': False,
        'skip_site_re': None,
    },
    # ── Dedicated packet-loss files ───────────────────────────────────────────
    'H_3G_PL': {
        # Huawei 3G IP-path monitoring file
        'date_col': 'Date',
        'site_col': 'NODEBNAME',
        'skip_extra_hdr': False,
        'skip_site_re': None,
    },
    'H_4G_PL': {
        # Huawei 4G main KPIs file (contains DL Packet Loss Rate of Data Service)
        'date_col': 'Date',
        'site_col': 'eNodeB Name',
        'skip_extra_hdr': False,
        'skip_site_re': None,
    },
    'N_3G_PL': {
        # Nokia 3G packet-loss KPI file – same Nokia two-row header pattern
        'date_col': 'Period start time',
        'site_col': 'WBTS name',
        'skip_extra_hdr': True,
        'skip_site_re': r'^\d+$',
    },
    'N_4G_PL': {
        # Nokia 4G SCTP S1 counters – PL computed via COMPUTED_KPIS formula
        'date_col': 'Period start time',
        'site_col': 'LNBTS name',
        'skip_extra_hdr': True,
        'skip_site_re': None,
    },
    'Z_3G_PL': {
        # ZTE 3G IP-path monitoring file – multiple rows per site/day → averaged
        'date_col': 'Begin Time',
        'site_col': 'Office Name',
        'skip_extra_hdr': False,
        'skip_site_re': None,
    },
    'Z_4G_PL': {
        # ZTE 4G IP-path monitoring file – multiple rows per site/day → averaged
        'date_col': 'Begin Time',
        'site_col': 'Managed Element',
        'skip_extra_hdr': False,
        'skip_site_re': None,
    },
}


# ─── HELPERS ──────────────────────────────────────────────────────────────────

def _normalise(text) -> str:
    """Replace non-breaking spaces and strip; used to normalise column headers."""
    if text is None:
        return ''
    return str(text).replace('\xa0', ' ').strip()


def _to_float(val) -> float | None:
    """Return float or None for missing / Huawei '/0' / 'NIL' values."""
    if val is None:
        return None
    s = str(val).strip()
    if s in ('/0', 'NIL', 'N/A', '-', ''):
        return None
    try:
        return float(val)
    except (ValueError, TypeError):
        return None


def _parse_date(val) -> date | None:
    """Convert datetime / date / ISO-string to a date object."""
    if isinstance(val, datetime):
        return val.date()
    if isinstance(val, date):
        return val
    if isinstance(val, str):
        try:
            return datetime.strptime(val[:10], '%Y-%m-%d').date()
        except ValueError:
            pass
    return None


def _normalise_site_name(site_name: str | None) -> str:
    """
    Normalise a site name string for reliable direct comparison:
    - Replace non-breaking spaces with standard space
    - Strip leading and trailing whitespace
    - Convert to uppercase
    - Replace dashes '-' with underscores '_'
    - Collapse multiple spaces/underscores into a single underscore
    """
    if not site_name:
        return ''
    s = _normalise(site_name).upper()
    s = s.replace('-', '_')
    s = re.sub(r'[\s_]+', '_', s)
    return s.strip('_')


def _date_to_week_key(d: date) -> str:
    """ISO week label e.g. '2026W19'."""
    iso = d.isocalendar()
    return f'{iso[0]}W{iso[1]:02d}'


def _date_to_month_key(d: date) -> str:
    """Month label e.g. '2026M05'."""
    return f'{d.year}M{d.month:02d}'


# ─── VENDOR FILE READER ───────────────────────────────────────────────────────

def read_vendor_file(
    filepath: str,
    vendor_gen: str,
) -> dict[str, dict[date, dict[str, float]]]:
    """
    Parse one vendor raw file.
    Returns {site_code: {date: {sheet_name: value}}}.
    """
    cfg = FILE_CONFIGS[vendor_gen]
    logger.debug(f'  Reading {vendor_gen}: {os.path.basename(filepath)}')

    wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
    ws = wb.active
    row_iter = ws.iter_rows(values_only=True)

    # Read header row
    try:
        first_row = next(row_iter)
    except StopIteration:
        wb.close()
        logger.warning(f'  {vendor_gen}: file is empty')
        return {}
    headers = [_normalise(h) for h in first_row]

    # Nokia files have an extra KPI-code row between header and data
    if cfg['skip_extra_hdr']:
        try:
            next(row_iter)  # skip KPI codes row
        except StopIteration:
            pass

    # Locate date and site columns
    date_col = cfg['date_col']
    site_col = cfg['site_col']

    if date_col not in headers:
        logger.warning(f'  {vendor_gen}: date column "{date_col}" not found; '
                       f'available: {headers[:10]}')
        return {}
    if site_col not in headers:
        logger.warning(f'  {vendor_gen}: site column "{site_col}" not found; '
                       f'available: {headers[:10]}')
        return {}

    date_idx = headers.index(date_col)
    site_idx = headers.index(site_col)

    # Build KPI column index → (sheet, multiplier) map  (direct columns)
    kpi_cols: dict[int, tuple[str, float]] = {}
    for sheet_name, vmap in KPI_MAP.items():
        if vendor_gen in vmap:
            col_name, mult = vmap[vendor_gen]
            norm_col = _normalise(col_name)
            if norm_col in headers:
                kpi_cols[headers.index(norm_col)] = (sheet_name, mult)
            else:
                logger.warning(f'  {vendor_gen}: KPI column "{col_name}" '
                               f'not found for sheet "{sheet_name}"')

    # Build computed KPI specs (e.g. Nokia 4G PL formula)
    computed_specs = []
    for sheet_name, col_a, col_b, fn in COMPUTED_KPIS.get(vendor_gen, []):
        na, nb = _normalise(col_a), _normalise(col_b)
        if na in headers and nb in headers:
            computed_specs.append((sheet_name, headers.index(na), headers.index(nb), fn))
        else:
            missing = col_a if na not in headers else col_b
            logger.warning(f'  {vendor_gen}: computed KPI column "{missing}" '
                           f'not found for sheet "{sheet_name}"')

    skip_re = re.compile(cfg['skip_site_re']) if cfg.get('skip_site_re') else None

    # Use accumulator to handle files with multiple rows per site/day
    # (e.g. ZTE packet-loss file with one row per IP path).
    # Structure: {site_name: {date: {sheet_name: [values]}}}
    accum: dict[str, dict[date, dict[str, list[float]]]] = defaultdict(
        lambda: defaultdict(lambda: defaultdict(list))
    )
    skipped = 0

    for row in row_iter:
        if not row or all(v is None for v in row):
            continue

        site_raw = row[site_idx] if site_idx < len(row) else None
        if not site_raw:
            continue
        site_name = str(site_raw).strip()
        if skip_re and skip_re.match(site_name):
            skipped += 1
            continue

        d = _parse_date(row[date_idx] if date_idx < len(row) else None)
        if not d:
            continue

        for col_idx, (sheet_name, mult) in kpi_cols.items():
            raw = row[col_idx] if col_idx < len(row) else None
            val = _to_float(raw)
            if val is not None:
                accum[site_name][d][sheet_name].append(round(val * mult, 6))

        for sheet_name, idx_a, idx_b, fn in computed_specs:
            val_a = row[idx_a] if idx_a < len(row) else None
            val_b = row[idx_b] if idx_b < len(row) else None
            val = fn(val_a, val_b)
            if val is not None:
                accum[site_name][d][sheet_name].append(val)

    wb.close()

    # Collapse accumulator: average multiple values for the same site/date/sheet
    result: dict[str, dict[date, dict[str, float]]] = {}
    for site_name, date_map in accum.items():
        result[site_name] = {}
        for d, sheet_map in date_map.items():
            result[site_name][d] = {
                s: round(sum(vals) / len(vals), 6)
                for s, vals in sheet_map.items()
            }

    logger.info(f'  {vendor_gen}: {len(result)} sites | {skipped} rows skipped')
    return result


# ─── AGGREGATION ──────────────────────────────────────────────────────────────

def _aggregate(
    daily: dict[str, dict[date, dict[str, float]]],
    key_fn,
) -> dict[str, dict[str, dict[str, float]]]:
    """
    Aggregate daily data to weekly or monthly using key_fn(date) → period_key.
    Additive sheets (traffic volumes) are summed; all others are averaged.
    Returns {site_code: {period_key: {sheet: aggregated_value}}}.
    """
    buckets: dict[str, dict[str, dict[str, list[float]]]] = defaultdict(
        lambda: defaultdict(lambda: defaultdict(list))
    )
    for site_name, date_map in daily.items():
        for d, sheet_map in date_map.items():
            pk = key_fn(d)
            for sheet, val in sheet_map.items():
                buckets[site_name][pk][sheet].append(val)

    result: dict[str, dict[str, dict[str, float]]] = {}
    for site_name, pk_map in buckets.items():
        result[site_name] = {}
        for pk, sheet_map in pk_map.items():
            result[site_name][pk] = {
                s: round(sum(vals), 4) if s in ADDITIVE_SHEETS
                   else round(sum(vals) / len(vals), 4)
                for s, vals in sheet_map.items()
            }
    return result


def aggregate_weekly(daily):
    return _aggregate(daily, _date_to_week_key)


def aggregate_monthly(daily):
    return _aggregate(daily, _date_to_month_key)


# ─── CONDITIONAL FORMATTING ───────────────────────────────────────────────────

_RED_FILL   = PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid')
_GREEN_FILL = PatternFill(start_color='00B050', end_color='00B050', fill_type='solid')

# CF rules per sheet:
#   3G PacketLoss : HW >=1 → Red else Green; ZTE >=0.01 → Red else Green; Nokia → no color
#   4G PacketLoss : HW/Nokia >=1 → Red else Green; ZTE >=0.01 → Red else Green
_PL_CF_RULES = {
    '3G PacketLoss': [
        # (vendor_in_col_C, threshold, operator)  — Nokia intentionally absent
        ('HUAWEI', 1,    '>='),
        ('ZTE',    0.01, '>='),
    ],
    '4G PacketLoss': [
        ('HUAWEI', 1,    '>='),
        ('NOKIA',  1,    '>='),
        ('ZTE',    0.01, '>='),
    ],
}


def _apply_pl_cf(ws, sheet_name: str, last_data_col: int) -> None:
    """Apply conditional formatting to a PacketLoss sheet."""
    import openpyxl.utils as xl_utils
    rules = _PL_CF_RULES.get(sheet_name, [])
    if not rules:
        return

    # Range: col I (9) → last date col, rows 3 → last data row
    last_row = ws.max_row or 3
    first_col_letter = xl_utils.get_column_letter(9)
    last_col_letter  = xl_utils.get_column_letter(last_data_col)
    cf_range = f'{first_col_letter}3:{last_col_letter}{last_row}'

    # Clear any existing CF on this sheet to avoid accumulation on repeated runs
    ws.conditional_formatting = openpyxl.worksheet.worksheet.ConditionalFormattingList()

    for vendor, threshold, op in rules:
        # Red: cell meets threshold for this vendor
        ws.conditional_formatting.add(
            cf_range,
            FormulaRule(
                formula=[f'AND($C3="{vendor}",{first_col_letter}3{op}{threshold})'],
                fill=_RED_FILL,
                stopIfTrue=True,
            ),
        )
        # Green: cell does NOT meet threshold for this vendor (and has a value)
        inv_op = '<' if op == '>=' else '>'
        ws.conditional_formatting.add(
            cf_range,
            FormulaRule(
                formula=[f'AND($C3="{vendor}",{first_col_letter}3{inv_op}{threshold},{first_col_letter}3<>"")'],
                fill=_GREEN_FILL,
            ),
        )


# ─── OCM FILE WRITER ──────────────────────────────────────────────────────────

def update_ocm_file(
    filepath: str,
    data: dict,          # {site_code: {period_key: {sheet: value}}}
    period: str,         # 'daily' | 'weekly' | 'monthly'
    log_fn=None,
    sheets_filter: set[str] | None = None,  # if set, only write these sheets
) -> None:
    """
    Write KPI values into the OCM Excel workbook.

    OCM file layout (all three granularities):
      Row 1  (1-based) : empty / title row
      Row 2  (1-based) : header — site metadata in cols A-H, period labels from col I
      Row 3+ (1-based) : one site per row
      Col A             : 'Nom du Site'  (used for direct full site name lookup)
    """
    def _log(msg):
        if log_fn:
            log_fn(msg)
        else:
            logger.info(msg)

    _log(f'Loading: {os.path.basename(filepath)}')
    wb = openpyxl.load_workbook(filepath)

    sheets_to_write = [s for s in SHEETS
                       if (sheets_filter is None or s in sheets_filter)]

    for sheet_name in sheets_to_write:
        if sheet_name not in wb.sheetnames:
            _log(f'  WARNING: sheet "{sheet_name}" missing – skipped')
            continue

        ws = wb[sheet_name]

        # ── Build period_key → column-number map from header row (row 2) ──────
        period_col: dict = {}
        header_row = ws[2]
        for cell in header_row:
            hv = cell.value
            if hv is None:
                continue
            if period == 'daily':
                if isinstance(hv, datetime):
                    period_col[hv.date()] = cell.column
                elif isinstance(hv, date):
                    period_col[hv] = cell.column
                elif isinstance(hv, str):
                    d = _parse_date(hv)
                    if d:
                        period_col[d] = cell.column
            elif isinstance(hv, str):
                if period == 'weekly' and re.match(r'^\d{4}W\d{2}$', hv):
                    period_col[hv] = cell.column
                elif period == 'monthly' and re.match(r'^\d{4}M\d{2}$', hv):
                    period_col[hv] = cell.column

        # ── Build site_name → row-number map from column A (Nom du Site) ─────
        site_row: dict[str, int] = {}
        for row in ws.iter_rows(min_row=3, max_col=1, values_only=False):
            site_cell = row[0]  # column A (0-indexed = index 0)
            if site_cell.value:
                norm_name = _normalise_site_name(str(site_cell.value))
                if norm_name:
                    site_row[norm_name] = site_cell.row

        # ── Write values ───────────────────────────────────────────────────────
        written = 0
        for site_name, period_map in data.items():
            norm_key = _normalise_site_name(site_name)
            if norm_key not in site_row:
                continue
            rn = site_row[norm_key]
            for pk, sheet_map in period_map.items():
                if sheet_name not in sheet_map:
                    continue
                if pk not in period_col:
                    continue
                cn = period_col[pk]
                decimals = 2 if sheet_name in TRAFFIC_SHEETS else 4
                ws.cell(row=rn, column=cn, value=round(sheet_map[sheet_name], decimals))
                written += 1

        _log(f'  {sheet_name}: {written} cells written')

        # ── Conditional formatting for PacketLoss sheets ───────────────────────
        if sheet_name in ('3G PacketLoss', '4G PacketLoss') and period_col:
            _apply_pl_cf(ws, sheet_name, max(period_col.values()))

    wb.save(filepath)
    _log(f'Saved: {os.path.basename(filepath)}')


# Default OCM filenames (used by CLI and as default browse hints)
OCM_DEFAULT_FILENAMES = {
    'daily':   'OCM Network Key Performances Indicators Sites Per Sites Daily Level.xlsx',
    'weekly':  'OCM Network Key Performances Indicators Sites Per Sites Weekly Level.xlsx',
    'monthly': 'OCM Network Key Performances Indicators Sites Per Sites Monthly Level.xlsx',
}

ALL_GRANULARITIES = {'daily', 'weekly', 'monthly'}


# ─── MAIN ENTRY POINT ─────────────────────────────────────────────────────────

def process_vendor_files(
    vendor_files: dict[str, str],
    ocm_files: dict[str, str],
    granularities: set[str] | None = None,
    sheets_filter: set[str] | None = None,
    log_callback=None,
) -> None:
    """
    Main entry point called by the GUI.

    vendor_files  : {vendor_gen: filepath}  e.g. {'H_2G': '...', 'N_3G': '...'}
    ocm_files     : {'daily': path, 'weekly': path, 'monthly': path}
                    Each key is optional – only provide the ones you want to update.
    granularities : subset of {'daily', 'weekly', 'monthly'}.
                    Defaults to all keys present in ocm_files.
    sheets_filter : if set, only these OCM sheets are written and only vendor keys
                    that map to these sheets are read. Use PL_SHEETS for PL-only mode.
    log_callback  : optional callable(str) for UI progress messages.
    """
    if granularities is None:
        granularities = set(ocm_files.keys())
    granularities = granularities & ALL_GRANULARITIES

    # Determine which vendor keys are relevant for sheets_filter
    if sheets_filter is not None:
        relevant_vkeys: set[str] = set()
        for sheet, vmap in KPI_MAP.items():
            if sheet in sheets_filter:
                relevant_vkeys.update(vmap.keys())
        for vg, specs in COMPUTED_KPIS.items():
            for sheet_name, *_ in specs:
                if sheet_name in sheets_filter:
                    relevant_vkeys.add(vg)
        vendor_files = {vg: fp for vg, fp in vendor_files.items()
                        if vg in relevant_vkeys}

    def log(msg: str):
        if log_callback:
            log_callback(msg)
        else:
            logger.info(msg)

    if not granularities:
        log('No granularities selected – nothing to do.')
        return

    # ── Validate OCM file paths ───────────────────────────────────────────────
    for gran in granularities:
        fp = ocm_files.get(gran, '')
        if not fp or not os.path.isfile(fp):
            log(f'ERROR: OCM {gran} file not found: {fp}')
            return

    # ── Collect all daily data ────────────────────────────────────────────────
    # Structure: {site_code: {date: {sheet: value}}}
    all_daily: dict[str, dict[date, dict[str, float]]] = defaultdict(
        lambda: defaultdict(dict)
    )

    for vg, fp in vendor_files.items():
        if not fp or not os.path.isfile(fp):
            log(f'  Skipping {vg}: no file provided')
            continue
        log(f'Reading {vg} …')
        try:
            vdata = read_vendor_file(fp, vg)
            for site_code, date_map in vdata.items():
                for d, sheet_map in date_map.items():
                    for sheet, val in sheet_map.items():
                        existing = all_daily[site_code][d].get(sheet)
                        if existing is not None and sheet in ADDITIVE_SHEETS:
                            all_daily[site_code][d][sheet] = round(existing + val, 6)
                        else:
                            all_daily[site_code][d][sheet] = val
        except Exception as exc:
            import traceback
            log(f'  ERROR reading {vg}: {exc}')
            log(traceback.format_exc())

    total_sites = len(all_daily)
    if total_sites == 0:
        log('No data loaded – nothing to write.')
        return
    log(f'Total sites with data: {total_sites}')

    # ── Update selected granularities ─────────────────────────────────────────
    if 'daily' in granularities:
        log('\n--- Updating Daily file ---')
        update_ocm_file(ocm_files['daily'], all_daily, period='daily',
                        log_fn=log_callback, sheets_filter=sheets_filter)

    if 'weekly' in granularities:
        log('\n--- Aggregating to Weekly ---')
        weekly_data = aggregate_weekly(all_daily)
        update_ocm_file(ocm_files['weekly'], weekly_data, period='weekly',
                        log_fn=log_callback, sheets_filter=sheets_filter)

    if 'monthly' in granularities:
        log('\n--- Aggregating to Monthly ---')
        monthly_data = aggregate_monthly(all_daily)
        update_ocm_file(ocm_files['monthly'], monthly_data, period='monthly',
                        log_fn=log_callback, sheets_filter=sheets_filter)

    done = ' + '.join(sorted(granularities))
    log(f'\nDone. Updated: {done}.')


# ─── CLI ──────────────────────────────────────────────────────────────────────

if __name__ == '__main__':
    import argparse

    logging.basicConfig(level=logging.INFO, format='%(message)s')

    parser = argparse.ArgumentParser(description='Update OCM Network KPI reports from vendor raw files.')
    parser.add_argument('--ocm-dir', default=os.path.join('files', 'ocm net'),
                        help='Folder containing the three OCM Excel files')
    parser.add_argument('--raw-dir', default=os.path.join('files', 'ocm net', 'RAW FILES'),
                        help='Folder containing vendor raw files')

    args = parser.parse_args()

    # Auto-detect files by filename prefix
    # IMPORTANT: more-specific prefixes must come before overlapping ones
    # (e.g. 'H PACKET' before 'H P', 'Z PACKET' before 'Z P')
    prefixes = {
        'H_2G':    ('H 2G', 'H2G'),
        'H_3G':    ('H 3G', 'H3G'),
        'H_4G':    ('H 4G', 'H4G'),
        'H_3G_PL': ('HUAWEI 3G PACKET',),
        'H_4G_PL': ('HUAWEI 4G PACKET',),
        'N_2G':    ('N 2G', 'N2G'),
        'N_3G':    ('N 3G', 'N3G'),
        'N_4G':    ('N 4G', 'N4G'),
        'N_3G_PL': ('NOKIA 3G PACKET',),
        'N_4G_PL': ('NOKIA 4G PACKET',),
        'Z_3G_PL': ('ZTE 3G PACKET',),
        'Z_4G_PL': ('ZTE 4G PACKET',),
        'Z_2G':    ('Z P',),           # 'Z Performance Management...2G'
        'Z_3G':    ('Z 3G', 'Z3G'),
        'Z_4G':    ('Z-4G', 'Z 4G', 'Z4G'),
    }

    vendor_files: dict[str, str] = {}
    for fname in os.listdir(args.raw_dir):
        if not fname.lower().endswith('.xlsx'):
            continue
        for vg, pfxs in prefixes.items():
            if any(fname.startswith(p) for p in pfxs) and vg not in vendor_files:
                vendor_files[vg] = os.path.join(args.raw_dir, fname)
                break

    logger.info('Detected vendor files:')
    for vg, fp in vendor_files.items():
        logger.info(f'  {vg}: {os.path.basename(fp)}')

    ocm_files = {
        gran: os.path.join(args.ocm_dir, fname)
        for gran, fname in OCM_DEFAULT_FILENAMES.items()
    }
    process_vendor_files(vendor_files, ocm_files)
