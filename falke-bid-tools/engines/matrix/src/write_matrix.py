"""
FALKE Matrix Pipeline — Excel Writer (Fresh Workbook)
======================================================
Writes normalized bid data into a brand-new openpyxl Workbook().

No template is loaded or copied — the FEB 26 file's merged cells, conditional
formatting, and drawing objects caused corruption on save.  This module
creates a clean xlsx that mirrors the FEB 26 row/column structure closely
enough for side-by-side value comparison.

Pipeline position:
    list[NormalizedBid]  →  write_matrix()  →  bid-comparison .xlsx

Layout:
  Col A  : CSI code / label
  Col B  : Row description
  Col C+ : Contractor groups, each 3 columns wide:
             +0  COST SUBTOTALS (main comparison number)
             +1  $/SF           (subtotal ÷ gsf)
             +2  blank separator

  Row 1  : Project title
  Row 2  : Project details
  Row 3  : blank
  Row 4  : Column headers (CSI / Building System / COST SUBTOTALS / $/SF / …)
  Row 5  : Contractor names
  Row 6  : Project label per contractor
  Row 7  : GSF per contractor
  Row 8  : blank
  Rows 9+: CSI division data
  …      : blank separator, footer section, blank separator, qualifications
"""

from __future__ import annotations

from decimal import Decimal
from pathlib import Path
from typing import Optional

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill

from src.audit import AuditItem, AuditStatus
from src.normalized_models import CellState, NormalizedBid, NormalizedDivision
from src.run_config import RunInputs

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

# Project identity is a PER-RUN input (RunInputs), never hardcoded. These
# fallbacks exist only so a programmatic caller that passes a gsf but no
# RunInputs still produces a generic, non-client-specific title.
DEFAULT_SF_BASIS_LABEL = "GSF"

# Column widths
COL_A_WIDTH: float = 15.0
COL_B_WIDTH: float = 45.0
CONTRACTOR_COST_WIDTH: float = 18.0
CONTRACTOR_SF_WIDTH: float = 10.0
CONTRACTOR_SEP_WIDTH: float = 4.0

# Number format for dollar amounts
AMOUNT_FORMAT = "#,##0.00"

# ---------------------------------------------------------------------------
# Division and footer row definitions (FEB 26 CSI sequence)
# ---------------------------------------------------------------------------

DIVISION_ROWS: list[tuple[str, str]] = [
    ("DIV 01 00 00", "General Requirements"),
    ("DIV 02 00 00", "Existing Conditions"),
    ("DIV 03 00 00", "Concrete"),
    ("DIV 04 00 00", "Masonry"),
    ("DIV 05 00 00", "Metals"),
    ("DIV 06 00 00", "Wood, Plastics & Composites"),
    ("DIV 07 00 00", "Thermal & Moisture Protection"),
    ("DIV 08 00 00", "Openings"),
    ("DIV 09 00 00", "Finishes"),
    ("DIV 10 00 00", "Specialties"),
    ("DIV 11 00 00", "Equipment"),
    ("DIV 12 00 00", "Furnishings"),
    ("DIV 13 00 00", "Special Construction"),
    ("DIV 21 00 00", "Fire Suppression"),
    ("DIV 22 00 00", "Plumbing"),
    ("DIV 23 00 00", "HVAC"),
    ("DIV 25 00 00", "Integrated Automation"),
    ("DIV 26 00 00", "Electrical"),
    ("DIV 27 00 00", "Communications"),
    ("DIV 28 00 00", "Electronic Safety & Security"),
]

FOOTER_ROWS: list[tuple[str, str]] = [
    ("CONSTRUCTION_SUBTOTAL", "Construction Cost Subtotal"),
    ("GL_INSURANCE",          "General Liability Insurance"),
    ("BUILDERS_RISK",         "Builders Risk Insurance"),
    ("GC_FEE",                "GC Fee / OH&P"),
    ("FEES_SUBTOTAL",         "Fees Subtotal"),
    ("GRAND_TOTAL",           "GRAND TOTAL"),
    ("BOND",                  "Bond (Alternate)"),
]

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _to_float(amount: Optional[Decimal]) -> float:
    """Convert a Decimal (or None) to float. Returns 0.0 for None."""
    if amount is None:
        return 0.0
    return float(amount)


def _cell_amount(state: CellState, amount: Optional[Decimal]) -> float:
    """
    Return the numeric value to write for a CellValue, given its state.

    AMOUNT / EXPLICIT_ZERO / ALLOWANCE → float (or 0.0)
    NULL_BLANK / EXCLUDED / BY_OWNER_OTHERS → 0.0
    """
    if state in (CellState.AMOUNT, CellState.EXPLICIT_ZERO, CellState.ALLOWANCE):
        return _to_float(amount)
    return 0.0


def _sort_bids(bids: list[NormalizedBid]) -> list[NormalizedBid]:
    """
    Order columns by leveled_total ASCENDING — lowest leveled bid first, the
    natural board reading order, fully firm-agnostic (Marvin §2.5, Floyd C7).
    A bid with leveled_total None (e.g. no grand total extracted) sorts LAST so
    it never masquerades as the low bid. Ties break on contractor_name for
    determinism.
    """
    def _key(b: NormalizedBid) -> tuple[int, float, str]:
        lt = b.footer.leveled_total
        if lt is None:
            return (1, 0.0, b.contractor_name)
        return (0, float(lt), b.contractor_name)

    return sorted(bids, key=_key)


def _col_start(contractor_index: int) -> int:
    """
    Return the 1-based openpyxl column index for the COST SUBTOTALS column
    of the given contractor (0-based index).

    Layout: col A(1)=CSI, col B(2)=description, col C(3)=contractor 0 start.
    Each contractor group is 3 columns wide.
    """
    return 3 + contractor_index * 3  # C=3, F=6, I=9, L=12 …


# ---------------------------------------------------------------------------
# Worksheet construction
# ---------------------------------------------------------------------------

def _write_header_rows(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    bids: list[NormalizedBid],
    gsf: int,
    run: RunInputs,
) -> None:
    """Write rows 1–8: title, details, blank, column headers, contractor info.

    Project identity (title, details, per-contractor label) comes from the
    per-run RunInputs — never a hardcoded project (M1). The $/SF header carries
    the confirmed SF-basis label so the board knows what the denominator means
    (M2 / scoping §1.4).
    """
    bold = Font(bold=True)
    sf_label = run.sf_basis_label or DEFAULT_SF_BASIS_LABEL

    # Row 1 — project title
    ws.cell(row=1, column=1).value = f"{run.project_name} — Bid Comparison Matrix"
    ws.cell(row=1, column=1).font = bold

    # Row 2 — project details
    details = f"Project: {run.project_name} | {run.project_address} | {gsf:,.0f} {sf_label}"
    if run.rfp_label:
        details += f" | {run.rfp_label}"
    ws.cell(row=2, column=1).value = details

    # Row 3 — blank (intentional)

    # Row 4 — column headers
    ws.cell(row=4, column=1).value = "CSI"
    ws.cell(row=4, column=1).font = bold
    ws.cell(row=4, column=2).value = "Building System"
    ws.cell(row=4, column=2).font = bold

    for i, bid in enumerate(bids):
        cost_col = _col_start(i)
        sf_col = cost_col + 1
        ws.cell(row=4, column=cost_col).value = "COST SUBTOTALS"
        ws.cell(row=4, column=cost_col).font = bold
        ws.cell(row=4, column=sf_col).value = f"$/{sf_label}"
        ws.cell(row=4, column=sf_col).font = bold

    # Row 5 — contractor names
    for i, bid in enumerate(bids):
        cost_col = _col_start(i)
        ws.cell(row=5, column=cost_col).value = bid.contractor_name
        ws.cell(row=5, column=cost_col).font = bold

    # Row 6 — project label per contractor
    for i in range(len(bids)):
        cost_col = _col_start(i)
        ws.cell(row=6, column=cost_col).value = run.project_name

    # Row 7 — SF basis per contractor
    for i in range(len(bids)):
        cost_col = _col_start(i)
        ws.cell(row=7, column=cost_col).value = gsf

    # Row 8 — blank (intentional)


def _descriptions_match(a: str, b: str) -> bool:
    """
    Return True when two line-item description strings are semantically the same.
    Uses word overlap ≥ 60% (no external libraries needed).
    """
    a_words = set(a.lower().split())
    b_words = set(b.lower().split())
    if not a_words or not b_words:
        return False
    overlap = len(a_words & b_words) / max(len(a_words), len(b_words))
    return overlap >= 0.6


def _lookup_item_amount(
    div: "NormalizedDivision | None",
    target_desc: str,
) -> Optional[float]:
    """
    Return the float amount for target_desc from a NormalizedDivision's
    line_item_cells.  Tries exact key first, then _descriptions_match.
    Returns None when no match or state maps to 0.0/blank.
    """
    if div is None:
        return None

    # Exact match first
    if target_desc in div.line_item_cells:
        cell = div.line_item_cells[target_desc]
        if cell.state in (CellState.AMOUNT, CellState.EXPLICIT_ZERO, CellState.ALLOWANCE):
            return _to_float(cell.amount)
        return None

    # Fuzzy match
    for key, cell in div.line_item_cells.items():
        if _descriptions_match(target_desc, key):
            if cell.state in (CellState.AMOUNT, CellState.EXPLICIT_ZERO, CellState.ALLOWANCE):
                return _to_float(cell.amount)
            return None

    return None


def _build_unified_descriptions(
    bids: list[NormalizedBid],
    csi_code: str,
) -> list[str]:
    """
    For a given CSI division, collect all unique line-item descriptions across
    all contractors in display order.

    Algorithm: iterate contractors in order; for each contractor's line_item_cells
    append a description only if no existing entry in the running list matches it
    (case-insensitive substring OR ≥60% word overlap).

    LUMP_SUM contractors with no real items (all NULL_BLANK) are skipped —
    their placeholder descriptions must not pollute the unified list when they
    carry no pricing signal.
    """
    all_descs: list[str] = []

    for bid in bids:
        div = _find_div(bid, csi_code)
        if div is None:
            continue

        # Skip divisions where all line items are NULL_BLANK (lump-sum placeholder rows)
        has_priced_item = any(
            cell.state in (CellState.AMOUNT, CellState.EXPLICIT_ZERO, CellState.ALLOWANCE)
            for cell in div.line_item_cells.values()
        )
        if not has_priced_item and div.cost_structure.value == "LUMP_SUM":
            continue

        for desc in div.line_item_cells:
            already_present = any(
                _descriptions_match(desc, existing) or desc.lower() in existing.lower()
                or existing.lower() in desc.lower()
                for existing in all_descs
            )
            if not already_present:
                all_descs.append(desc)

    return all_descs


def _find_div(bid: NormalizedBid, csi_code: str) -> "NormalizedDivision | None":
    """Return the first NormalizedDivision matching csi_code for a bid, or None."""
    for div in bid.divisions:
        if div.csi_code == csi_code:
            return div
    return None


def _write_division_rows(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    bids: list[NormalizedBid],
    start_row: int,
    gsf: int,
) -> tuple[int, dict[str, int]]:
    """
    Write dynamic division rows starting at start_row.

    For each CSI division:
      1. Division header row (bold): CSI code | Division name
      2. One row per unified line-item description (across all contractors)
      3. Bold subtotal row: blank | DIVISION NAME SUBTOTAL | sub amounts
      4. Blank spacer row

    Returns (next_row_after_divisions, subtotal_row_by_csi_code).
    subtotal_row_by_csi_code maps each CSI code to the Excel row number of its
    SUBTOTAL row — used by the caller to apply audit-driven color fills.
    """
    bold = Font(bold=True)

    # Build per-bid, per-division lookups:
    #   bid_div_lookup[bid_index][csi_code] = aggregated subtotal float
    bid_div_lookup: list[dict[str, float]] = []
    for bid in bids:
        aggregated: dict[str, float] = {}
        for div in bid.divisions:
            amount = _cell_amount(div.subtotal_cell.state, div.subtotal_cell.amount)
            aggregated[div.csi_code] = aggregated.get(div.csi_code, 0.0) + amount
        bid_div_lookup.append(aggregated)

    row = start_row
    subtotal_row_by_csi: dict[str, int] = {}

    for csi_code, div_name in DIVISION_ROWS:
        # --- Division header row (bold) ---
        c_csi = ws.cell(row=row, column=1)
        c_csi.value = csi_code
        c_csi.font = bold

        c_name = ws.cell(row=row, column=2)
        c_name.value = div_name
        c_name.font = bold

        row += 1

        # --- Gather per-bid NormalizedDivision objects for this CSI code ---
        bid_divs: list["NormalizedDivision | None"] = [
            _find_div(bid, csi_code) for bid in bids
        ]

        # --- Unified description list across all contractors ---
        all_descs = _build_unified_descriptions(bids, csi_code)

        # --- One row per line-item description ---
        for desc in all_descs:
            ws.cell(row=row, column=2).value = desc

            for i, div in enumerate(bid_divs):
                cost_col = _col_start(i)
                amount = _lookup_item_amount(div, desc)
                if amount is not None:
                    c = ws.cell(row=row, column=cost_col)
                    c.value = amount
                    c.number_format = AMOUNT_FORMAT

            row += 1

        # --- Subtotal row (bold) ---
        subtotal_label = div_name.upper() + " SUBTOTAL"
        c_sub_label = ws.cell(row=row, column=2)
        c_sub_label.value = subtotal_label
        c_sub_label.font = bold

        subtotal_row_by_csi[csi_code] = row  # record for audit-fill pass

        for i, bid in enumerate(bids):
            cost_col = _col_start(i)
            sf_col = cost_col + 1
            amount = bid_div_lookup[i].get(csi_code, 0.0)
            sf_val = round(amount / gsf, 2) if gsf > 0 else 0.0

            c_cost = ws.cell(row=row, column=cost_col)
            c_cost.value = amount
            c_cost.number_format = AMOUNT_FORMAT
            c_cost.font = bold

            c_sf = ws.cell(row=row, column=sf_col)
            c_sf.value = sf_val
            c_sf.number_format = AMOUNT_FORMAT
            c_sf.font = bold

        row += 1

        # --- Blank spacer ---
        row += 1

    return row, subtotal_row_by_csi  # next row after all divisions


def _write_footer_rows(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    bids: list[NormalizedBid],
    start_row: int,
    gsf: int,
) -> tuple[int, list[dict]]:
    """
    Write a blank separator then the footer rows (construction subtotal,
    insurance, fees, grand total, bond).

    FEES_SUBTOTAL is computed in Python as gl + br + gc_fee.

    Returns (next_row_after_footer, list[per_bid_footer_summary]).
    """
    bold = Font(bold=True)
    row = start_row + 1  # +1 blank separator

    summaries: list[dict] = []
    for bid in bids:
        summaries.append({})

    for key, label in FOOTER_ROWS:
        ws.cell(row=row, column=1).value = key
        ws.cell(row=row, column=2).value = label
        if key == "GRAND_TOTAL":
            ws.cell(row=row, column=2).font = bold

        for i, bid in enumerate(bids):
            cost_col = _col_start(i)
            footer = bid.footer

            if key == "CONSTRUCTION_SUBTOTAL":
                val = _cell_amount(footer.construction_subtotal.state,
                                   footer.construction_subtotal.amount)
            elif key == "GL_INSURANCE":
                val = _cell_amount(footer.general_liability_insurance.state,
                                   footer.general_liability_insurance.amount)
            elif key == "BUILDERS_RISK":
                val = _cell_amount(footer.builders_risk_insurance.state,
                                   footer.builders_risk_insurance.amount)
            elif key == "GC_FEE":
                val = _cell_amount(footer.gc_fee.state, footer.gc_fee.amount)
            elif key == "FEES_SUBTOTAL":
                # Compute in Python — no formulas
                gl  = summaries[i].get("GL_INSURANCE", 0.0)
                br  = summaries[i].get("BUILDERS_RISK", 0.0)
                gc  = summaries[i].get("GC_FEE", 0.0)
                val = gl + br + gc
            elif key == "GRAND_TOTAL":
                val = _cell_amount(footer.grand_total.state, footer.grand_total.amount)
            elif key == "BOND":
                val = _cell_amount(footer.bond.state, footer.bond.amount)
            else:
                val = 0.0

            summaries[i][key] = val

            c = ws.cell(row=row, column=cost_col)
            c.value = val
            c.number_format = AMOUNT_FORMAT
            if key == "GRAND_TOTAL":
                c.font = bold

        row += 1

    return row, summaries


def _write_alternates(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    bids: list[NormalizedBid],
    start_row: int,
) -> int:
    """Write bid alternates in their OWN clearly-labeled section (M7).

    Alternates (add/deduct options) are NEVER folded into the base/leveled
    total — the base comparison stays apples-to-apples. Each contractor's
    alternates are listed under their column. If no bidder submitted any
    alternate, the section is omitted entirely. Returns the next free row.
    """
    if not any(bid.footer.alternates for bid in bids):
        return start_row

    bold = Font(bold=True)
    row = start_row + 1  # blank separator

    ws.cell(row=row, column=1).value = "ALTERNATES"
    ws.cell(row=row, column=1).font = bold
    ws.cell(row=row, column=2).value = (
        "Bid Alternates (add/deduct — NOT included in base comparison)"
    )
    ws.cell(row=row, column=2).font = bold
    row += 1

    # One row per (contractor, alternate). Description in col B, amount under the
    # contractor's cost column — kept visually separate from the base divisions.
    for i, bid in enumerate(bids):
        cost_col = _col_start(i)
        for alt in bid.footer.alternates:
            ws.cell(row=row, column=1).value = bid.contractor_name
            ws.cell(row=row, column=2).value = alt.description
            c = ws.cell(row=row, column=cost_col)
            if alt.amount is not None:
                c.value = float(alt.amount)
                c.number_format = AMOUNT_FORMAT
            else:
                c.value = alt.display
            row += 1

    return row


def _write_qualifications(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    bids: list[NormalizedBid],
    start_row: int,
) -> None:
    """Write one qualifications row per contractor, separated by a blank row."""
    row = start_row + 1  # +1 blank separator

    ws.cell(row=row, column=1).value = "QUALIFICATIONS"
    ws.cell(row=row, column=1).font = Font(bold=True)
    ws.cell(row=row, column=2).value = "Contractor Qualifications"
    ws.cell(row=row, column=2).font = Font(bold=True)

    row += 1
    for i, bid in enumerate(bids):
        cost_col = _col_start(i)
        ws.cell(row=row, column=cost_col - 1).value = bid.contractor_name
        qual_cell = ws.cell(row=row, column=cost_col)
        qual_cell.value = bid.qualifications_text or ""
        qual_cell.alignment = Alignment(wrap_text=True)


def _set_column_widths(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    num_contractors: int,
) -> None:
    """Set column widths for readability."""
    from openpyxl.utils import get_column_letter

    ws.column_dimensions["A"].width = COL_A_WIDTH
    ws.column_dimensions["B"].width = COL_B_WIDTH

    for i in range(num_contractors):
        cost_col = _col_start(i)
        sf_col   = cost_col + 1
        sep_col  = cost_col + 2

        ws.column_dimensions[get_column_letter(cost_col)].width = CONTRACTOR_COST_WIDTH
        ws.column_dimensions[get_column_letter(sf_col)].width   = CONTRACTOR_SF_WIDTH
        ws.column_dimensions[get_column_letter(sep_col)].width  = CONTRACTOR_SEP_WIDTH


# ---------------------------------------------------------------------------
# Audit sheet fills and AUDIT worksheet
# ---------------------------------------------------------------------------

# Cell fill constants — PatternFill is safe on fresh workbooks (no existing styles to corrupt)
RED_FILL    = PatternFill("solid", fgColor="FFCCCC")   # soft red
YELLOW_FILL = PatternFill("solid", fgColor="FFF2CC")   # soft yellow
GREEN_FILL  = PatternFill("solid", fgColor="CCFFCC")   # soft green

_STATUS_FILL = {
    AuditStatus.RED:    RED_FILL,
    AuditStatus.YELLOW: YELLOW_FILL,
    AuditStatus.GREEN:  GREEN_FILL,
}

_STATUS_SORT_KEY = {
    AuditStatus.RED:    0,
    AuditStatus.YELLOW: 1,
    AuditStatus.GREEN:  2,
}


def _apply_audit_fills(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    bids: list[NormalizedBid],
    subtotal_row_by_csi: dict[str, int],
    audit_items: list[AuditItem],
) -> None:
    """
    Color-code SUBTOTAL cells in the Bid_Form sheet based on audit findings.

    A division subtotal cell is colored by the WORST status of any division-
    scoped AuditItem for that (contractor, division): RED wins over YELLOW;
    GREEN triggers no fill. This now includes the C4-promoted remap/reclass/
    unrecognized/unmatched codes (UNRECOGNIZED_CODE_FORMAT, CODE_SPLIT_UNMATCHED
    → RED; CODE_FORMAT_REMAPPED, KNOWN_FIRM_RECLASSIFIED → YELLOW) alongside the
    existing EXPLICIT_EXCLUSION (RED) and SCOPE_GAP_IMPLICIT (YELLOW).
    """
    # Build index: (contractor_name, csi_code) → worst AuditStatus.
    from openpyxl.utils import get_column_letter

    _STATUS_RANK = {AuditStatus.RED: 0, AuditStatus.YELLOW: 1, AuditStatus.GREEN: 2}
    worst: dict[tuple[str, str], AuditStatus] = {}
    for item in audit_items:
        if item.division_csi is None:
            continue
        key = (item.contractor_name, item.division_csi)
        if key not in worst or _STATUS_RANK[item.status] < _STATUS_RANK[worst[key]]:
            worst[key] = item.status

    for i, bid in enumerate(bids):
        cost_col = _col_start(i)
        col_letter = get_column_letter(cost_col)
        name = bid.contractor_name

        for csi_code, subtotal_row in subtotal_row_by_csi.items():
            status = worst.get((name, csi_code))
            if status == AuditStatus.RED:
                ws[f"{col_letter}{subtotal_row}"].fill = RED_FILL
            elif status == AuditStatus.YELLOW:
                ws[f"{col_letter}{subtotal_row}"].fill = YELLOW_FILL
            # GREEN / None: no fill change (default)


def _write_audit_sheet(
    wb: openpyxl.Workbook,
    audit_items: list[AuditItem],
) -> None:
    """
    Create and populate the AUDIT worksheet in wb.

    Layout:
      Row 1: Title
      Row 2: Subtitle
      Row 3: blank
      Row 4: Column headers
      Row 5+: One row per AuditItem sorted RED→YELLOW→GREEN, then contractor, then division
      (2 blank rows after last item)
      Summary block: totals by status
    """
    ws = wb.create_sheet(title="AUDIT")
    bold = Font(bold=True)

    # --- Column widths ---
    from openpyxl.utils import get_column_letter
    col_widths = [10, 28, 30, 16, 35, 18, 60]  # A through G
    for col_idx, width in enumerate(col_widths, start=1):
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    # --- Row 1: Title ---
    ws.cell(row=1, column=1).value = (
        "FALKE Matrix — Extraction & Normalization Audit Report"
    )
    ws.cell(row=1, column=1).font = bold

    # --- Row 2: Subtitle ---
    ws.cell(row=2, column=1).value = (
        "Generated by ARA Pipeline | Items requiring action before bid award"
    )

    # --- Row 3: blank ---

    # --- Row 4: Column headers ---
    headers = ["Status", "Code", "Contractor", "Division", "Line Item", "Value", "Message"]
    for col_idx, header in enumerate(headers, start=1):
        c = ws.cell(row=4, column=col_idx)
        c.value = header
        c.font = bold

    # --- Sort items: RED first, then YELLOW, then GREEN; then by contractor; then division ---
    sorted_items = sorted(
        audit_items,
        key=lambda a: (
            _STATUS_SORT_KEY[a.status],
            a.contractor_name,
            a.division_csi or "",
        ),
    )

    # --- Rows 5+: One row per AuditItem ---
    data_start_row = 5
    for row_offset, item in enumerate(sorted_items):
        row = data_start_row + row_offset
        fill = _STATUS_FILL[item.status]

        values = [
            item.status.value,
            item.code.value,
            item.contractor_name,
            item.division_csi or "",
            item.line_item_desc or "",
            item.value or "",
            item.message,
        ]
        for col_idx, val in enumerate(values, start=1):
            c = ws.cell(row=row, column=col_idx)
            c.value = val
            c.fill = fill
            # Bold the Status cell text
            if col_idx == 1:
                c.font = Font(bold=True)

    # --- Summary block (2 blank rows after last data row) ---
    last_data_row = data_start_row + len(sorted_items) - 1
    summary_start = last_data_row + 3  # 2 blank rows then summary

    red_count    = sum(1 for a in audit_items if a.status == AuditStatus.RED)
    yellow_count = sum(1 for a in audit_items if a.status == AuditStatus.YELLOW)
    green_count  = sum(1 for a in audit_items if a.status == AuditStatus.GREEN)
    total_count  = len(audit_items)

    summary_lines = [
        (f"Total items audited:   {total_count}", None),
        (f"RED Critical:          {red_count}  — must resolve before award", RED_FILL),
        (f"YELLOW Review:         {yellow_count}  — verify before finalizing", YELLOW_FILL),
        (f"GREEN Verified:        {green_count}  — clean", GREEN_FILL),
    ]
    for i, (text, fill) in enumerate(summary_lines):
        c = ws.cell(row=summary_start + i, column=1)
        c.value = text
        c.font = bold
        if fill:
            c.fill = fill


# ---------------------------------------------------------------------------
# Primary entry point
# ---------------------------------------------------------------------------

def write_matrix(
    bids: list[NormalizedBid],
    output_path: str | Path,
    run: RunInputs,
    audit_items: Optional[list[AuditItem]] = None,
) -> list[dict]:
    """
    Write normalized bid data into a fresh openpyxl Workbook.

    Project identity (title/address/SF basis) is supplied per-run via ``run``
    (RunInputs) — never hardcoded (M1/M2). ``run.gross_sf`` is the confirmed
    $/SF denominator; ``run.sf_basis_label`` labels the $/SF header.

    Steps:
      1. Sort bids by leveled_total ascending (Floyd C7).
      2. Create a new Workbook and write header rows (1–8) + single-bid notice (N3).
      3. Write CSI division rows (row 9+).
      4. Apply audit-driven color fills to subtotal rows (if audit_items provided).
      5. Write footer section (blank separator + footer rows).
      6. Write alternates section (M7) then qualifications section.
      7. Set column widths.
      8. Write AUDIT sheet (if audit_items provided).
      9. Save to output_path.

    Returns a list of per-bid summary dicts for reporting.
    """
    output_path = Path(output_path)
    gsf = int(run.gross_sf)

    # Step 1: Sort bids (leveled-total ascending)
    ordered_bids = _sort_bids(bids)
    print(f"  [write_matrix] Contractor order: "
          f"{[b.contractor_name for b in ordered_bids]}")

    # Step 2: Create fresh workbook
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Bid_Form"

    # Write header rows (rows 1–8)
    _write_header_rows(ws, ordered_bids, gsf, run)

    # N3: single-bidder notice — a leveled "comparison" of one bid is not a
    # comparison. Render it verbatim near the header; no fabricated competition.
    if len(ordered_bids) == 1:
        notice = ws.cell(row=3, column=1)
        notice.value = "Single bid — no competitive comparison available."
        notice.font = Font(bold=True, italic=True)

    # Step 3: Write division rows starting at row 9
    DIVISION_START_ROW = 9
    next_row, subtotal_row_by_csi = _write_division_rows(ws, ordered_bids, DIVISION_START_ROW, gsf)

    # Step 4: Apply audit color fills to subtotal rows
    if audit_items:
        _apply_audit_fills(ws, ordered_bids, subtotal_row_by_csi, audit_items)

    # Step 5: Write footer rows (includes blank separator before)
    next_row, footer_summaries = _write_footer_rows(ws, ordered_bids, next_row, gsf)

    # Step 6: Write alternates (own labeled section) then qualifications
    next_row = _write_alternates(ws, ordered_bids, next_row)
    _write_qualifications(ws, ordered_bids, next_row)

    # Step 7: Column widths
    _set_column_widths(ws, len(ordered_bids))

    # Step 8: Write AUDIT sheet
    if audit_items:
        _write_audit_sheet(wb, audit_items)

    # Step 9: Save
    wb.save(output_path)
    print(f"  [write_matrix] Saved fresh workbook → {output_path}")
    print(f"  [write_matrix] Sheet dimensions: "
          f"{ws.max_row} rows × {ws.max_column} cols")

    # Build per-bid summary dicts (same shape as original for pipeline.py).
    # Row numbers are no longer sequential (dynamic per-item layout), so we
    # report 0 as the row sentinel and let pipeline.py print amounts only.
    summaries: list[dict] = []
    for i, bid in enumerate(ordered_bids):
        # Aggregate subtotal per CSI code (handles duplicate codes)
        seen: dict[str, float] = {}
        for div in bid.divisions:
            amount = _cell_amount(div.subtotal_cell.state, div.subtotal_cell.amount)
            seen[div.csi_code] = seen.get(div.csi_code, 0.0) + amount

        divisions_written = [
            {
                "csi_code": csi_code,
                "row": 0,  # dynamic layout — row number not fixed
                "amount": seen.get(csi_code, 0.0),
                "state": "AMOUNT",
            }
            for csi_code, _ in DIVISION_ROWS
        ]

        summaries.append({
            "contractor": bid.contractor_name,
            "matched": True,
            "name_col": _col_start(i),
            "divisions_written": divisions_written,
            "footer_written": footer_summaries[i],
            "warnings": [],
        })

    return summaries
