"""
FALKE Matrix Pipeline — Normalization Layer Output Models
=========================================================
These Pydantic models represent the output contract of the normalization
rule engine.  They are designed to be consumed directly by the Excel writer
without further interpretation.

Pipeline position:
    BidDocument  →  [Normalization Rule Engine (normalize.py)]
    →  NormalizedBid (this schema)  →  [Excel Writer]

Every cell in the matrix is represented as a CellValue, which carries both
a CellState (semantic classification) and a display string (board-ready).
The Excel writer must use display, not raw amounts, for cell text rendering.
"""

from __future__ import annotations

from decimal import Decimal
from enum import Enum
from typing import Optional

from pydantic import BaseModel, Field

# Re-export enums from models that the Excel writer will need
from src.models import (
    CostStructure,
    ExtractionConfidence,
    FormType,
    GrandTotalConfidence,
    InputType,
)


# ---------------------------------------------------------------------------
# Cell-level semantics
# ---------------------------------------------------------------------------

class CellState(str, Enum):
    """
    Semantic state of a single matrix cell.  The Excel writer must render
    each state distinctly — never treat EXCLUDED as $0 or NULL_BLANK as zero.
    """
    AMOUNT = "AMOUNT"
    """Contractor priced this item — display the dollar amount."""

    EXPLICIT_ZERO = "EXPLICIT_ZERO"
    """Contractor explicitly entered $0 — item is in scope at no cost."""

    NULL_BLANK = "NULL_BLANK"
    """
    Cell was blank in the source document.  Potential scope gap.
    Displayed as '-'.  When field-median > $20K, SCOPE_GAP_IMPLICIT is
    added to flags.
    """

    EXCLUDED = "EXCLUDED"
    """
    Contractor explicitly excluded this item from scope.
    Display as 'EXCL'.  Requires a plug number before the total is used
    for cross-bid comparison.
    """

    BY_OWNER_OTHERS = "BY_OWNER_OTHERS"
    """
    Item is marked 'By Others' / 'By Owner'.  Display as 'BY OTHERS'.
    Must be excluded from the contractor's leveled construction total.
    """

    ALLOWANCE = "ALLOWANCE"
    """
    Item is an allowance (estimate, not firm price).  Display as 'ALLOW $X'.
    Included in division subtotal but NOT in the hard-cost leveled total.
    Flagged separately in the board summary.
    """


class CellValue(BaseModel):
    """A single leveled matrix cell — the atom of the normalized output."""

    state: CellState

    amount: Optional[Decimal] = None
    """
    Set when state is AMOUNT, EXPLICIT_ZERO, or ALLOWANCE.
    None for NULL_BLANK, EXCLUDED, BY_OWNER_OTHERS.
    """

    display: str
    """
    Board-display string, ready for Excel rendering:
      AMOUNT          → '$120,000'
      EXPLICIT_ZERO   → '$0'
      NULL_BLANK      → '-'
      EXCLUDED        → 'EXCL'
      BY_OWNER_OTHERS → 'BY OTHERS'
      ALLOWANCE       → 'ALLOW $50,000'
    """

    is_reclassified: bool = False
    """True when this cell's division was PIPELINE_REMAPPED from a wrong division."""

    reclassified_from: Optional[str] = None
    """Original CSI division code if is_reclassified=True."""

    flags: list[str] = Field(default_factory=list)
    """
    Machine-readable flags on this cell.  Examples:
      'SCOPE_GAP_IMPLICIT'      — NULL_BLANK in a division where field-median > $20K
      'ARITHMETIC_DISCREPANCY'  — line items do not sum to stated subtotal
    """


# ---------------------------------------------------------------------------
# Division-level aggregation
# ---------------------------------------------------------------------------

class NormalizedDivision(BaseModel):
    """One CSI division's normalized bid data for a single contractor."""

    csi_code: str
    """Canonical Falke division code: DIV XX 00 00."""

    division_name: str

    line_item_cells: dict[str, CellValue] = Field(default_factory=dict)
    """
    Keyed by canonical sub-line label (from canon.CANONICAL_DIVISIONS).
    May also include contractor-native labels when no canonical mapping exists.
    """

    subtotal_cell: CellValue
    """Division-level rolled-up cell.  Displayed in the subtotal row of the matrix."""

    cost_structure: CostStructure
    """Pricing structure from the source DivisionBid."""


# ---------------------------------------------------------------------------
# Footer / bid-level aggregation
# ---------------------------------------------------------------------------

class NormalizedAlternate(BaseModel):
    """One add/deduct bid alternate, kept separate from the base comparison (M7)."""

    description: str
    amount: Optional[Decimal] = None
    display: str


class NormalizedFooter(BaseModel):
    """
    The fee, insurance, and total section of the normalized bid.
    Parallels BidFooter but with fully resolved CellValues and computed fields.
    """

    construction_subtotal: CellValue
    """Sum of all division subtotals (including BY_OWNER_OTHERS, excluding nothing)."""

    general_liability_insurance: CellValue
    builders_risk_insurance: CellValue
    gc_fee: CellValue
    grand_total: CellValue
    bond: CellValue

    gc_fee_pct: Optional[Decimal] = None
    """
    Computed: gc_fee / construction_subtotal * 100.
    None when either value is missing or zero (Rule 5 Phase 1).
    """

    grand_total_confidence: GrandTotalConfidence
    confidence_flags: list[str] = Field(default_factory=list)

    leveled_total: Optional[Decimal] = None
    """
    Grand total minus the sum of BY_OWNER_OTHERS line-item amounts.
    Allowances are retained in this total — they are contractual.
    BY_OWNER_OTHERS items are excluded — they are not the contractor's direct cost.
    None when grand_total is not set.
    """

    alternates: list[NormalizedAlternate] = Field(default_factory=list)
    """
    Bid alternates (add/deduct options), surfaced for their OWN section in the
    matrix — never folded into the base/leveled total (M7, instrument-separation).
    """


# ---------------------------------------------------------------------------
# Summary flags
# ---------------------------------------------------------------------------

class BidSummaryFlag(BaseModel):
    """
    A board-memo-ready flag on a normalized bid.  Consumed by the board
    summary generator AND by audit_bids(), which promotes flags whose
    flag_type matches an AuditCode into first-class AuditItems on the AUDIT
    sheet (the C4 contract — remap/reclass/ambiguous flags must reach the
    board-facing sheet and feed the cell-coloring pass).
    """

    flag_type: str
    """
    Machine-readable type key.  Examples:
      'SCOPE_GAP_IMPLICIT', 'GC_FEE_OUTLIER', 'ALLOWANCE_PRESENT',
      'CODE_FORMAT_REMAPPED', 'KNOWN_FIRM_RECLASSIFIED',
      'UNRECOGNIZED_CODE_FORMAT', 'KNOWN_FIRM_AMBIGUOUS', 'CODE_SPLIT_UNMATCHED'
    """

    message: str
    """Human-readable, board-memo-ready description of the flag."""

    severity: str
    """'info' | 'warning' | 'critical'"""

    division_csi: Optional[str] = None
    """
    Canonical CSI code this flag attaches to (DIV XX 00 00), or None for a
    bid-level flag.  audit_bids() uses this to place the promoted AuditItem on
    the right division row and to drive _apply_audit_fills cell coloring (C4).
    """

    line_item_desc: Optional[str] = None
    """The specific line-item description involved, when the flag is line-scoped."""

    value: Optional[str] = None
    """The concrete value (e.g. native code, '15') surfaced on the AUDIT row."""


# ---------------------------------------------------------------------------
# Top-level normalized bid
# ---------------------------------------------------------------------------

class NormalizedBid(BaseModel):
    """
    Top-level normalization artifact: one contractor's complete bid after
    all rule-engine transformations.  The Excel writer consumes this directly.
    """

    contractor_name: str
    project_name: Optional[str] = None

    form_type: FormType
    bid_document_input_type: InputType
    extraction_confidence: ExtractionConfidence

    divisions: list[NormalizedDivision] = Field(default_factory=list)
    footer: NormalizedFooter

    qualifications_text: str
    """
    Concatenation of notes + qualifications + exclusions + assumptions + terms
    from BidQualifications.  Plain text, newline-separated sections.
    """

    # --- Allowance accounting ---
    total_allowance_value: Decimal = Decimal("0")
    """Sum of all line items where is_allowance=True."""

    allowance_count: int = 0
    """Number of allowance line items across all divisions."""

    # --- Scope gap accounting ---
    explicit_exclusion_count: int = 0
    """Number of line items where is_excluded=True."""

    implicit_gap_count: int = 0
    """
    Count of NULL_BLANK cells in divisions where field-median > $20K.
    Set after cross-bid median computation; 0 before that step.
    """

    # --- Pass-through and generated warnings ---
    extraction_warnings: list[str] = Field(default_factory=list)
    """Warnings generated by the extraction layer — passed through unchanged."""

    normalization_warnings: list[str] = Field(default_factory=list)
    """Warnings generated by the normalization rule engine."""

    summary_flags: list[BidSummaryFlag] = Field(default_factory=list)
    """Structured flags consumed by the board-memo generator."""
