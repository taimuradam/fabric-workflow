"""
calculations.py — Textile lot costing business rules (single source of truth).

This module is intentionally PURE: it imports no tkinter and no openpyxl, so the
same formulas drive both the on-screen display and the Excel export, and can be
unit-tested in isolation.

The vocabulary comes from the paper costing sheet this app replaces:
  * A "lot" is a batch of raw fabric bought by weight (kg) from a supplier.
  * The lot is cut/stitched into finished "products" (bedsheets, pillows, ...).
  * Some fabric is unusable "wastage" — it produces no sellable pieces but its
    cost still has to be recovered from the products that DID sell.

All numeric inputs are Optional[float]: None means "blank or not-yet-valid". The
UI holds partial/invalid input all the time (user mid-typing), so every function
degrades gracefully — a missing input just makes the dependent outputs None
instead of raising. Nothing here should ever crash the GUI.
"""

from dataclasses import dataclass, field
from typing import List, Optional


# Meters below which we treat a fabric dimension as effectively zero, to avoid
# dividing by zero when the user hasn't filled a field in yet.
_EPSILON = 1e-9

# Inches -> meters. Width is entered in inches; the meters-per-kg formula needs
# the width expressed in meters.
_INCH_TO_METER = 0.0254

# Reconciliation tolerance. Weights are entered to ~2 decimals, so anything under
# this is rounding noise, not a real data-entry mismatch.
_RECON_TOLERANCE = 0.01


def parse_number(value) -> Optional[float]:
    """Parse a user-entered value into a float, tolerantly.

    Returns None for empty strings, whitespace, or anything non-numeric — this is
    the friendly-validation primitive used everywhere so the app never crashes on
    half-typed or garbage input. Accepts numbers with thousands separators
    ("1,234.5") and surrounding whitespace.
    """
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", "")
    if text == "":
        return None
    try:
        return float(text)
    except ValueError:
        return None


@dataclass
class LotInfo:
    """Top-section inputs describing the raw fabric lot."""

    reference: str = ""            # e.g. "CVC-3"
    fabric_type: str = ""          # e.g. "CVC"
    date: str = ""                 # ISO date string, e.g. "2026-07-09"
    gsm: Optional[float] = None    # grams per square meter
    width_in: Optional[float] = None       # fabric width in inches
    total_kg: Optional[float] = None       # total weight received from supplier
    rate_per_meter: Optional[float] = None # price paid to supplier per meter
    transport_cost: Optional[float] = None # flat transport fee for the lot


@dataclass
class Product:
    """One row in the products table."""

    name: str = ""
    weight_kg: Optional[float] = None  # kg of fabric this product used
    pieces: Optional[float] = None     # number of finished pieces (0/blank = none)
    is_wastage: bool = False           # exactly one row should carry this flag


@dataclass
class ProductResult:
    """Computed, display-ready values for a single product row."""

    product: Product
    weight_per_piece: Optional[float] = None  # None => shown as "N/A"
    cost_per_piece: Optional[float] = None
    revenue: Optional[float] = None


@dataclass
class Results:
    """Everything derived from a (LotInfo, [Product]) pair.

    Both the GUI and the Excel exporter read from one of these so the numbers can
    never drift apart between screen and spreadsheet.
    """

    # --- Lot-level derived values (all Optional; None => not computable yet) ---
    meters_per_kg: Optional[float] = None
    total_meters: Optional[float] = None
    fabric_cost: Optional[float] = None
    total_cost: Optional[float] = None
    base_cost_per_kg: Optional[float] = None
    wastage_weight: float = 0.0
    wastage_cost_per_kg: Optional[float] = None
    adjusted_cost_per_kg: Optional[float] = None

    # --- Per-row results (same order as the input products list) ---
    product_results: List[ProductResult] = field(default_factory=list)

    # --- Summary / totals ---
    total_weight_produced: float = 0.0
    total_pieces: float = 0.0
    total_payment: Optional[float] = None   # = total_cost (what we paid)
    total_receipt: float = 0.0              # sum of product revenues (what we recover)
    profit: Optional[float] = None          # receipt - payment

    # --- Reconciliation (data-entry sanity check) ---
    recon_mismatch: bool = False
    recon_message: str = ""
    recon_sum_weights: float = 0.0


def kg_from_meters(gsm, width_in, total_meters) -> Optional[float]:
    """Derive the lot's weight (kg) from its length in meters.

    The supplier's invoice states fabric in meters, so the UI asks for meters
    and derives the weight. This is the exact inverse of the meters-per-kg
    formula in compute(): a strip 1 m long and `width` wide weighs
    (width_m × GSM) grams, so `meters` of it weigh
    meters × width_m × GSM / 1000 kg.

    Returns None (tolerantly) when any input is missing, non-numeric, or
    non-positive — same contract as the rest of this module.
    """
    gsm_v = parse_number(gsm)
    width_v = parse_number(width_in)
    meters_v = parse_number(total_meters)
    if not gsm_v or not width_v or meters_v is None:
        return None
    if gsm_v <= _EPSILON or width_v <= _EPSILON:
        return None
    return meters_v * (width_v * _INCH_TO_METER) * gsm_v / 1000.0


def _wastage_weight(products: List[Product]) -> float:
    """Weight of the (first) row flagged as wastage, or 0.0 if none is flagged.

    Only one row is meant to be wastage; if several are flagged we defensively use
    the first so a mis-click can't silently double-count.
    """
    for p in products:
        if p.is_wastage:
            return parse_number(p.weight_kg) or 0.0
    return 0.0


def compute(lot: LotInfo, products: List[Product]) -> Results:
    """Compute all derived values for a lot. Never raises on bad input."""
    r = Results()

    gsm = parse_number(lot.gsm)
    width_in = parse_number(lot.width_in)
    total_kg = parse_number(lot.total_kg)
    rate = parse_number(lot.rate_per_meter)
    transport = parse_number(lot.transport_cost)

    # --- Meters per KG -------------------------------------------------------
    # A square meter of this fabric weighs GSM grams. A strip 1 m long and
    # `width` wide weighs (width_m * GSM) grams; so 1000 g (=1 kg) buys
    # 1000 / (width_m * GSM) meters of fabric.
    if gsm and width_in and gsm > _EPSILON and width_in > _EPSILON:
        width_m = width_in * _INCH_TO_METER
        r.meters_per_kg = 1000.0 / (width_m * gsm)

    # --- Total meters in the lot --------------------------------------------
    if r.meters_per_kg is not None and total_kg is not None:
        r.total_meters = r.meters_per_kg * total_kg

    # --- Fabric cost = meters * supplier rate --------------------------------
    if r.total_meters is not None and rate is not None:
        r.fabric_cost = r.total_meters * rate

    # --- Total cost = fabric + flat transport --------------------------------
    if r.fabric_cost is not None:
        # Transport is optional; treat a blank transport fee as 0.
        r.total_cost = r.fabric_cost + (transport or 0.0)

    # --- Base cost per KG ----------------------------------------------------
    if r.total_cost is not None and total_kg and total_kg > _EPSILON:
        r.base_cost_per_kg = r.total_cost / total_kg

    # --- Wastage cost spreading ---------------------------------------------
    # Wastage fabric produces nothing, but we still paid for it. This is a
    # cost-of-creation tool, so every rupee we spent must land on a finished
    # piece — nothing may leak off the books. The waste cost is therefore spread
    # over the SELLABLE weight only (total_kg - wastage_weight), because that is
    # the only fabric that can carry it. Doing so makes the allocated per-piece
    # costs add back up to the total cost exactly (full recovery).
    #
    #   wastage_cost         = wastage_weight * base_cost_per_kg
    #   sellable_weight      = total_kg - wastage_weight
    #   wastage_cost_per_kg  = wastage_cost / sellable_weight   # extra per good kg
    #   adjusted_cost_per_kg = base_cost_per_kg + wastage_cost_per_kg
    r.wastage_weight = _wastage_weight(products)
    if r.base_cost_per_kg is not None and total_kg and total_kg > _EPSILON:
        sellable_weight = total_kg - r.wastage_weight
        if sellable_weight > _EPSILON:
            wastage_cost = r.wastage_weight * r.base_cost_per_kg
            r.wastage_cost_per_kg = wastage_cost / sellable_weight
            r.adjusted_cost_per_kg = r.base_cost_per_kg + r.wastage_cost_per_kg
        else:
            # No sellable fabric to carry the cost (all waste / empty). Fall back
            # to the base rate rather than dividing by zero.
            r.wastage_cost_per_kg = 0.0
            r.adjusted_cost_per_kg = r.base_cost_per_kg

    # --- Per-product results -------------------------------------------------
    total_weight = 0.0
    total_pieces = 0.0
    total_receipt = 0.0

    for p in products:
        pr = ProductResult(product=p)
        weight = parse_number(p.weight_kg)
        pieces = parse_number(p.pieces)

        if weight is not None:
            total_weight += weight

        # The wastage row is pure cost — no pieces, no per-piece math, no revenue.
        if not p.is_wastage:
            if pieces is not None:
                total_pieces += pieces

            # Weight per piece: skip (show "N/A") when pieces is 0/blank so we
            # never divide by zero.
            if weight is not None and pieces and pieces > _EPSILON:
                pr.weight_per_piece = weight / pieces
                if r.adjusted_cost_per_kg is not None:
                    pr.cost_per_piece = pr.weight_per_piece * r.adjusted_cost_per_kg
                    pr.revenue = pieces * pr.cost_per_piece
                    total_receipt += pr.revenue

        r.product_results.append(pr)

    r.total_weight_produced = total_weight
    r.total_pieces = total_pieces
    r.total_receipt = total_receipt
    r.total_payment = r.total_cost
    if r.total_payment is not None:
        r.profit = r.total_receipt - r.total_payment

    # --- Reconciliation check ------------------------------------------------
    # Fabric received should equal fabric turned into products + wastage. If the
    # entered Total KG doesn't match the summed product weights, the operator
    # likely mistyped something. We WARN but never block — the numbers may be
    # intentionally provisional.
    r.recon_sum_weights = total_weight
    if total_kg is not None and abs(total_kg - total_weight) > _RECON_TOLERANCE:
        r.recon_mismatch = True
        r.recon_message = (
            f"⚠ Total KG Received ({total_kg:,.2f}) doesn't match sum of "
            f"product weights ({total_weight:,.2f}) — check your entries"
        )

    return r
