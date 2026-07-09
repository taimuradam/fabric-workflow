"""
excel_export.py — Build a formatted .xlsx that mirrors the paper costing sheet.

Consumes the SAME Results object the GUI displays (from calculations.compute), so
the spreadsheet always agrees with what the user saw on screen.
"""

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from calculations import LotInfo, Results

# PKR currency format: "Rs" prefix, thousands separator, 2 decimals.
_CURRENCY_FMT = '"Rs" #,##0.00'
_NUMBER_FMT = "#,##0.00"

_HEADER_FILL = PatternFill("solid", fgColor="DDE6F0")   # light blue
_WASTAGE_FILL = PatternFill("solid", fgColor="FCE8D5")  # light orange
_TOTAL_FILL = PatternFill("solid", fgColor="EDEDED")    # light grey

_BOLD = Font(bold=True)
_TITLE_FONT = Font(bold=True, size=14)
_WARN_FONT = Font(bold=True, color="C55A11")            # orange
_THIN = Side(style="thin", color="BBBBBB")
_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)


def _dash(value):
    """Render None as an em dash, otherwise pass the value through."""
    return "—" if value is None else value


def export(results: Results, lot: LotInfo, path: str) -> None:
    """Write the full costing sheet to `path`."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Costing"

    # Column widths for a readable products table (A..F).
    for col, width in zip("ABCDEF", (26, 14, 12, 14, 16, 18)):
        ws.column_dimensions[col].width = width

    row = 1

    # ------------------------------------------------------------------ Header
    ws.cell(row=row, column=1, value="Textile Lot Costing").font = _TITLE_FONT
    row += 2

    def label_value(label, value, money=False, number=False):
        """Emit a 'Label | value' pair on the current row and advance."""
        nonlocal row
        lc = ws.cell(row=row, column=1, value=label)
        lc.font = _BOLD
        vc = ws.cell(row=row, column=2, value=_dash(value))
        if money and value is not None:
            vc.number_format = _CURRENCY_FMT
        elif number and value is not None:
            vc.number_format = _NUMBER_FMT
        row += 1

    label_value("Lot Reference", lot.reference or "—")
    label_value("Fabric Type", lot.fabric_type or "—")
    label_value("Date", lot.date or "—")
    row += 1

    # ------------------------------------------------------- Lot info block
    ws.cell(row=row, column=1, value="Lot Details").font = _BOLD
    row += 1
    label_value("GSM", lot.gsm, number=True)
    label_value("Width (inches)", lot.width_in, number=True)
    label_value("Total KG Received", lot.total_kg, number=True)
    label_value("Rate per Meter", lot.rate_per_meter, money=True)
    label_value("Transport Cost", lot.transport_cost, money=True)
    label_value("Meters per KG", results.meters_per_kg, number=True)
    label_value("Total Meters", results.total_meters, number=True)
    label_value("Fabric Cost", results.fabric_cost, money=True)
    label_value("Total Cost", results.total_cost, money=True)
    label_value("Base Cost per KG", results.base_cost_per_kg, money=True)
    label_value("Wastage Cost per KG", results.wastage_cost_per_kg, money=True)
    label_value("Adjusted Cost per KG", results.adjusted_cost_per_kg, money=True)
    row += 1

    # ------------------------------------------------------- Products table
    ws.cell(row=row, column=1, value="Products").font = _BOLD
    row += 1

    headers = ["Product", "Weight (kg)", "Pieces", "Wt./Pc", "Cost/Pc", "Revenue"]
    for col, text in enumerate(headers, start=1):
        c = ws.cell(row=row, column=col, value=text)
        c.font = _BOLD
        c.fill = _HEADER_FILL
        c.border = _BORDER
        c.alignment = Alignment(horizontal="center")
    row += 1

    money_cols = {5, 6}   # Cost/Pc, Revenue
    number_cols = {2, 3, 4}  # Weight, Pieces, Wt/Pc

    for pr in results.product_results:
        p = pr.product
        name = p.name or "—"
        if p.is_wastage:
            # Make the wastage row unmistakable.
            name = f"{name}  (WASTAGE)" if p.name else "Wastage"

        values = [
            name,
            _dash(_num(p.weight_kg)),
            "—" if p.is_wastage else _dash(_num(p.pieces)),
            "N/A" if (not p.is_wastage and pr.weight_per_piece is None) else
            ("—" if p.is_wastage else _dash(pr.weight_per_piece)),
            "—" if p.is_wastage else _dash(pr.cost_per_piece),
            "—" if p.is_wastage else _dash(pr.revenue),
        ]
        for col, value in enumerate(values, start=1):
            c = ws.cell(row=row, column=col, value=value)
            c.border = _BORDER
            if p.is_wastage:
                c.fill = _WASTAGE_FILL
            if isinstance(value, (int, float)):
                if col in money_cols:
                    c.number_format = _CURRENCY_FMT
                elif col in number_cols:
                    c.number_format = _NUMBER_FMT
        row += 1

    # Bold totals row for the products table.
    total_row = [
        "TOTAL",
        results.total_weight_produced,
        results.total_pieces,
        "",
        "",
        results.total_receipt,
    ]
    for col, value in enumerate(total_row, start=1):
        c = ws.cell(row=row, column=col, value=value)
        c.font = _BOLD
        c.fill = _TOTAL_FILL
        c.border = _BORDER
        if isinstance(value, (int, float)):
            if col == 6:
                c.number_format = _CURRENCY_FMT
            elif col in number_cols:
                c.number_format = _NUMBER_FMT
    row += 2

    # ------------------------------------------------------- Summary block
    ws.cell(row=row, column=1, value="Summary").font = _BOLD
    row += 1
    label_value("Total Payment (Cost)", results.total_payment, money=True)
    label_value("Total Receipt (Revenue)", results.total_receipt, money=True)
    profit_label = ws.cell(row=row, column=1, value="Profit")
    profit_label.font = _BOLD
    profit_cell = ws.cell(row=row, column=2, value=_dash(results.profit))
    if results.profit is not None:
        profit_cell.number_format = _CURRENCY_FMT
        # Green profit / red loss, matching the on-screen colouring.
        color = "2E7D32" if results.profit >= 0 else "C62828"
        profit_cell.font = Font(bold=True, color=color)
    row += 2

    # ------------------------------------------------------- Reconciliation note
    if results.recon_mismatch:
        c = ws.cell(row=row, column=1, value=results.recon_message)
        c.font = _WARN_FONT
        row += 1

    wb.save(path)


def _num(value):
    """Coerce a stored input into a float for the spreadsheet, or None."""
    from calculations import parse_number
    return parse_number(value)
