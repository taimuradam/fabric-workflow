"""
main.py — Textile Lot Costing desktop app (CustomTkinter GUI + entry point).

Single-window layout, top to bottom:
    Toolbar (New / Save / Open / Export)
    Lot Info card + Lot Calculations stat grid
    Products table (manual CTk widgets) + Add Row
    Reconciliation warning (only when weights don't add up)
    Summary card (large, colour-coded Profit)

All heavy lifting lives in calculations.py; this file is purely presentation and
wiring. Every input change funnels into a single recompute() so the whole screen
stays consistent, and all number parsing is tolerant so bad input never crashes.

CustomTkinter has no Treeview-equivalent, so the products table is built by hand:
a CTkScrollableFrame whose rows are strips of always-editable CTkEntry widgets
(plus a wastage CTkCheckBox and a delete button). This keeps a consistent flat,
rounded look instead of embedding a native-themed ttk.Treeview.
"""

from datetime import date
from tkinter import filedialog, messagebox

import customtkinter as ctk

import calculations
import excel_export
import storage
from calculations import LotInfo, Product

# ----------------------------------------------------------------- palette ---
# One restrained accent (deep teal) plus neutrals. Colours are given as plain
# hex; appearance mode is fixed to light so we don't need light/dark pairs.
TEAL        = "#0f766e"
TEAL_HOVER  = "#0b5f58"
TEAL_TINT   = "#dcebe8"   # highlighted "Adjusted Cost/KG" tile + wastage rows
BG          = "#eef0ec"   # window background (off-white)
CARD        = "#ffffff"
CARD_BORDER = "#dde2dd"
TILE_BG     = "#f3f5f2"   # read-only / computed field tint
TEXT        = "#1f2a30"
MUTED       = "#67757c"
SECONDARY   = "#e4e8e3"   # secondary button fill
SECONDARY_H = "#d6dbd5"   # secondary button hover
GREEN       = "#15803d"
RED         = "#b91c1c"
ORANGE      = "#c05621"

# Product table column widths (pixels). Header labels and row widgets share these
# so the columns line up even though they live in separate frames.
COL = {
    "name": 210, "weight": 95, "pieces": 95, "wastage": 78,
    "wtpc": 90, "costpc": 120, "revenue": 132, "del": 40,
}


def fmt(value, money=False, decimals=2):
    """Format a computed value for display; None -> em dash."""
    if value is None:
        return "—"
    if money:
        return f"Rs {value:,.{decimals}f}"
    return f"{value:,.{decimals}f}"


class ProductRow:
    """One editable product row in the manual table (widgets + state)."""

    def __init__(self, table, app, product):
        self.app = app
        self.frame = ctk.CTkFrame(table, fg_color="transparent")
        self.wastage_var = ctk.IntVar(value=1 if product.is_wastage else 0)

        # --- editable cells ---
        self.name = ctk.CTkEntry(self.frame, width=COL["name"], font=app.font_body,
                                 fg_color=CARD, border_color=CARD_BORDER)
        self.name.insert(0, product.name or "")

        self.weight = ctk.CTkEntry(self.frame, width=COL["weight"], justify="right",
                                   font=app.font_body, fg_color=CARD, border_color=CARD_BORDER)
        if product.weight_kg not in (None, ""):
            self.weight.insert(0, str(product.weight_kg))

        self.pieces = ctk.CTkEntry(self.frame, width=COL["pieces"], justify="right",
                                   font=app.font_body, fg_color=CARD, border_color=CARD_BORDER)
        if product.pieces not in (None, ""):
            self.pieces.insert(0, str(product.pieces))

        # --- wastage checkbox (centred in a fixed-width holder) ---
        waste_holder = ctk.CTkFrame(self.frame, width=COL["wastage"], height=32,
                                    fg_color="transparent")
        waste_holder.pack_propagate(False)
        self.check = ctk.CTkCheckBox(waste_holder, text="", width=24,
                                     variable=self.wastage_var, onvalue=1, offvalue=0,
                                     fg_color=TEAL, hover_color=TEAL_HOVER,
                                     command=lambda: app._on_wastage(self))
        self.check.pack(expand=True)

        # --- read-only computed cells (tinted labels — clearly not editable) ---
        self.wtpc = self._calc_label(app, COL["wtpc"])
        self.costpc = self._calc_label(app, COL["costpc"])
        self.revenue = self._calc_label(app, COL["revenue"])

        # --- delete button ---
        self.delete = ctk.CTkButton(
            self.frame, text="✕", width=COL["del"], font=app.font_body,
            fg_color="transparent", text_color=MUTED, hover_color="#f6dede",
            command=lambda: app._delete_row(self),
        )

        # Lay the strip out left-to-right; fixed widths keep header alignment.
        pad = 4
        for w in (self.name, self.weight, self.pieces, waste_holder,
                  self.wtpc, self.costpc, self.revenue, self.delete):
            w.pack(side="left", padx=pad, pady=2)

        # Live recompute on typing.
        for e in (self.name, self.weight, self.pieces):
            e.bind("<KeyRelease>", app.recompute)

        self._apply_wastage_style(product.is_wastage)

    def _calc_label(self, app, width):
        return ctk.CTkLabel(self.frame, text="—", width=width, height=30,
                            font=app.font_body, text_color=TEXT,
                            fg_color=TILE_BG, corner_radius=6, anchor="e",
                            padx=8)

    def is_wastage(self):
        return self.wastage_var.get() == 1

    def to_product(self):
        w = self.weight.get().strip()
        p = self.pieces.get().strip()
        return Product(
            name=self.name.get().strip(),
            weight_kg=w if w != "" else None,
            pieces=p if p != "" else None,
            is_wastage=self.is_wastage(),
        )

    def _apply_wastage_style(self, is_wastage):
        """Wastage rows get a teal tint and no pieces (pure cost)."""
        self.frame.configure(fg_color=TEAL_TINT if is_wastage else "transparent")
        if is_wastage:
            self.pieces.delete(0, "end")
            self.pieces.configure(state="disabled")
        else:
            self.pieces.configure(state="normal")

    def paint(self, result):
        """Write the computed Wt./Pc, Cost/Pc and Revenue cells."""
        if self.is_wastage():
            wtpc = costpc = rev = "—"
        elif result is not None:
            wtpc = "N/A" if result.weight_per_piece is None else fmt(result.weight_per_piece)
            costpc = fmt(result.cost_per_piece, money=True)
            rev = fmt(result.revenue, money=True)
        else:
            wtpc = costpc = rev = "—"
        self.wtpc.configure(text=wtpc)
        self.costpc.configure(text=costpc)
        self.revenue.configure(text=rev)

    def destroy(self):
        self.frame.destroy()


class CostingApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Textile Lot Costing")
        self.geometry("1180x940")
        self.minsize(1040, 720)

        ctk.set_appearance_mode("light")
        # Base theme is fine; we override accent colours per-widget so only the
        # primary button + checkboxes carry teal (secondary buttons stay neutral).
        self.configure(fg_color=BG)

        # Fonts (created after root exists, as CTkFont requires).
        self.font_title = ctk.CTkFont(size=22, weight="bold")
        self.font_h2    = ctk.CTkFont(size=16, weight="bold")
        self.font_body  = ctk.CTkFont(size=14)
        self.font_label = ctk.CTkFont(size=13)
        self.font_tile  = ctk.CTkFont(size=19, weight="bold")
        self.font_sum   = ctk.CTkFont(size=17)
        self.font_sumb  = ctk.CTkFont(size=17, weight="bold")
        self.font_profit = ctk.CTkFont(size=26, weight="bold")

        self.rows = []          # list[ProductRow] in display order
        self.vars = {}          # lot-info StringVars

        # Scrollable page so the whole layout stays reachable on small screens.
        self.page = ctk.CTkScrollableFrame(self, fg_color="transparent")
        self.page.pack(fill="both", expand=True, padx=18, pady=14)

        self._build_toolbar()
        self._build_lot_card()
        self._build_calc_card()
        self._build_products_card()
        self._build_warning()
        self._build_summary_card()

        # Wire lot-info live updates now that all widgets exist.
        for var in self.vars.values():
            var.trace_add("write", self.recompute)

        self._new_lot(confirm=False)  # start with a clean, dated sheet

    # ------------------------------------------------------------ card helper
    def _card(self, title=None, hint=None):
        card = ctk.CTkFrame(self.page, fg_color=CARD, border_width=1,
                            border_color=CARD_BORDER, corner_radius=12)
        card.pack(fill="x", pady=(0, 16))
        if title:
            ctk.CTkLabel(card, text=title, font=self.font_h2,
                         text_color=TEXT).pack(anchor="w", padx=20, pady=(16, 0))
        if hint:
            ctk.CTkLabel(card, text=hint, font=self.font_label,
                         text_color=MUTED).pack(anchor="w", padx=20, pady=(2, 0))
        return card

    def _button(self, parent, text, command, primary=False):
        if primary:
            return ctk.CTkButton(parent, text=text, command=command,
                                 font=self.font_body, fg_color=TEAL,
                                 hover_color=TEAL_HOVER, text_color="#ffffff",
                                 height=38, corner_radius=8)
        return ctk.CTkButton(parent, text=text, command=command,
                             font=self.font_body, fg_color=SECONDARY,
                             hover_color=SECONDARY_H, text_color=TEXT,
                             height=38, corner_radius=8)

    # ------------------------------------------------------------------ toolbar
    def _build_toolbar(self):
        bar = ctk.CTkFrame(self.page, fg_color="transparent")
        bar.pack(fill="x", pady=(0, 16))
        ctk.CTkLabel(bar, text="Textile Lot Costing", font=self.font_title,
                     text_color=TEXT).pack(side="left")
        # Primary action on the far right, secondaries to its left.
        self._button(bar, "Export to Excel", self._export_excel,
                     primary=True).pack(side="right", padx=(8, 0))
        for text, cmd in [("Open Lot", self._open_lot),
                          ("Save Lot", self._save_lot),
                          ("New Lot", self._new_lot)]:
            self._button(bar, text, cmd).pack(side="right", padx=(8, 0))

    # --------------------------------------------------------------- lot card
    def _build_lot_card(self):
        card = self._card("Lot Information",
                          "Details of the fabric received from the supplier.")
        grid = ctk.CTkFrame(card, fg_color="transparent")
        grid.pack(fill="x", padx=20, pady=(12, 18))

        fields = [
            ("reference", "Lot Reference"), ("fabric_type", "Fabric Type"),
            ("date", "Date"), ("gsm", "GSM"),
            ("width_in", "Width (inches)"), ("total_kg", "Total KG Received"),
            ("rate_per_meter", "Rate per Meter"), ("transport_cost", "Transport Cost"),
        ]
        cols = 4
        for i in range(cols):
            grid.grid_columnconfigure(i, weight=1, uniform="lot")
        for i, (key, label) in enumerate(fields):
            r, c = divmod(i, cols)
            cell = ctk.CTkFrame(grid, fg_color="transparent")
            cell.grid(row=r, column=c, sticky="ew", padx=8, pady=8)
            ctk.CTkLabel(cell, text=label, font=self.font_label,
                         text_color=MUTED).pack(anchor="w")
            var = ctk.StringVar()
            self.vars[key] = var
            ctk.CTkEntry(cell, textvariable=var, font=self.font_body,
                         fg_color=CARD, border_color=CARD_BORDER,
                         height=36).pack(fill="x", pady=(4, 0))

    # -------------------------------------------------------------- calc card
    def _build_calc_card(self):
        card = self._card("Lot Calculations",
                          "Worked out automatically — you can’t type in these.")
        grid = ctk.CTkFrame(card, fg_color="transparent")
        grid.pack(fill="x", padx=20, pady=(12, 18))

        self.lot_calc_labels = {}
        tiles = [
            ("meters_per_kg", "Meters per KG", False),
            ("total_meters", "Total Meters", False),
            ("fabric_cost", "Fabric Cost", True),
            ("total_cost", "Total Cost", True),
            ("base_cost_per_kg", "Base Cost per KG", True),
            ("wastage_cost_per_kg", "Wastage Cost per KG", True),
            ("adjusted_cost_per_kg", "Adjusted Cost per KG", True),
        ]
        cols = 4
        for i in range(cols):
            grid.grid_columnconfigure(i, weight=1, uniform="calc")
        for i, (key, label, _money) in enumerate(tiles):
            r, c = divmod(i, cols)
            accent = key == "adjusted_cost_per_kg"
            tile = ctk.CTkFrame(grid, fg_color=TEAL_TINT if accent else TILE_BG,
                                corner_radius=8)
            tile.grid(row=r, column=c, sticky="ew", padx=6, pady=6)
            ctk.CTkLabel(tile, text=label, font=self.font_label,
                         text_color=MUTED).pack(anchor="w", padx=14, pady=(12, 0))
            val = ctk.CTkLabel(tile, text="—", font=self.font_tile,
                               text_color=TEAL_HOVER if accent else TEXT)
            val.pack(anchor="w", padx=14, pady=(2, 12))
            self.lot_calc_labels[key] = val

    # ---------------------------------------------------------- products card
    def _build_products_card(self):
        card = self._card("Products",
                          "Type into Product, Weight and Pieces. Tick one row as wastage.")
        # Add Row button on the header line.
        header_bar = ctk.CTkFrame(card, fg_color="transparent")
        header_bar.pack(fill="x", padx=20, pady=(8, 0))
        self._button(header_bar, "+ Add Row",
                     lambda: (self._add_row(Product(name="")), self.recompute())
                     ).pack(side="right")

        # Column header row (fixed, above the scroll area) — same widths as rows.
        head = ctk.CTkFrame(card, fg_color="transparent")
        head.pack(fill="x", padx=20, pady=(10, 0))
        headers = [("name", "Product"), ("weight", "Weight (kg)"),
                   ("pieces", "Pieces"), ("wastage", "Wastage"),
                   ("wtpc", "Wt./Pc"), ("costpc", "Cost/Pc"),
                   ("revenue", "Revenue"), ("del", "")]
        for key, text in headers:
            anchor = "w" if key == "name" else ("center" if key == "wastage" else "e")
            ctk.CTkLabel(head, text=text, width=COL[key], anchor=anchor,
                         font=self.font_label, text_color=MUTED).pack(
                side="left", padx=4)

        self.table = ctk.CTkScrollableFrame(card, fg_color="transparent", height=260)
        self.table.pack(fill="x", padx=16, pady=(4, 16))

    # ---------------------------------------------------------- warning label
    def _build_warning(self):
        self.warning = ctk.CTkLabel(
            self.page, text="", font=self.font_sumb, text_color=ORANGE,
            anchor="w", justify="left", wraplength=1000)
        # Packed/unpacked by _refresh_warning; sits above the summary card.

    # ---------------------------------------------------------- summary card
    def _build_summary_card(self):
        card = self._card("Summary")
        self.summary_card = card
        box = ctk.CTkFrame(card, fg_color="transparent")
        box.pack(fill="x", padx=20, pady=(12, 16))
        box.grid_columnconfigure(0, weight=1)

        self.total_labels = {}
        rows = [
            ("total_weight", "Total Weight Produced (kg)"),
            ("total_pieces", "Total Pieces"),
            ("total_payment", "Total Payment (Cost)"),
            ("total_receipt", "Total Receipt (Revenue)"),
        ]
        for i, (key, label) in enumerate(rows):
            ctk.CTkLabel(box, text=label, font=self.font_sum,
                         text_color=MUTED).grid(row=i, column=0, sticky="w", pady=6)
            val = ctk.CTkLabel(box, text="—", font=self.font_sumb, text_color=TEXT)
            val.grid(row=i, column=1, sticky="e", pady=6)
            self.total_labels[key] = val

        ctk.CTkLabel(box, text="Profit", font=self.font_h2,
                     text_color=TEXT).grid(row=len(rows), column=0, sticky="w",
                                           pady=(12, 4))
        self.profit_label = ctk.CTkLabel(box, text="—", font=self.font_profit,
                                         text_color=TEXT)
        self.profit_label.grid(row=len(rows), column=1, sticky="e", pady=(12, 4))

    # --------------------------------------------------------------- products
    def _add_row(self, product=None):
        product = product or Product(name="")
        row = ProductRow(self.table, self, product)
        row.frame.pack(fill="x")
        self.rows.append(row)

    def _delete_row(self, row):
        row.destroy()
        self.rows.remove(row)
        self.recompute()

    def _on_wastage(self, row):
        """Enforce a single wastage row: checking one clears the others."""
        if row.is_wastage():
            for other in self.rows:
                if other is not row and other.is_wastage():
                    other.wastage_var.set(0)
                    other._apply_wastage_style(False)
        row._apply_wastage_style(row.is_wastage())
        self.recompute()

    def _ordered_products(self):
        return [row.to_product() for row in self.rows]

    # ------------------------------------------------------------- recompute
    def _current_lot(self):
        v = self.vars
        return LotInfo(
            reference=v["reference"].get().strip(),
            fabric_type=v["fabric_type"].get().strip(),
            date=v["date"].get().strip(),
            gsm=v["gsm"].get(),
            width_in=v["width_in"].get(),
            total_kg=v["total_kg"].get(),
            rate_per_meter=v["rate_per_meter"].get(),
            transport_cost=v["transport_cost"].get(),
        )

    def recompute(self, *_):
        """Recompute everything and repaint every derived field on screen."""
        lot = self._current_lot()
        products = self._ordered_products()
        results = calculations.compute(lot, products)

        money_keys = {"fabric_cost", "total_cost", "base_cost_per_kg",
                      "wastage_cost_per_kg", "adjusted_cost_per_kg"}
        for key, label in self.lot_calc_labels.items():
            label.configure(text=fmt(getattr(results, key), money=key in money_keys))

        for row, pr in zip(self.rows, results.product_results):
            row.paint(pr)

        self.total_labels["total_weight"].configure(
            text=fmt(results.total_weight_produced))
        self.total_labels["total_pieces"].configure(
            text=fmt(results.total_pieces, decimals=0))
        self.total_labels["total_payment"].configure(
            text=fmt(results.total_payment, money=True))
        self.total_labels["total_receipt"].configure(
            text=fmt(results.total_receipt, money=True))

        if results.profit is None:
            self.profit_label.configure(text="—", text_color=TEXT)
        else:
            self.profit_label.configure(
                text=fmt(results.profit, money=True),
                text_color=GREEN if results.profit >= 0 else RED)

        self._refresh_warning(results)

    def _refresh_warning(self, results):
        if results.recon_mismatch:
            self.warning.configure(text=results.recon_message)
            if not self.warning.winfo_ismapped():
                self.warning.pack(fill="x", padx=6, pady=(0, 10),
                                  before=self.summary_card)
        else:
            self.warning.configure(text="")
            if self.warning.winfo_ismapped():
                self.warning.pack_forget()

    # -------------------------------------------------------------- toolbar ops
    def _clear_rows(self):
        for row in self.rows:
            row.destroy()
        self.rows.clear()

    def _new_lot(self, confirm=True):
        if confirm and not messagebox.askyesno(
                "New Lot", "Clear all fields and start a new lot?"):
            return
        for key, var in self.vars.items():
            var.set(date.today().isoformat() if key == "date" else "")
        self._clear_rows()
        # Seed with two blank rows plus a wastage row for convenience.
        self._add_row(Product(name=""))
        self._add_row(Product(name=""))
        self._add_row(Product(name="Wastage", is_wastage=True))
        self.recompute()

    def _load_into_ui(self, lot, products):
        self.vars["reference"].set(lot.reference or "")
        self.vars["fabric_type"].set(lot.fabric_type or "")
        self.vars["date"].set(lot.date or "")
        for key in ("gsm", "width_in", "total_kg", "rate_per_meter", "transport_cost"):
            val = getattr(lot, key)
            self.vars[key].set("" if val is None else str(val))
        self._clear_rows()
        for p in products:
            self._add_row(p)
        self.recompute()

    def _save_lot(self):
        lot = self._current_lot()
        products = self._ordered_products()
        path = filedialog.asksaveasfilename(
            title="Save Lot", defaultextension=".json",
            initialdir=str(storage.lots_dir()),
            initialfile=storage.default_filename(lot, "json"),
            filetypes=[("Lot files", "*.json")],
        )
        if not path:
            return
        try:
            storage.save_lot(lot, products, path)
            messagebox.showinfo("Save Lot", "Lot saved successfully.")
        except Exception as exc:  # noqa: BLE001 — surface any IO error friendly
            messagebox.showerror("Save Lot", f"Could not save the lot:\n{exc}")

    def _open_lot(self):
        path = filedialog.askopenfilename(
            title="Open Lot", initialdir=str(storage.lots_dir()),
            filetypes=[("Lot files", "*.json"), ("All files", "*.*")],
        )
        if not path:
            return
        try:
            lot, products = storage.load_lot(path)
            self._load_into_ui(lot, products)
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("Open Lot", f"Could not open the lot:\n{exc}")

    def _export_excel(self):
        lot = self._current_lot()
        products = self._ordered_products()
        results = calculations.compute(lot, products)
        path = filedialog.asksaveasfilename(
            title="Export to Excel", defaultextension=".xlsx",
            initialfile=storage.default_filename(lot, "xlsx"),
            filetypes=[("Excel files", "*.xlsx")],
        )
        if not path:
            return
        try:
            excel_export.export(results, lot, path)
            messagebox.showinfo("Export to Excel", "Excel file created successfully.")
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("Export to Excel", f"Could not export:\n{exc}")


def main():
    app = CostingApp()
    app.mainloop()


if __name__ == "__main__":
    main()
