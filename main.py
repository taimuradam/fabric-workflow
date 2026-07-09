"""
main.py — Textile Lot Costing desktop app (tkinter GUI + entry point).

Single-window layout, top to bottom:
    Toolbar (New / Save / Open / Export)
    Lot Info panel + live lot calculations
    Products table (inline-editable ttk.Treeview) + Add/Delete buttons
    Reconciliation warning (only when weights don't add up)
    Totals summary (bold)

All heavy lifting lives in calculations.py; this file is purely presentation and
wiring. Every input change funnels into a single recompute() so the whole screen
stays consistent, and all number parsing is tolerant so bad input never crashes.
"""

import tkinter as tk
from datetime import date
from tkinter import filedialog, messagebox, ttk

import calculations
import excel_export
import storage
from calculations import LotInfo, Product

# Colours for the profit figure and the reconciliation warning.
_GREEN = "#1B7F3B"
_RED = "#C62828"
_ORANGE = "#C55A11"

# Product table columns: (internal id, heading, width, editable?)
_COLUMNS = [
    ("name", "Product", 200, True),
    ("weight", "Weight (kg)", 100, True),
    ("pieces", "Pieces", 80, True),
    ("wastage", "Wastage?", 80, False),   # toggled by clicking, not typed
    ("wtpc", "Wt./Pc", 90, False),
    ("costpc", "Cost/Pc", 110, False),
    ("revenue", "Revenue", 120, False),
]


def fmt(value, money=False, decimals=2):
    """Format a computed value for display; None -> em dash."""
    if value is None:
        return "—"
    if money:
        return f"Rs {value:,.{decimals}f}"
    return f"{value:,.{decimals}f}"


class CostingApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Textile Lot Costing")
        self.geometry("1024x760")
        self.minsize(900, 640)

        # iid -> Product mapping; tree child order defines product order.
        self.row_products = {}
        self._row_counter = 0
        self._editor = None  # active inline-edit Entry, if any

        self._build_toolbar()
        self._build_lot_panel()
        self._build_products_panel()
        self._build_warning_and_totals()

        self._new_lot(confirm=False)  # start with a clean, dated sheet

    # ------------------------------------------------------------------ layout
    def _build_toolbar(self):
        bar = ttk.Frame(self, padding=(12, 10))
        bar.pack(fill="x")
        for text, cmd in [
            ("New Lot", self._new_lot),
            ("Save Lot", self._save_lot),
            ("Open Lot", self._open_lot),
            ("Export to Excel", self._export_excel),
        ]:
            ttk.Button(bar, text=text, command=cmd).pack(side="left", padx=(0, 8))

    def _build_lot_panel(self):
        frame = ttk.LabelFrame(self, text="Lot Information", padding=12)
        frame.pack(fill="x", padx=12, pady=(0, 10))

        # --- editable inputs (left) ---
        self.vars = {}
        input_fields = [
            ("reference", "Lot Reference"),
            ("fabric_type", "Fabric Type"),
            ("date", "Date"),
            ("gsm", "GSM"),
            ("width_in", "Width (inches)"),
            ("total_kg", "Total KG Received"),
            ("rate_per_meter", "Rate per Meter"),
            ("transport_cost", "Transport Cost"),
        ]
        inputs = ttk.Frame(frame)
        inputs.grid(row=0, column=0, sticky="nw", padx=(0, 30))
        for i, (key, label) in enumerate(input_fields):
            r, c = divmod(i, 2)
            cell = ttk.Frame(inputs)
            cell.grid(row=r, column=c, sticky="w", padx=(0, 24), pady=6)
            ttk.Label(cell, text=label).pack(anchor="w")
            var = tk.StringVar()
            var.trace_add("write", lambda *_: self.recompute())
            ttk.Entry(cell, textvariable=var, width=22).pack(anchor="w")
            self.vars[key] = var

        # --- read-only lot calculations (right) ---
        calc = ttk.LabelFrame(frame, text="Lot Calculations", padding=10)
        calc.grid(row=0, column=1, sticky="nw")
        self.lot_calc_labels = {}
        calc_fields = [
            ("meters_per_kg", "Meters per KG"),
            ("total_meters", "Total Meters"),
            ("fabric_cost", "Fabric Cost"),
            ("total_cost", "Total Cost"),
            ("base_cost_per_kg", "Base Cost per KG"),
            ("wastage_cost_per_kg", "Wastage Cost per KG"),
            ("adjusted_cost_per_kg", "Adjusted Cost per KG"),
        ]
        for i, (key, label) in enumerate(calc_fields):
            ttk.Label(calc, text=label + ":").grid(row=i, column=0, sticky="w", pady=2)
            val = ttk.Label(calc, text="—", width=16, anchor="e")
            val.grid(row=i, column=1, sticky="e", padx=(12, 0), pady=2)
            self.lot_calc_labels[key] = val

    def _build_products_panel(self):
        frame = ttk.LabelFrame(self, text="Products", padding=12)
        frame.pack(fill="both", expand=True, padx=12, pady=(0, 10))

        btns = ttk.Frame(frame)
        btns.pack(fill="x", pady=(0, 8))
        ttk.Button(btns, text="Add Row", command=self._add_row).pack(side="left")
        ttk.Button(btns, text="Delete Selected Row",
                   command=self._delete_row).pack(side="left", padx=(8, 0))
        ttk.Label(
            btns,
            text="Double-click a Product / Weight / Pieces cell to edit. "
                 "Click the Wastage cell to mark the waste row.",
            foreground="#666",
        ).pack(side="left", padx=(16, 0))

        table = ttk.Frame(frame)
        table.pack(fill="both", expand=True)

        columns = [c[0] for c in _COLUMNS]
        self.tree = ttk.Treeview(table, columns=columns, show="headings", height=8)
        for key, heading, width, _editable in _COLUMNS:
            self.tree.heading(key, text=heading)
            anchor = "w" if key == "name" else "center"
            self.tree.column(key, width=width, anchor=anchor, stretch=(key == "name"))

        vsb = ttk.Scrollbar(table, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        # Highlight the wastage row.
        self.tree.tag_configure("wastage", background="#FCE8D5")

        self.tree.bind("<Double-1>", self._on_double_click)
        self.tree.bind("<Button-1>", self._on_single_click)

    def _build_warning_and_totals(self):
        # Reconciliation warning (packed/unpacked as needed).
        self.warning_var = tk.StringVar(value="")
        self.warning_label = tk.Label(
            self, textvariable=self.warning_var, fg=_ORANGE,
            font=("TkDefaultFont", 11, "bold"), anchor="w", justify="left",
        )
        # Not packed yet; _refresh_warning controls visibility.

        totals = ttk.LabelFrame(self, text="Summary", padding=12)
        totals.pack(fill="x", padx=12, pady=(0, 14))
        self.totals_frame = totals

        self.total_labels = {}
        big = ("TkDefaultFont", 12, "bold")
        rows = [
            ("total_weight", "Total Weight Produced (kg)"),
            ("total_pieces", "Total Pieces"),
            ("total_payment", "Total Payment (Cost)"),
            ("total_receipt", "Total Receipt (Revenue)"),
        ]
        for i, (key, label) in enumerate(rows):
            ttk.Label(totals, text=label + ":", font=big).grid(
                row=i, column=0, sticky="w", pady=3)
            val = ttk.Label(totals, text="—", font=big, anchor="e", width=18)
            val.grid(row=i, column=1, sticky="e", padx=(16, 0), pady=3)
            self.total_labels[key] = val

        # Profit gets its own coloured, larger label.
        ttk.Label(totals, text="Profit:", font=("TkDefaultFont", 14, "bold")).grid(
            row=len(rows), column=0, sticky="w", pady=(8, 0))
        self.profit_label = tk.Label(
            totals, text="—", font=("TkDefaultFont", 14, "bold"), anchor="e", width=18)
        self.profit_label.grid(row=len(rows), column=1, sticky="e",
                               padx=(16, 0), pady=(8, 0))

    # --------------------------------------------------------------- products
    def _add_row(self, product=None):
        """Append a product row (blank, or from a given Product)."""
        product = product or Product(name="", weight_kg=None, pieces=None)
        iid = f"row{self._row_counter}"
        self._row_counter += 1
        self.row_products[iid] = product
        self.tree.insert("", "end", iid=iid, values=self._row_values(product))
        self._apply_row_tags(iid)
        self.recompute()

    def _delete_row(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Delete Row", "Select a row to delete first.")
            return
        for iid in sel:
            self.tree.delete(iid)
            self.row_products.pop(iid, None)
        self.recompute()

    def _row_values(self, product, result=None):
        """Build the 7 display cell values for a product row."""
        wastage = "✓" if product.is_wastage else ""
        if product.is_wastage:
            wtpc = costpc = revenue = "—"
        elif result is not None:
            wtpc = "N/A" if result.weight_per_piece is None else fmt(result.weight_per_piece)
            costpc = fmt(result.cost_per_piece, money=True)
            revenue = fmt(result.revenue, money=True)
        else:
            wtpc = costpc = revenue = "—"
        return [
            product.name,
            "" if product.weight_kg in (None, "") else product.weight_kg,
            "" if product.pieces in (None, "") else product.pieces,
            wastage,
            wtpc,
            costpc,
            revenue,
        ]

    def _apply_row_tags(self, iid):
        product = self.row_products[iid]
        self.tree.item(iid, tags=("wastage",) if product.is_wastage else ())

    def _ordered_products(self):
        return [self.row_products[iid] for iid in self.tree.get_children()]

    # ------------------------------------------------------- inline editing
    def _on_single_click(self, event):
        """Toggle the wastage flag when its cell is clicked."""
        if self.tree.identify_region(event.x, event.y) != "cell":
            return
        col = self.tree.identify_column(event.x)  # like "#4"
        iid = self.tree.identify_row(event.y)
        if not iid:
            return
        col_index = int(col[1:]) - 1
        if _COLUMNS[col_index][0] != "wastage":
            return
        # Enforce single wastage row: turning one on clears the rest.
        turning_on = not self.row_products[iid].is_wastage
        for other_iid, product in self.row_products.items():
            product.is_wastage = turning_on and other_iid == iid
            self._apply_row_tags(other_iid)
        self.recompute()

    def _on_double_click(self, event):
        """Open an inline Entry over an editable text cell."""
        if self.tree.identify_region(event.x, event.y) != "cell":
            return
        col = self.tree.identify_column(event.x)
        iid = self.tree.identify_row(event.y)
        if not iid:
            return
        col_index = int(col[1:]) - 1
        key, _heading, _width, editable = _COLUMNS[col_index]
        if not editable:
            return

        x, y, w, h = self.tree.bbox(iid, col)
        current = self.tree.set(iid, key)
        self._destroy_editor()

        entry = ttk.Entry(self.tree)
        entry.place(x=x, y=y, width=w, height=h)
        entry.insert(0, current)
        entry.focus_set()
        entry.select_range(0, "end")

        def commit(_event=None):
            self._commit_edit(iid, key, entry.get())
            self._destroy_editor()

        entry.bind("<Return>", commit)
        entry.bind("<KP_Enter>", commit)
        entry.bind("<FocusOut>", commit)
        entry.bind("<Escape>", lambda e: self._destroy_editor())
        self._editor = entry

    def _commit_edit(self, iid, key, raw):
        product = self.row_products.get(iid)
        if product is None:
            return
        raw = raw.strip()
        if key == "name":
            product.name = raw
        elif key == "weight":
            product.weight_kg = raw if raw != "" else None
        elif key == "pieces":
            product.pieces = raw if raw != "" else None
        self.recompute()

    def _destroy_editor(self):
        if self._editor is not None:
            self._editor.destroy()
            self._editor = None

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

        # Lot calculation labels.
        money_keys = {"fabric_cost", "total_cost", "base_cost_per_kg",
                      "wastage_cost_per_kg", "adjusted_cost_per_kg"}
        for key, label in self.lot_calc_labels.items():
            label.config(text=fmt(getattr(results, key), money=key in money_keys))

        # Per-row computed columns.
        for iid, pr in zip(self.tree.get_children(), results.product_results):
            product = self.row_products[iid]
            self.tree.item(iid, values=self._row_values(product, pr))
            self._apply_row_tags(iid)

        # Totals.
        self.total_labels["total_weight"].config(text=fmt(results.total_weight_produced))
        self.total_labels["total_pieces"].config(text=fmt(results.total_pieces, decimals=0))
        self.total_labels["total_payment"].config(text=fmt(results.total_payment, money=True))
        self.total_labels["total_receipt"].config(text=fmt(results.total_receipt, money=True))

        if results.profit is None:
            self.profit_label.config(text="—", fg="black")
        else:
            self.profit_label.config(
                text=fmt(results.profit, money=True),
                fg=_GREEN if results.profit >= 0 else _RED,
            )

        self._refresh_warning(results)

    def _refresh_warning(self, results):
        if results.recon_mismatch:
            self.warning_var.set(results.recon_message)
            if not self.warning_label.winfo_ismapped():
                # Show it right above the Summary frame.
                self.warning_label.pack(fill="x", padx=14, pady=(0, 6),
                                        before=self.totals_frame)
        else:
            self.warning_var.set("")
            if self.warning_label.winfo_ismapped():
                self.warning_label.pack_forget()

    # -------------------------------------------------------------- toolbar ops
    def _new_lot(self, confirm=True):
        if confirm and not messagebox.askyesno(
                "New Lot", "Clear all fields and start a new lot?"):
            return
        self._destroy_editor()
        for key, var in self.vars.items():
            var.set(date.today().isoformat() if key == "date" else "")
        for iid in list(self.tree.get_children()):
            self.tree.delete(iid)
        self.row_products.clear()
        # Seed with a few blank rows incl. a wastage row for convenience.
        self._add_row(Product(name="", weight_kg=None, pieces=None))
        self._add_row(Product(name="", weight_kg=None, pieces=None))
        self._add_row(Product(name="Wastage", weight_kg=None, pieces=None, is_wastage=True))
        self.recompute()

    def _load_into_ui(self, lot, products):
        self._destroy_editor()
        self.vars["reference"].set(lot.reference or "")
        self.vars["fabric_type"].set(lot.fabric_type or "")
        self.vars["date"].set(lot.date or "")
        for key in ("gsm", "width_in", "total_kg", "rate_per_meter", "transport_cost"):
            val = getattr(lot, key)
            self.vars[key].set("" if val is None else str(val))
        for iid in list(self.tree.get_children()):
            self.tree.delete(iid)
        self.row_products.clear()
        for p in products:
            self._add_row(p)
        self.recompute()

    def _save_lot(self):
        lot = self._current_lot()
        products = self._ordered_products()
        path = filedialog.asksaveasfilename(
            title="Save Lot",
            defaultextension=".json",
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
            title="Open Lot",
            initialdir=str(storage.lots_dir()),
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
            title="Export to Excel",
            defaultextension=".xlsx",
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
