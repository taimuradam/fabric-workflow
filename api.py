"""
api.py — The Python <-> JavaScript bridge for the pywebview frontend.

An instance of Api is handed to pywebview as `js_api`, which exposes its public
methods to the browser as `window.pywebview.api.<method>(...)` (each returns a
JS Promise). This layer does NOTHING clever with the numbers — it just marshals
plain dicts/lists from JS into the existing dataclasses, calls the unchanged
calculations / storage / excel_export modules, and marshals the results back.

Keeping every business rule in calculations.py means the pywebview UI produces
byte-for-byte the same figures the old tkinter UI did.
"""

from dataclasses import asdict

import webview

import calculations
import excel_export
import storage
from calculations import LotInfo, Product


def _lot_from_dict(d: dict) -> LotInfo:
    """Build a LotInfo from the plain dict the frontend sends.

    Numeric fields arrive as strings (or None) straight from the HTML inputs;
    calculations.parse_number handles the cleanup, so we pass them through as-is.
    """
    d = d or {}
    return LotInfo(
        reference=d.get("reference", "") or "",
        fabric_type=d.get("fabric_type", "") or "",
        date=d.get("date", "") or "",
        gsm=d.get("gsm"),
        width_in=d.get("width_in"),
        total_kg=d.get("total_kg"),
        rate_per_meter=d.get("rate_per_meter"),
        transport_cost=d.get("transport_cost"),
    )


def _products_from_list(items) -> list:
    """Build the Product list from the frontend's array of row dicts."""
    products = []
    for p in items or []:
        products.append(
            Product(
                name=p.get("name", "") or "",
                weight_kg=p.get("weight_kg"),
                pieces=p.get("pieces"),
                is_wastage=bool(p.get("is_wastage", False)),
            )
        )
    return products


def _normalize_dialog_result(result):
    """create_file_dialog returns a str, a tuple/list of paths, or None.

    Collapse all of those to a single path string (or None if cancelled).
    """
    if not result:
        return None
    if isinstance(result, (list, tuple)):
        return result[0] if result else None
    return result


class Api:
    """Methods on this class are callable from JavaScript via pywebview."""

    def __init__(self):
        # Set by main.py right after the window is created; needed for the
        # native file dialogs.
        self.window = None

    # ---------------------------------------------------------------- compute
    def compute(self, lot_dict, products_list):
        """Run the costing calculation and return a JSON-serializable dict.

        dataclasses.asdict recursively converts Results (incl. the nested
        ProductResult -> Product) into plain dicts the frontend can read.
        """
        lot = _lot_from_dict(lot_dict)
        products = _products_from_list(products_list)
        results = calculations.compute(lot, products)
        return asdict(results)

    # ------------------------------------------------------------- save / load
    def save_lot(self, lot_dict, products_list):
        """Open a native Save dialog and write the lot to JSON."""
        lot = _lot_from_dict(lot_dict)
        products = _products_from_list(products_list)
        try:
            path = _normalize_dialog_result(
                self.window.create_file_dialog(
                    webview.SAVE_DIALOG,
                    directory=str(storage.lots_dir()),
                    save_filename=storage.default_filename(lot, "json"),
                    file_types=("Lot files (*.json)", "All files (*.*)"),
                )
            )
            if not path:
                return {"ok": False, "cancelled": True}
            if not path.lower().endswith(".json"):
                path += ".json"
            storage.save_lot(lot, products, path)
            return {"ok": True, "path": path}
        except Exception as exc:  # noqa: BLE001 — report to the UI, never crash
            return {"ok": False, "error": str(exc)}

    def load_lot(self):
        """Open a native Open dialog and return the loaded lot as plain dicts."""
        try:
            path = _normalize_dialog_result(
                self.window.create_file_dialog(
                    webview.OPEN_DIALOG,
                    directory=str(storage.lots_dir()),
                    allow_multiple=False,
                    file_types=("Lot files (*.json)", "All files (*.*)"),
                )
            )
            if not path:
                return {"ok": False, "cancelled": True}
            lot, products = storage.load_lot(path)
            return {
                "ok": True,
                "lot": asdict(lot),
                "products": [asdict(p) for p in products],
            }
        except Exception as exc:  # noqa: BLE001
            return {"ok": False, "error": str(exc)}

    # -------------------------------------------------------------- excel export
    def export_excel(self, lot_dict, products_list, results_dict=None):
        """Open a native Save dialog and write the formatted .xlsx.

        results_dict is accepted for API symmetry with the frontend but is not
        trusted for the numbers — we recompute from the raw lot/products so the
        spreadsheet always reflects the exact same calculation the app uses.
        """
        lot = _lot_from_dict(lot_dict)
        products = _products_from_list(products_list)
        try:
            results = calculations.compute(lot, products)
            path = _normalize_dialog_result(
                self.window.create_file_dialog(
                    webview.SAVE_DIALOG,
                    directory=str(storage.lots_dir()),
                    save_filename=storage.default_filename(lot, "xlsx"),
                    file_types=("Excel files (*.xlsx)", "All files (*.*)"),
                )
            )
            if not path:
                return {"ok": False, "cancelled": True}
            if not path.lower().endswith(".xlsx"):
                path += ".xlsx"
            excel_export.export(results, lot, path)
            return {"ok": True, "path": path}
        except Exception as exc:  # noqa: BLE001
            return {"ok": False, "error": str(exc)}
