"""
storage.py — Save/load a lot's inputs to/from a local JSON file.

Only the raw INPUTS are persisted (lot info + product rows); everything else is
recomputed on load via calculations.compute, so a saved file can never contain
stale derived numbers.
"""

import json
import sys
from dataclasses import asdict
from pathlib import Path
from typing import List, Tuple

from calculations import LotInfo, Product


def app_dir() -> Path:
    """Directory the app lives in — works both as a script and a frozen exe.

    When PyInstaller freezes the app, sys.argv[0] points at the .exe, so this
    resolves to the folder the user double-clicked from. In dev it's the repo.
    """
    return Path(sys.argv[0]).resolve().parent


def lots_dir() -> Path:
    """The 'lots' subfolder next to the app, created on demand."""
    d = app_dir() / "lots"
    d.mkdir(parents=True, exist_ok=True)
    return d


def default_filename(lot: LotInfo, extension: str) -> str:
    """Suggested filename like 'CVC-3_2026-07-09.xlsx'.

    Falls back to 'lot' / 'nodate' when those fields are blank, and strips
    characters that are illegal in filenames.
    """
    ref = (lot.reference or "lot").strip() or "lot"
    date = (lot.date or "nodate").strip() or "nodate"
    stem = f"{ref}_{date}"
    for bad in '<>:"/\\|?*':
        stem = stem.replace(bad, "-")
    return f"{stem}.{extension.lstrip('.')}"


def save_lot(lot: LotInfo, products: List[Product], path: str) -> None:
    """Write lot info + product rows to a JSON file."""
    data = {
        "version": 1,
        "lot": asdict(lot),
        "products": [asdict(p) for p in products],
    }
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)


def load_lot(path: str) -> Tuple[LotInfo, List[Product]]:
    """Read a JSON file back into a (LotInfo, [Product]) pair.

    Uses per-field .get() so a file written by an older/newer version with
    missing or extra keys still loads instead of throwing.
    """
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)

    lot_data = data.get("lot", {})
    lot = LotInfo(
        reference=lot_data.get("reference", ""),
        fabric_type=lot_data.get("fabric_type", ""),
        date=lot_data.get("date", ""),
        gsm=lot_data.get("gsm"),
        width_in=lot_data.get("width_in"),
        total_kg=lot_data.get("total_kg"),
        rate_per_meter=lot_data.get("rate_per_meter"),
        transport_cost=lot_data.get("transport_cost"),
    )

    products: List[Product] = []
    for pd in data.get("products", []):
        products.append(
            Product(
                name=pd.get("name", ""),
                weight_kg=pd.get("weight_kg"),
                pieces=pd.get("pieces"),
                is_wastage=bool(pd.get("is_wastage", False)),
            )
        )

    return lot, products
