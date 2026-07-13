"""
main.py — Textile Lot Costing desktop app (CustomTkinter GUI + entry point).

Layout per the design handoff (fixed 1200×900, three columns, no page scroll):

    Title bar   — app mark + "Fabric Costing Calculator"
    Toolbar     — "Lot Costing" + live lot context · New / Open / Save / Export
    Body (grid, 3 columns)
      LEFT  (300px)  Fabric Lot card — all lot inputs, incl. the Wastage field
      CENTER (flex)  Products card — header row + rows (ONLY the rows scroll)
                     + reconciliation status bar
      RIGHT (296px)  Cost breakdown card + Lot cost summary card

All heavy lifting lives in calculations.py; this file is purely presentation and
wiring. Every input change funnels into a single recompute() so the whole screen
stays consistent, and all number parsing is tolerant so bad input never crashes.

Wastage is entered as a single field in the LEFT card (not a per-row checkbox).
Before every compute()/save/export, a wastage Product is synthesised from that
field, so calculations.py and the on-disk JSON format are unchanged — old saved
lots (wastage as an is_wastage row) load correctly, and new saves remain readable
by older builds.
"""

import os
import sys
from datetime import date, datetime
from pathlib import Path
from tkinter import filedialog, messagebox

import customtkinter as ctk

import calculations
import excel_export
import storage
from calculations import LotInfo, Product

# ------------------------------------------------------------ design tokens ---
# Exact palette from the design handoff README.
TEAL             = "#0f766e"   # primary button fill, section-label accents
TEAL_HOVER       = "#0b5f58"   # primary hover, cost-per-piece + adjusted-cost text
TEAL_TINT        = "#dcebe8"   # cost cell, adjusted tile, "+ Add product" fill
TEAL_TINT_BORDER = "#bcd7d1"   # border of tinted teal buttons/cells
BG               = "#eef0ec"   # window background, toolbar
CARD             = "#ffffff"   # title bar, cards, entry fill
CARD_BORDER      = "#dde2dd"   # card borders / hairlines
INPUT_BORDER     = "#b9c4bd"   # editable-entry outlines — deliberately darker
                               # than CARD_BORDER: a 1px near-white hairline
                               # disappears entirely on some monitors and the
                               # operator couldn't see where to type
HAIRLINE         = "#eef0ec"   # inner row separators / dividers
ROW_LINE         = "#f2f4f1"   # product-row bottom border
TEXT             = "#1f2a30"   # primary text
MUTED            = "#67757c"   # labels, secondary text
FAINT            = "#9aa6ab"   # unit suffixes, column headers
PLACEHOLDER      = "#c0c8c1"   # empty-row borders & em-dashes
SECONDARY        = "#e4e8e3"   # secondary button fill
SECONDARY_H      = "#d6dbd5"   # secondary button hover
GREEN            = "#15803d"   # reconciliation "adds up"
WASTE_BG         = "#fbf1ea"   # wastage entry fill
WASTE_BORDER     = "#e6c9b6"   # wastage entry border
WASTE_TEXT       = "#c05621"   # wastage labels / values, mismatch warning
RECON_BG         = "#f2f7f3"   # reconciliation status bar fill

# Products table column sizing, shared by the header row and every product row so
# the grids line up (CTk has no Treeview; this is the manual-table approach).
#   col 0 Product (expands) · 1 Weight · 2 Pieces · 3 Wt/Pc · 4 Cost/pc · 5 ✕
COL_MIN = {1: 84, 2: 68, 3: 62, 4: 122, 5: 30}
# Extra width is shared, not dumped on the name column: Product grows 3× while
# Weight / Pieces / Cost each grow 1× (keeps fullscreen proportions sane).
COL_WEIGHT = {0: 3, 1: 1, 2: 1, 3: 0, 4: 1, 5: 0}


def fmt(value, money=False, decimals=2):
    """Format a computed value for display; None -> em dash."""
    if value is None:
        return "—"
    if money:
        return f"Rs {value:,.{decimals}f}"
    return f"{value:,.{decimals}f}"


def kgfmt(value):
    """Compact weight for sentences: 500.0 -> '500', 18.5 -> '18.5'."""
    return f"{value:g}"


def resource_path(relative: str) -> str:
    """Absolute path to a bundled asset, both as a script and as a frozen exe.

    PyInstaller unpacks --add-data files into a temp dir exposed as
    sys._MEIPASS; when running from source, fall back to this file's folder.
    """
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, relative)


# --------------------------------------------------------- app data folder ---
# NOTHING is written next to the .exe. Everything lives in one user-chosen
# "Fabric Lot Files" folder: settings.json plus "Saved Lots" (lot JSONs) and
# "Excel Reports" (default export destination) subfolders. The only thing
# outside it is a one-line locator in the OS per-user config area (%APPDATA%
# on Windows) that records where that folder is — without it the app couldn't
# find its own settings at launch.

def _locator_path() -> Path:
    import os
    if os.name == "nt":
        base = Path(os.environ.get("APPDATA", str(Path.home())))
    else:
        base = Path.home() / ".config"
    d = base / "Fabric Costing"
    d.mkdir(parents=True, exist_ok=True)
    return d / "location.txt"


def read_locator():
    """The recorded Fabric Lot Files folder, or None if unset/deleted."""
    try:
        text = _locator_path().read_text(encoding="utf-8").strip()
        if text and Path(text).is_dir():
            return Path(text)
    except Exception:  # noqa: BLE001
        pass
    return None


def write_locator(data_dir: Path) -> None:
    try:
        _locator_path().write_text(str(data_dir), encoding="utf-8")
    except Exception:  # noqa: BLE001 — non-critical; app still works this run
        pass


def load_settings(data_dir) -> dict:
    """Preferences from <Fabric Lot Files>/settings.json ({} if unavailable)."""
    if not data_dir:
        return {}
    try:
        import json
        with open(Path(data_dir) / "settings.json", "r", encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except Exception:  # noqa: BLE001 — missing/corrupt file: start fresh
        return {}


def save_settings(data_dir, settings: dict) -> None:
    if not data_dir:
        return
    try:
        import json
        with open(Path(data_dir) / "settings.json", "w", encoding="utf-8") as f:
            json.dump(settings, f, indent=2)
    except Exception:  # noqa: BLE001 — non-critical; never block the export
        pass


def report_filename(lot) -> str:
    """Human-friendly report name, e.g. 'Lot CVC-3 - CVC - 2026-07-09.xlsx'.

    Lot reference first so every export of the same lot sorts together in the
    folder; ISO date so repeated exports line up chronologically.
    """
    parts = ["Lot " + ((lot.reference or "").strip() or "untitled")]
    if (lot.fabric_type or "").strip():
        parts.append(lot.fabric_type.strip())
    if (lot.date or "").strip():
        parts.append(lot.date.strip())
    stem = " - ".join(parts)
    for bad in '<>:"/\\|?*':
        stem = stem.replace(bad, "-")
    return f"{stem}.xlsx"


def configure_table_grid(frame):
    """Apply the shared products-table column scheme to a grid container."""
    for col in range(6):
        frame.grid_columnconfigure(col, weight=COL_WEIGHT[col],
                                   minsize=COL_MIN.get(col, 0))


class UnitEntry(ctk.CTkFrame):
    """An entry with a faint unit hint inside it (kg / Rs / in).

    CTkEntry has no inner padding, so overlaying a label risks collisions.
    Instead the *frame* draws the entry look (fill, border, radius) and packs a
    borderless entry beside a small FAINT unit label — the README-sanctioned
    nearest-clean-CTk-equivalent.
    """

    def __init__(self, parent, textvariable, *, font, height=40, width=200,
                 prefix=None, suffix=None, justify="left", fg=CARD,
                 border=INPUT_BORDER, placeholder=None):
        # width matters: a CTkFrame defaults to 200px, and with
        # pack_propagate(False) that request would blow out fixed table columns.
        # border_width=2: a 1px hairline can vanish on some monitors after DPI
        # scaling rounds it away — 2px survives everywhere.
        super().__init__(parent, fg_color=fg, border_width=2,
                         border_color=border, corner_radius=6, height=height,
                         width=width)
        self.pack_propagate(False)
        unit_font = ctk.CTkFont(size=12)
        if prefix:
            ctk.CTkLabel(self, text=prefix, font=unit_font, text_color=FAINT,
                         width=1).pack(side="left", padx=(10, 0))
        self.entry = ctk.CTkEntry(self, textvariable=textvariable, font=font,
                                  fg_color=fg, border_width=0, justify=justify,
                                  placeholder_text=placeholder)
        self.entry.pack(side="left", fill="both", expand=True,
                        padx=(8 if not prefix else 4, 4), pady=2)
        if suffix:
            ctk.CTkLabel(self, text=suffix, font=unit_font, text_color=FAINT,
                         width=1).pack(side="right", padx=(0, 10))

    def set_border(self, color):
        self.configure(border_color=color)

    def bind_keyrelease(self, handler):
        self.entry.bind("<KeyRelease>", handler)


class ProductRow:
    """One editable product row: Product · Weight · Pieces · Wt/Pc · Cost/pc · ✕."""

    def __init__(self, table, app, product):
        self.app = app
        self.frame = ctk.CTkFrame(table, fg_color="transparent")
        configure_table_grid(self.frame)

        self.name_var = ctk.StringVar(value=product.name or "")
        self.weight_var = ctk.StringVar(
            value="" if product.weight_kg in (None, "") else str(product.weight_kg))
        self.pieces_var = ctk.StringVar(
            value="" if product.pieces in (None, "") else str(product.pieces))

        # width=60: request small so the weight=1 column stretches it; a larger
        # request would overflow the scroll canvas and break header alignment.
        self.name = UnitEntry(self.frame, self.name_var, font=app.font_body,
                              placeholder="Add a product…", width=60)
        self.name.grid(row=0, column=0, sticky="ew", padx=(0, 8), pady=(0, 6))

        self.weight = UnitEntry(self.frame, self.weight_var, font=app.font_body,
                                suffix="kg", justify="right", width=COL_MIN[1])
        self.weight.grid(row=0, column=1, sticky="ew", padx=(0, 8), pady=(0, 6))

        self.pieces = UnitEntry(self.frame, self.pieces_var, font=app.font_body,
                                justify="right", width=COL_MIN[2])
        self.pieces.grid(row=0, column=2, sticky="ew", padx=(0, 8), pady=(0, 6))

        # Wt / Pc — read-only computed label.
        self.wtpc = ctk.CTkLabel(self.frame, text="—", font=app.font_small,
                                 text_color=MUTED, anchor="e")
        self.wtpc.grid(row=0, column=3, sticky="ew", padx=(0, 8), pady=(0, 6))

        # Cost per piece — the visual hero: teal-tinted cell, 17 bold teal text.
        self.cost_cell = ctk.CTkFrame(self.frame, fg_color=TEAL_TINT,
                                      corner_radius=6, height=40,
                                      width=COL_MIN[4])
        self.cost_cell.grid(row=0, column=4, sticky="ew", padx=(0, 8), pady=(0, 6))
        self.cost_cell.pack_propagate(False)
        self.cost = ctk.CTkLabel(self.cost_cell, text="—", font=app.font_cost,
                                 text_color=TEAL_HOVER, anchor="e")
        self.cost.pack(fill="both", expand=True, padx=10)

        # Delete — always visible (non-technical user won't find hover-only).
        self.delete = ctk.CTkButton(
            self.frame, text="✕", width=COL_MIN[5], height=40, font=app.font_body,
            fg_color="transparent", text_color=MUTED, hover_color="#f6dede",
            command=lambda: app._delete_row(self))
        self.delete.grid(row=0, column=5, pady=(0, 6))

        for cell in (self.name, self.weight, self.pieces):
            cell.bind_keyrelease(app.recompute)

        # The row's own <Configure> is the reliable resize signal: it fires
        # AFTER the row has its new geometry (the card's fires before, so a
        # sync triggered there reads stale column widths).
        self.frame.bind("<Configure>", lambda e: app._schedule_header_sync())

    def to_product(self):
        w = self.weight_var.get().strip()
        p = self.pieces_var.get().strip()
        return Product(
            name=self.name_var.get().strip(),
            weight_kg=w if w != "" else None,
            pieces=p if p != "" else None,
            is_wastage=False,
        )

    def paint(self, result):
        """Write the computed cells; style empty rows as placeholders."""
        empty = not (self.name_var.get().strip() or self.weight_var.get().strip()
                     or self.pieces_var.get().strip())
        border = PLACEHOLDER if empty else INPUT_BORDER
        for cell in (self.name, self.weight, self.pieces):
            cell.set_border(border)

        if result is None or empty:
            self.wtpc.configure(text="—", text_color=PLACEHOLDER if empty else MUTED)
            self.cost.configure(text="—",
                                text_color=PLACEHOLDER if empty else TEAL_HOVER)
            return
        # pieces = 0 -> weight_per_piece is None -> "N/A" (unchanged semantics)
        if result.weight_per_piece is None:
            has_inputs = self.weight_var.get().strip() and self.pieces_var.get().strip()
            self.wtpc.configure(text="N/A" if has_inputs else "—", text_color=MUTED)
        else:
            self.wtpc.configure(text=f"{fmt(result.weight_per_piece)} kg",
                                text_color=MUTED)
        self.cost.configure(text=fmt(result.cost_per_piece, money=True),
                            text_color=TEAL_HOVER)

    def destroy(self):
        self.frame.destroy()


class CostingApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Fabric Costing Calculator")
        self._set_window_icon()
        self._header_sync_pending = False
        self._set_initial_geometry()
        self.minsize(1120, 840)
        ctk.set_appearance_mode("light")
        self.configure(fg_color=BG)

        # ------------------------------------------------------------- fonts
        self.font_title   = ctk.CTkFont(size=19, weight="bold")   # toolbar title
        self.font_h2      = ctk.CTkFont(size=16, weight="bold")   # card headings
        self.font_section = ctk.CTkFont(size=11, weight="bold")   # UPPERCASE labels
        self.font_field   = ctk.CTkFont(size=12)                  # field labels
        self.font_body    = ctk.CTkFont(size=15)                  # entries
        self.font_emph    = ctk.CTkFont(size=17, weight="bold")   # total received
        self.font_cost    = ctk.CTkFont(size=17, weight="bold")   # cost per piece
        self.font_adj     = ctk.CTkFont(size=27, weight="bold")   # adjusted tile
        self.font_big     = ctk.CTkFont(size=32, weight="bold")   # lot cost
        self.font_value   = ctk.CTkFont(size=14, weight="bold")   # breakdown values
        self.font_small   = ctk.CTkFont(size=13)                  # captions, wt/pc
        self.font_tiny    = ctk.CTkFont(size=11)                  # helper text

        # Where quick-save writes. None until the first save/open; Save As and
        # Open rebind it, so Save always updates "the file this lot lives in".
        self._current_save_path = None
        # The Fabric Lot Files folder (settings + Saved Lots). Found via the
        # per-user locator, or adopted from an older build's layout.
        self._data_dir = read_locator() or self._adopt_legacy_layout()
        self._settings = load_settings(self._data_dir)

        self.rows = []
        self.vars = {k: ctk.StringVar() for k in (
            "reference", "fabric_type", "date", "gsm", "width_in",
            "total_meters_in", "rate_per_meter", "transport_cost",
            "wastage_kg")}

        self._build_toolbar()
        self._build_body()

        # First launch (or the configured folder was deleted): ask where the
        # app's files should live, once the window is up to parent the dialog.
        if self._data_dir is None:
            self.after(600, self._ensure_data_dir)

        for var in self.vars.values():
            var.trace_add("write", self.recompute)
        for key in ("reference", "fabric_type", "date"):
            self.vars[key].trace_add("write", self._update_context)

        self._new_lot(confirm=False)

    # -------------------------------------------------------------- chrome ---
    def _set_window_icon(self):
        """Use the app's fabric-weave icon for the window / taskbar.

        CustomTkinter applies its own default icon ~200ms after launch on
        Windows, so the icon is set again after that to win the race. The .ico
        format is Windows-only; other platforms just keep their default.
        """
        if sys.platform != "win32":
            return
        icon = resource_path(os.path.join("assets", "icon.ico"))
        def apply():
            try:
                self.iconbitmap(icon)
            except Exception:  # noqa: BLE001 — missing asset: keep default
                pass
        apply()
        self.after(300, apply)

    def _set_initial_geometry(self):
        """Open near-screen-sized and centered instead of a fixed 1200×900.

        Targets the full screen minus room for the OS title bar and taskbar,
        clamped to sane bounds, and centers the window (biased slightly upward)
        so the bottom edge clears the Windows taskbar. CTk multiplies the size
        part of geometry() by its window-scaling factor — but passes +x+y
        offsets through unscaled — so size is converted to logical units while
        the position stays in physical pixels.
        """
        try:
            scaling = ctk.ScalingTracker.get_window_scaling(self)
        except Exception:  # noqa: BLE001
            scaling = 1.0
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        phys_w = max(min(sw - 140, 1880), 1000)
        phys_h = max(min(sh - 100, 1160), 760)
        # Center horizontally; center vertically within the space above the
        # taskbar (~60px allowance for taskbar + window title bar).
        x = max(0, (sw - phys_w) // 2)
        y = max(0, (sh - 60 - phys_h) // 2)
        geo = f"{int(phys_w / scaling)}x{int(phys_h / scaling)}+{x}+{y}"
        self.geometry(geo)
        # CTk's set_*_scaling() re-applies its stored window size asynchronously
        # and can clobber a geometry set during startup — re-assert once settled.
        self.after(250, lambda: self.geometry(geo))

    def _build_toolbar(self):
        bar = ctk.CTkFrame(self, fg_color=BG, corner_radius=0, height=58)
        bar.pack(fill="x")
        bar.pack_propagate(False)
        ctk.CTkLabel(bar, text="Lot Costing", font=self.font_title,
                     text_color=TEXT).pack(side="left", padx=(18, 8))
        self.context_label = ctk.CTkLabel(bar, text="", font=self.font_small,
                                          text_color=MUTED)
        self.context_label.pack(side="left")

        def btn(text, cmd, primary=False):
            if primary:
                return ctk.CTkButton(bar, text=text, command=cmd, height=36,
                                     corner_radius=6, font=self.font_body,
                                     fg_color=TEAL, hover_color=TEAL_HOVER,
                                     text_color="#ffffff")
            return ctk.CTkButton(bar, text=text, command=cmd, height=36,
                                 corner_radius=6, font=self.font_body,
                                 fg_color=SECONDARY, hover_color=SECONDARY_H,
                                 text_color=TEXT)

        btn("Export to Excel", self._export_excel, primary=True).pack(
            side="right", padx=(8, 18))
        for text, cmd in [("Save As", self._save_lot_as),
                          ("Save", self._save_lot), ("Open", self._open_lot),
                          ("New Lot", self._new_lot)]:
            btn(text, cmd).pack(side="right", padx=(8, 0))
        ctk.CTkFrame(self, fg_color=CARD_BORDER, height=1,
                     corner_radius=0).pack(fill="x")

    def _update_context(self, *_):
        parts = []
        for key in ("reference", "fabric_type"):
            v = self.vars[key].get().strip()
            if v:
                parts.append(v)
        raw = self.vars["date"].get().strip()
        if raw:
            try:
                dt = datetime.strptime(raw, "%Y-%m-%d")
                parts.append(f"{dt.day} {dt.strftime('%b %Y')}")
            except ValueError:
                parts.append(raw)
        self.context_label.configure(
            text=("·  " + "  ·  ".join(parts)) if parts else "")

    # ---------------------------------------------------------------- body ---
    def _build_body(self):
        body = ctk.CTkFrame(self, fg_color=BG, corner_radius=0)
        body.pack(fill="both", expand=True, padx=16, pady=16)
        body.grid_rowconfigure(0, weight=1)
        body.grid_columnconfigure(1, weight=1)

        left = ctk.CTkFrame(body, fg_color=CARD, border_width=1,
                            border_color=CARD_BORDER, corner_radius=8, width=300)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 16))
        left.grid_propagate(False)
        left.pack_propagate(False)
        self._build_left_card(left)

        center = ctk.CTkFrame(body, fg_color=CARD, border_width=1,
                              border_color=CARD_BORDER, corner_radius=8)
        center.grid(row=0, column=1, sticky="nsew")
        self.center_card = center
        self._build_products_card(center)

        right = ctk.CTkFrame(body, fg_color="transparent", width=296)
        right.grid(row=0, column=2, sticky="nsew", padx=(16, 0))
        right.grid_propagate(False)
        right.pack_propagate(False)
        self._build_right_column(right)

    # ------------------------------------------------------------ left card ---
    def _field(self, parent, label, key, **unit_kwargs):
        """Label + UnitEntry block bound to self.vars[key]."""
        block = ctk.CTkFrame(parent, fg_color="transparent")
        ctk.CTkLabel(block, text=label, font=self.font_field,
                     text_color=MUTED, anchor="w").pack(fill="x")
        entry = UnitEntry(block, self.vars[key],
                          font=unit_kwargs.pop("font", self.font_body),
                          **unit_kwargs)
        entry.pack(fill="x", pady=(4, 0))
        return block

    def _divider(self, parent):
        ctk.CTkFrame(parent, fg_color=HAIRLINE, height=1,
                     corner_radius=0).pack(fill="x", pady=(16, 14))

    def _build_left_card(self, card):
        inner = ctk.CTkFrame(card, fg_color="transparent")
        inner.pack(fill="both", expand=True, padx=18, pady=(14, 12))

        ctk.CTkLabel(inner, text="Fabric Lot", font=self.font_h2,
                     text_color=TEXT, anchor="w").pack(fill="x")
        ctk.CTkLabel(inner, text="What you bought from the supplier.",
                     font=self.font_field, text_color=MUTED,
                     anchor="w").pack(fill="x", pady=(0, 10))

        # --- FABRIC RECEIVED -------------------------------------------------
        ctk.CTkLabel(inner, text="F A B R I C   I N F O R M A T I O N",
                     font=self.font_section, text_color=TEAL,
                     anchor="w").pack(fill="x", pady=(0, 6))
        self._field(inner, "Fabric type", "fabric_type").pack(fill="x", pady=(0, 12))

        two = ctk.CTkFrame(inner, fg_color="transparent")
        two.pack(fill="x", pady=(0, 12))
        two.grid_columnconfigure((0, 1), weight=1, uniform="two1")
        self._field(two, "Lot reference", "reference").grid(
            row=0, column=0, sticky="ew", padx=(0, 5))
        self._field(two, "Date", "date").grid(
            row=0, column=1, sticky="ew", padx=(5, 0))

        two2 = ctk.CTkFrame(inner, fg_color="transparent")
        two2.pack(fill="x", pady=(0, 12))
        two2.grid_columnconfigure((0, 1), weight=1, uniform="two2")
        self._field(two2, "GSM", "gsm", justify="right").grid(
            row=0, column=0, sticky="ew", padx=(0, 5))
        self._field(two2, "Width", "width_in", suffix="in", justify="right").grid(
            row=0, column=1, sticky="ew", padx=(5, 0))

        # The supplier's invoice states fabric in METERS; the weight in kg is
        # derived from GSM + width (see calculations.kg_from_meters) and shown
        # in the Cost breakdown card.
        self._field(inner, "Total fabric received", "total_meters_in",
                    suffix="m", justify="right", height=44,
                    font=self.font_emph).pack(fill="x")

        self._divider(inner)

        # --- COST FROM SUPPLIER ----------------------------------------------
        ctk.CTkLabel(inner, text="C O S T   I N F O R M A T I O N",
                     font=self.font_section, text_color=TEAL,
                     anchor="w").pack(fill="x", pady=(0, 6))
        self._field(inner, "Rate per meter", "rate_per_meter",
                    prefix="Rs").pack(fill="x", pady=(0, 12))
        self._field(inner, "Transport cost", "transport_cost",
                    prefix="Rs").pack(fill="x")

        self._divider(inner)

        # --- WASTAGE ----------------------------------------------------------
        # A single lot-level field (no per-row checkbox). Synthesised into a
        # wastage Product before compute/save so the logic + file format are
        # unchanged.
        ctk.CTkLabel(inner, text="W A S T A G E", font=self.font_section,
                     text_color=WASTE_TEXT, anchor="w").pack(fill="x", pady=(0, 6))
        self._field(inner, "Fabric that made nothing", "wastage_kg", suffix="kg",
                    justify="right", fg=WASTE_BG,
                    border=WASTE_BORDER).pack(fill="x")
        ctk.CTkLabel(inner, text="Its cost is spread over the pieces that did sell.",
                     font=self.font_tiny, text_color=FAINT,
                     anchor="w", wraplength=260).pack(fill="x", pady=(6, 0))

    # -------------------------------------------------------- products card ---
    def _build_products_card(self, card):
        head = ctk.CTkFrame(card, fg_color="transparent")
        head.pack(fill="x", padx=18, pady=(14, 6))
        titles = ctk.CTkFrame(head, fg_color="transparent")
        titles.pack(side="left", fill="x", expand=True)
        ctk.CTkLabel(titles, text="Products", font=self.font_h2,
                     text_color=TEXT, anchor="w").pack(fill="x")
        ctk.CTkLabel(titles,
                     text="Type the weight and pieces — the cost per piece is "
                          "the number you sell from.",
                     font=self.font_field, text_color=MUTED, anchor="w",
                     wraplength=420, justify="left").pack(fill="x")
        ctk.CTkButton(head, text="+ Add product", height=36, corner_radius=6,
                      font=self.font_body, fg_color=TEAL_TINT,
                      hover_color=TEAL_TINT_BORDER, text_color=TEAL,
                      border_width=1, border_color=TEAL_TINT_BORDER,
                      command=lambda: (self._add_row(Product(name="")),
                                       self.recompute())).pack(side="right")

        # Column header (fixed; starts on the shared grid scheme, then
        # _sync_product_header pins each column to the rows' real pixel widths).
        header = ctk.CTkFrame(card, fg_color="transparent")
        header.pack(fill="x", anchor="w", padx=(18, 0), pady=(6, 0))
        configure_table_grid(header)
        self.products_header = header
        cols = [("PRODUCT", "w", FAINT), ("WEIGHT", "e", FAINT),
                ("PIECES", "e", FAINT), ("WT / PC", "e", FAINT),
                ("COST PER PIECE", "e", TEAL_HOVER), ("", "e", FAINT)]
        self._header_labels = []
        for i, (text, anchor, color) in enumerate(cols):
            lbl = ctk.CTkLabel(header, text=text, font=self.font_section,
                               text_color=color, anchor=anchor)
            lbl.grid(row=0, column=i, sticky="ew", padx=(0, 8) if i < 5 else 0)
            self._header_labels.append(lbl)

        # Product rows — the ONLY scrolling region in the app.
        self.table = ctk.CTkScrollableFrame(card, fg_color="transparent")
        self.table.pack(fill="both", expand=True, padx=(18, 14), pady=(4, 6))

        # Reconciliation status bar (bottom of the card).
        self.recon_bar = ctk.CTkFrame(card, fg_color=RECON_BG, corner_radius=0,
                                      height=34)
        self.recon_bar.pack(fill="x", side="bottom", padx=1, pady=(0, 1))
        self.recon_bar.pack_propagate(False)
        self.recon_label = ctk.CTkLabel(self.recon_bar, text="",
                                        font=self.font_small, text_color=MUTED,
                                        anchor="w")
        self.recon_label.pack(fill="x", padx=16)

    def _schedule_header_sync(self, *_):
        """Coalesce bursts of <Configure> events into one sync ~30 ms later."""
        if self._header_sync_pending:
            return
        self._header_sync_pending = True
        self.after(30, self._run_header_sync)

    def _run_header_sync(self):
        self._header_sync_pending = False
        try:
            self._sync_product_header()
        except Exception:  # noqa: BLE001 — window mid-teardown etc.
            pass

    def _sync_product_header(self):
        """Pin the header columns to the first row's actual pixel geometry.

        The header and the rows live in different parents (the rows sit inside a
        CTkScrollableFrame whose canvas/scrollbar eat an unpredictable number of
        pixels, DPI-scaled), so two independently-computed grids drift apart.
        Measuring the real row and copying its column widths + left origin onto
        the header is exact under any DPI, scrollbar width, or window size.
        """
        if not self.rows:
            return
        row = self.rows[0].frame
        if not row.winfo_ismapped():
            return
        # Match the header's left edge to the rows' left edge.
        dx = row.winfo_rootx() - self.center_card.winfo_rootx()
        if dx > 0:
            self.products_header.pack_configure(padx=(dx, 0))
        # Copy each column's real width.
        for i in range(6):
            bbox = row.grid_bbox(column=i, row=0)
            if bbox and bbox[2] > 0:
                self.products_header.grid_columnconfigure(
                    i, minsize=bbox[2], weight=0)

    # ---------------------------------------------------------- right column ---
    def _build_right_column(self, container):
        # --- Cost breakdown card ---
        card = ctk.CTkFrame(container, fg_color=CARD, border_width=1,
                            border_color=CARD_BORDER, corner_radius=8)
        card.pack(fill="x")
        inner = ctk.CTkFrame(card, fg_color="transparent")
        inner.pack(fill="x", padx=18, pady=(14, 14))
        ctk.CTkLabel(inner, text="Cost breakdown", font=self.font_h2,
                     text_color=TEXT, anchor="w").pack(fill="x")
        ctk.CTkLabel(inner, text="Worked out step by step.", font=self.font_field,
                     text_color=MUTED, anchor="w").pack(fill="x", pady=(0, 6))

        self.breakdown = {}
        rows = [("meters_per_kg", "Meters per kg"),
                ("total_weight", "Total weight"),
                ("fabric_cost", "Fabric cost"),
                ("total_cost", "Total cost"),
                ("base_cost_per_kg", "Base cost / kg"),
                ("wastage_cost_per_kg", "Wastage cost / kg")]
        for key, label in rows:
            ctk.CTkFrame(inner, fg_color=HAIRLINE, height=1,
                         corner_radius=0).pack(fill="x")
            line = ctk.CTkFrame(inner, fg_color="transparent")
            line.pack(fill="x", pady=5)
            ctk.CTkLabel(line, text=label, font=self.font_small,
                         text_color=MUTED, anchor="w").pack(side="left")
            val = ctk.CTkLabel(line, text="—", font=self.font_value,
                               text_color=TEXT, anchor="e")
            val.pack(side="right")
            self.breakdown[key] = val

        # Adjusted cost per kg — the highlighted tile.
        tile = ctk.CTkFrame(inner, fg_color=TEAL_TINT, corner_radius=8)
        tile.pack(fill="x", pady=(10, 0))
        ctk.CTkLabel(tile, text="A D J U S T E D   C O S T   P E R   K G",
                     font=self.font_section, text_color=TEAL_HOVER,
                     anchor="w").pack(fill="x", padx=14, pady=(12, 0))
        self.adjusted_label = ctk.CTkLabel(tile, text="—", font=self.font_adj,
                                           text_color=TEAL_HOVER, anchor="w")
        self.adjusted_label.pack(fill="x", padx=14)
        ctk.CTkLabel(tile, text="Every piece is costed at this rate.",
                     font=self.font_tiny, text_color=TEAL_HOVER,
                     anchor="w").pack(fill="x", padx=14, pady=(0, 12))

        # --- Lot cost summary card ---
        card2 = ctk.CTkFrame(container, fg_color=CARD, border_width=1,
                             border_color=CARD_BORDER, corner_radius=8)
        card2.pack(fill="x", pady=(16, 0))
        inner2 = ctk.CTkFrame(card2, fg_color="transparent")
        inner2.pack(fill="x", padx=18, pady=(14, 14))
        ctk.CTkLabel(inner2, text="This lot cost you", font=self.font_small,
                     text_color=MUTED, anchor="w").pack(fill="x")
        self.lot_cost_label = ctk.CTkLabel(inner2, text="—", font=self.font_big,
                                           text_color=TEXT, anchor="w")
        self.lot_cost_label.pack(fill="x")
        self.spread_label = ctk.CTkLabel(inner2, text="", font=self.font_small,
                                         text_color=MUTED, anchor="w",
                                         wraplength=250, justify="left")
        self.spread_label.pack(fill="x", pady=(0, 8))

        self.summary = {}
        for key, label in [("total_pieces", "Total pieces"),
                           ("fabric_used", "Fabric used")]:
            ctk.CTkFrame(inner2, fg_color=HAIRLINE, height=1,
                         corner_radius=0).pack(fill="x")
            line = ctk.CTkFrame(inner2, fg_color="transparent")
            line.pack(fill="x", pady=5)
            ctk.CTkLabel(line, text=label, font=self.font_small,
                         text_color=MUTED, anchor="w").pack(side="left")
            val = ctk.CTkLabel(line, text="—", font=self.font_value,
                               text_color=TEXT, anchor="e")
            val.pack(side="right")
            self.summary[key] = val

        self.rupee_label = ctk.CTkLabel(
            inner2, text="✓  Wastage cost is included in each piece's cost.",
            font=self.font_field, text_color=GREEN, anchor="w",
            wraplength=250, justify="left")
        self.rupee_label.pack(fill="x", pady=(6, 0))

    # ------------------------------------------------------------- products ---
    def _add_row(self, product=None):
        row = ProductRow(self.table, self, product or Product(name=""))
        row.frame.pack(fill="x")
        self.rows.append(row)

    def _delete_row(self, row):
        row.destroy()
        self.rows.remove(row)
        self.recompute()

    def _products_for_compute(self):
        """Table rows + the wastage Product synthesised from the left-card field.

        This keeps calculations.py and the JSON format untouched: wastage still
        reaches compute()/save as an is_wastage row, exactly as before.
        """
        products = [row.to_product() for row in self.rows]
        wastage = self.vars["wastage_kg"].get().strip()
        if wastage != "":
            products.append(Product(name="Wastage", weight_kg=wastage,
                                    pieces=None, is_wastage=True))
        return products

    # ------------------------------------------------------------- recompute ---
    def _current_lot(self):
        v = self.vars
        return LotInfo(
            reference=v["reference"].get().strip(),
            fabric_type=v["fabric_type"].get().strip(),
            date=v["date"].get().strip(),
            gsm=v["gsm"].get(),
            width_in=v["width_in"].get(),
            # The operator enters meters; the weight the whole calculation
            # runs on (and the JSON stores) is derived here, so compute(),
            # reconciliation, and old save files all keep working on kg.
            total_kg=calculations.kg_from_meters(
                v["gsm"].get(), v["width_in"].get(),
                v["total_meters_in"].get()),
            rate_per_meter=v["rate_per_meter"].get(),
            transport_cost=v["transport_cost"].get(),
        )

    def recompute(self, *_):
        """Recompute everything and repaint every derived field on screen."""
        lot = self._current_lot()
        products = self._products_for_compute()
        results = calculations.compute(lot, products)

        # Product rows (zip stops before the synthesised wastage result).
        for row, pr in zip(self.rows, results.product_results):
            row.paint(pr)

        # Cost breakdown.
        self.breakdown["meters_per_kg"].configure(
            text="—" if results.meters_per_kg is None
            else f"{fmt(results.meters_per_kg)} m")
        derived_kg = calculations.parse_number(lot.total_kg)
        self.breakdown["total_weight"].configure(
            text="—" if derived_kg is None else f"{fmt(derived_kg)} kg")
        for key in ("fabric_cost", "total_cost", "base_cost_per_kg"):
            self.breakdown[key].configure(
                text=fmt(getattr(results, key), money=True))
        w = results.wastage_cost_per_kg
        if w is not None and w > 0:
            self.breakdown["wastage_cost_per_kg"].configure(
                text=f"+ {fmt(w, money=True)}", text_color=WASTE_TEXT)
        else:
            self.breakdown["wastage_cost_per_kg"].configure(
                text=fmt(w, money=True), text_color=TEXT)
        self.adjusted_label.configure(
            text=fmt(results.adjusted_cost_per_kg, money=True))

        # Lot cost summary ("This lot cost you" framing — no Receipt/Profit).
        self.lot_cost_label.configure(text=fmt(results.total_cost, money=True))
        total_kg = calculations.parse_number(lot.total_kg)
        if results.total_cost is not None and total_kg:
            self.spread_label.configure(
                text=f"spread across {fmt(results.total_pieces, decimals=0)} "
                     f"pieces from {kgfmt(total_kg)} kg.")
            self.rupee_label.configure(
                text="✓  Wastage cost is included in each piece's cost.")
        else:
            self.spread_label.configure(text="")
            self.rupee_label.configure(text="")
        self.summary["total_pieces"].configure(
            text=fmt(results.total_pieces, decimals=0))
        self.summary["fabric_used"].configure(
            text=f"{fmt(results.total_weight_produced)} kg")

        self._refresh_recon(results, total_kg)
        self._schedule_header_sync()

    def _refresh_recon(self, results, total_kg):
        """Reconciliation status bar: green when it adds up, orange on mismatch."""
        if total_kg is None:
            self.recon_label.configure(
                text="Enter the lot and product weights to check the fabric "
                     "balance.", text_color=MUTED)
            self.recon_bar.configure(fg_color=RECON_BG)
        elif results.recon_mismatch:
            self.recon_label.configure(text=results.recon_message,
                                       text_color=WASTE_TEXT)
            self.recon_bar.configure(fg_color=WASTE_BG)
        else:
            weights = [calculations.parse_number(r.weight_var.get())
                       for r in self.rows]
            weights = [w for w in weights if w is not None]
            wastage = calculations.parse_number(
                self.vars["wastage_kg"].get()) or 0.0
            parts = " + ".join(kgfmt(w) for w in weights) if weights else "0"
            self.recon_label.configure(
                text=f"✓  Fabric adds up — {parts} kg products + "
                     f"{kgfmt(wastage)} kg wastage = {kgfmt(total_kg)} kg "
                     f"received.", text_color=GREEN)
            self.recon_bar.configure(fg_color=RECON_BG)

    # ------------------------------------------------------------ toolbar ops ---
    def _clear_rows(self):
        for row in self.rows:
            row.destroy()
        self.rows.clear()

    def _new_lot(self, confirm=True):
        if confirm and not messagebox.askyesno(
                "New Lot", "Clear all fields and start a new lot?"):
            return
        self._current_save_path = None  # a fresh lot isn't bound to a file yet
        for key, var in self.vars.items():
            var.set(date.today().isoformat() if key == "date" else "")
        self._clear_rows()
        self._add_row(Product(name=""))
        self._add_row(Product(name=""))
        self.recompute()

    def _load_into_ui(self, lot, products):
        """Populate the UI from a loaded lot.

        Backward-compat split: older files store wastage as an is_wastage row.
        The FIRST flagged row's weight fills the Wastage field (mirroring
        calculations._wastage_weight's first-match rule); any extra flagged rows
        are defensively demoted to normal product rows.
        """
        self.vars["reference"].set(lot.reference or "")
        self.vars["fabric_type"].set(lot.fabric_type or "")
        self.vars["date"].set(lot.date or "")
        for key in ("gsm", "width_in", "rate_per_meter", "transport_cost"):
            val = getattr(lot, key)
            self.vars[key].set("" if val is None else str(val))

        # Files store the weight (kg, for compatibility); the UI shows meters.
        # Derive meters back via the same formula compute() uses; blank if the
        # gsm/width needed for the conversion are missing.
        meters = ""
        total_kg = calculations.parse_number(lot.total_kg)
        kg_per_meter = calculations.kg_from_meters(lot.gsm, lot.width_in, 1)
        if total_kg is not None and kg_per_meter:
            meters = f"{total_kg / kg_per_meter:g}"
        self.vars["total_meters_in"].set(meters)

        wastage_val = ""
        table_rows = []
        for p in products:
            if p.is_wastage and wastage_val == "":
                wastage_val = "" if p.weight_kg in (None, "") else str(p.weight_kg)
            else:
                table_rows.append(Product(name=p.name, weight_kg=p.weight_kg,
                                          pieces=p.pieces, is_wastage=False))
        self.vars["wastage_kg"].set(wastage_val)

        self._clear_rows()
        for p in table_rows:
            self._add_row(p)
        if not table_rows:
            self._add_row(Product(name=""))
        self.recompute()

    # ------------------------------------------------------ app data folder ---
    def _adopt_legacy_layout(self):
        """Adopt the folder a previous build configured, without re-asking.

        Older builds kept settings.json next to the exe with a "lots_dir" key
        pointing at the Fabric Lot Files folder. If that's present and valid,
        take it over: write the locator, move its settings into the folder,
        and tidy stray files into the new structure.
        """
        legacy = storage.app_dir() / "settings.json"
        try:
            import json
            with open(legacy, "r", encoding="utf-8") as f:
                old = json.load(f)
            d = Path(old.get("lots_dir", ""))
            if not d.is_dir():
                return None
            write_locator(d)
            merged = load_settings(d)
            for key, value in old.items():
                merged.setdefault(key, value)
            merged.pop("lots_dir", None)  # the locator owns this now
            save_settings(d, merged)
            legacy.unlink()
            self._organize_data_dir(d)
            return d
        except Exception:  # noqa: BLE001 — no/invalid legacy file
            return None

    def _ensure_data_dir(self) -> Path:
        """Return the Fabric Lot Files folder, asking where to put it on first use.

        The chosen location gets a "Fabric Lot Files" folder holding
        settings.json and a "Saved Lots" subfolder; its path is recorded in the
        per-user locator. Cancelling the picker falls back to Documents so the
        app always has somewhere to write. Nothing is stored next to the exe.
        """
        if self._data_dir and self._data_dir.is_dir():
            return self._data_dir
        docs = Path.home() / "Documents"
        default_base = docs if docs.exists() else Path.home()
        messagebox.showinfo(
            "App files folder",
            "Choose where the app should keep its files — saved lots, Excel "
            "reports, and settings.\n\n"
            'A folder called "Fabric Lot Files" will be created there.\n'
            "(If you press Cancel, it will be created in Documents.)")
        picked = filedialog.askdirectory(
            title="Choose where to keep the app's files",
            initialdir=str(default_base))
        base = Path(picked) if picked else default_base
        # Don't nest a second level if the user selected the folder itself.
        d = base if base.name == "Fabric Lot Files" else base / "Fabric Lot Files"
        try:
            d.mkdir(parents=True, exist_ok=True)
        except Exception:  # noqa: BLE001 — unwritable pick: fall back safely
            d = default_base / "Fabric Lot Files"
            d.mkdir(parents=True, exist_ok=True)
        self._data_dir = d
        write_locator(d)
        save_settings(d, self._settings)
        self._organize_data_dir(d)
        return d

    def _saved_lots_dir(self) -> Path:
        """The Saved Lots subfolder inside the Fabric Lot Files folder."""
        d = self._ensure_data_dir() / "Saved Lots"
        d.mkdir(parents=True, exist_ok=True)
        return d

    def _excel_reports_dir(self) -> Path:
        """The Excel Reports subfolder inside the Fabric Lot Files folder."""
        d = self._ensure_data_dir() / "Excel Reports"
        d.mkdir(parents=True, exist_ok=True)
        return d

    def _organize_data_dir(self, data_dir: Path):
        """Best-effort migration of older layouts into the folder structure.

        Moves lot JSONs from the old lots/ folder next to the exe, and any lot
        JSONs sitting in the Fabric Lot Files root (the previous build saved
        them there), into the Saved Lots subfolder.
        """
        import shutil
        saved = data_dir / "Saved Lots"
        try:
            saved.mkdir(parents=True, exist_ok=True)
            (data_dir / "Excel Reports").mkdir(parents=True, exist_ok=True)
        except Exception:  # noqa: BLE001
            return
        old_lots = storage.app_dir() / "lots"
        try:
            if old_lots.is_dir() and old_lots.resolve() != saved.resolve():
                for f in old_lots.glob("*.json"):
                    if not (saved / f.name).exists():
                        shutil.move(str(f), str(saved / f.name))
        except Exception:  # noqa: BLE001 — old files stay put
            pass
        try:
            for f in data_dir.glob("*.json"):
                if f.name != "settings.json" and not (saved / f.name).exists():
                    shutil.move(str(f), str(saved / f.name))
        except Exception:  # noqa: BLE001
            pass

    def _save_lot(self):
        """Quick save — no dialog.

        Writes to the lot's bound file (set by a previous Save, Save As, or
        Open); a never-saved lot defaults to {reference}_{date}.json in the
        Fabric Lot Files folder.
        """
        lot = self._current_lot()
        products = self._products_for_compute()  # incl. synthesised wastage row
        path = self._current_save_path or (
            self._saved_lots_dir() / storage.default_filename(lot, "json"))
        self._write_lot(lot, products, path)

    def _save_lot_as(self):
        """Pick a filename/destination; quick saves then update that file."""
        lot = self._current_lot()
        products = self._products_for_compute()
        if self._current_save_path:
            initialdir = str(self._current_save_path.parent)
            initialfile = self._current_save_path.name
        else:
            initialdir = str(self._saved_lots_dir())
            initialfile = storage.default_filename(lot, "json")
        path = filedialog.asksaveasfilename(
            title="Save Lot As", defaultextension=".json",
            initialdir=initialdir, initialfile=initialfile,
            filetypes=[("Lot files", "*.json")],
        )
        if not path:
            return
        self._write_lot(lot, products, Path(path))

    def _write_lot(self, lot, products, path):
        try:
            storage.save_lot(lot, products, str(path))
            self._current_save_path = Path(path)
            messagebox.showinfo("Save Lot", f"Saved {Path(path).name}")
        except Exception as exc:  # noqa: BLE001 — surface any IO error friendly
            messagebox.showerror("Save Lot", f"Could not save the lot:\n{exc}")

    def _open_lot(self):
        path = filedialog.askopenfilename(
            title="Open Lot", initialdir=str(self._saved_lots_dir()),
            filetypes=[("Lot files", "*.json"), ("All files", "*.*")],
        )
        if not path:
            return
        try:
            lot, products = storage.load_lot(path)
            self._load_into_ui(lot, products)
            # Quick saves now update the file that was opened.
            self._current_save_path = Path(path)
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("Open Lot", f"Could not open the lot:\n{exc}")

    def _export_excel(self):
        lot = self._current_lot()
        products = self._products_for_compute()  # incl. synthesised wastage row
        results = calculations.compute(lot, products)
        # Open in the folder the user last exported to; fall back to the
        # Excel Reports subfolder of the app's own data folder.
        remembered = self._settings.get("export_dir")
        if remembered and Path(remembered).is_dir():
            initialdir = remembered
        else:
            initialdir = str(self._excel_reports_dir())
        path = filedialog.asksaveasfilename(
            title="Export to Excel", defaultextension=".xlsx",
            initialdir=initialdir,
            initialfile=report_filename(lot),
            filetypes=[("Excel files", "*.xlsx")],
        )
        if not path:
            return
        # Remember wherever the user actually saved for next time.
        self._settings["export_dir"] = str(Path(path).parent)
        save_settings(self._ensure_data_dir(), self._settings)
        try:
            excel_export.export(results, lot, path)
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("Export to Excel", f"Could not export:\n{exc}")
            return
        # An exported lot is worth keeping: quick-save its JSON too (to the
        # lot's bound file, or its default in Saved Lots), so the report and
        # the reloadable lot can never drift apart.
        note = ""
        try:
            lot_path = self._current_save_path or (
                self._saved_lots_dir() / storage.default_filename(lot, "json"))
            storage.save_lot(lot, products, str(lot_path))
            self._current_save_path = Path(lot_path)
            note = f"\nLot also saved ({Path(lot_path).name})."
        except Exception as exc:  # noqa: BLE001 — report, but export succeeded
            note = f"\nNote: the lot itself could not be saved:\n{exc}"
        messagebox.showinfo("Export to Excel",
                            f"Excel file created successfully.{note}")


def _fit_scaling_to_screen():
    """Cap CTk's DPI scaling so the full content fits the screen.

    On Windows at 125/150 % display scaling, CustomTkinter multiplies every
    dimension by the DPI factor, so the ~900px-tall layout overflows a 1080p
    monitor and forces the user to fullscreen. This computes a scaling that
    fits and applies it BEFORE any CTk window exists — calling the set_*_scaling
    functions on a live window makes CTk's ScalingTracker asynchronously revert
    the window geometry, so the decision has to happen up front.
    """
    import tkinter
    try:
        # Make the process DPI-aware first so the probe reads true pixels.
        ctk.ScalingTracker.activate_high_dpi_awareness()
    except Exception:  # noqa: BLE001
        pass
    try:
        probe = tkinter.Tk()
        probe.withdraw()
        sw, sh = probe.winfo_screenwidth(), probe.winfo_screenheight()
        if probe.tk.call("tk", "windowingsystem") == "win32":
            dpi = probe.winfo_fpixels("1i") / 96.0  # 96 dpi == 100 % scaling
        else:
            dpi = 1.0  # macOS/X11 handle high-DPI natively
        probe.destroy()
    except Exception:  # noqa: BLE001 — no display etc.: keep defaults
        return
    fit = min(sw / 1240, (sh - 90) / 900, dpi)
    if fit < dpi - 0.01:
        ctk.set_widget_scaling(fit)
        ctk.set_window_scaling(fit)


def main():
    _fit_scaling_to_screen()
    app = CostingApp()
    app.mainloop()


if __name__ == "__main__":
    main()
