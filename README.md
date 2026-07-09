# Textile Lot Costing

A standalone Windows desktop app that replaces the paper costing sheet used for a
bedsheet / fabric business. You buy a **lot** of raw fabric by weight, convert it
into finished products (bedsheets, pillows, …), and this app tracks the cost of
the lot against the revenue those products recover — including spreading the cost
of unsellable **wastage** across everything that did sell.

All amounts are in **PKR (Rs)**.

- Live calculations as you type — nothing to recompute by hand.
- Products table you can add to / delete from, with one row marked as wastage.
- A reconciliation warning that catches data-entry mistakes (fabric received
  should equal the fabric turned into products + waste).
- **Save / Open** past lots (local JSON files).
- **Export to Excel** — a formatted `.xlsx` mirroring the paper sheet.
- Ships as a single double-click `.exe` — no Python, no terminal for the end user.

---

## Files

| File | Purpose |
|------|---------|
| `main.py` | The app window (tkinter GUI) and entry point. |
| `calculations.py` | All the business formulas (pure, tested). |
| `excel_export.py` | Builds the formatted Excel workbook. |
| `storage.py` | Save / load a lot to / from JSON. |
| `test_calculations.py` | Tests for the formulas. |
| `requirements.txt` | Python dependencies. |

---

## Running during development

You need Python 3.9+ installed.

### macOS / Linux

```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
python main.py
```

### Windows

```bat
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
python main.py
```

Run the tests any time with:

```bash
python test_calculations.py
```

---

## Building the Windows `.exe`

**Do this on a Windows machine** — PyInstaller cannot cross-compile a Windows
executable from macOS/Linux. Open a Command Prompt in the project folder and copy
-paste the block below:

```bat
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
pyinstaller --onefile --windowed --noconsole --name TextileCosting --collect-all openpyxl main.py
```

What the flags do:

- `--onefile` — bundle everything into a single `.exe`.
- `--windowed --noconsole` — GUI app; **no black terminal window ever appears**.
- `--name TextileCosting` — names the output `TextileCosting.exe`.
- `--collect-all openpyxl` — force-bundles openpyxl (without this, the Excel
  export can fail in the packaged app with a missing-module error).

### Where the `.exe` ends up

After the build finishes, your app is at:

```
dist\TextileCosting.exe
```

Double-click it to run. It creates a `lots\` folder next to itself the first time
you save a lot. You can copy `TextileCosting.exe` anywhere (desktop, USB drive) —
it's fully self-contained.

> **Test the built exe, not just the script:** double-click `dist\TextileCosting.exe`,
> enter a lot, and use **Export to Excel** to confirm openpyxl was bundled
> correctly.

---

## How the numbers work (business rules)

Entered per lot: **GSM**, **Width (inches)**, **Total KG Received**,
**Rate per Meter**, **Transport Cost**.

```
Meters per KG       = 1000 / (Width_inches × 0.0254 × GSM)
Total Meters        = Meters per KG × Total KG Received
Fabric Cost         = Total Meters × Rate per Meter
Total Cost          = Fabric Cost + Transport Cost
Base Cost per KG    = Total Cost / Total KG Received
```

**Wastage spreading** — wastage fabric produces nothing, but you still paid for
it. Because this is a cost tool, that cost is spread over the **sellable** weight
(the only fabric that can carry it) so every rupee spent is allocated back onto a
finished piece — the per-piece costs add up to the total cost exactly:

```
Sellable Weight      = Total KG Received − Wastage Weight
Wastage Cost per KG  = (Wastage Weight × Base Cost per KG) / Sellable Weight
Adjusted Cost per KG = Base Cost per KG + Wastage Cost per KG   ← each piece is costed at THIS
```

Per product row:

```
Weight per Piece = Weight / Pieces        (shows "N/A" if Pieces = 0)
Cost per Piece   = Weight per Piece × Adjusted Cost per KG
Revenue          = Pieces × Cost per Piece
```

**Reconciliation** — the app continuously compares **Total KG Received** with the
**sum of all product weights (including wastage)**. They should match (fabric in =
products + waste). If they differ by more than 0.01 kg, an orange warning appears,
but it never blocks calculation or export.

**Summary** — Total Payment = Total Cost, Total Receipt = sum of revenues, and
**Profit = Receipt − Payment** (green if positive, red if negative).
