
import re
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
from pathlib import Path
from typing import List, Dict, Optional, Tuple
from openpyxl import load_workbook, Workbook

# === Configuration ===
EXCEL_PATH = Path("molds.xlsx")
SHEET_NAME = "Sheet1"  # inventory lives here
LOG_SHEET = "Log"

# === Helpers ===
def normalize(s: str) -> str:
    if not s:
        return ""
    s = s.upper()
    s = re.sub(r'[\s\-._/\\]', '', s)  # remove spaces, dashes, dots, underscores, slashes
    return s

def find_header_row(ws) -> Optional[int]:
    # Find a row that contains PART#, QTY, and BOX (variants) — return 1-based row index
    max_scan = min(100, ws.max_row)
    targets = {"PART#", "QTY", "BOX", "BOX#", "BOX #", "BOX NUMBER"}
    for i in range(1, max_scan+1):
        vals = [(ws.cell(row=i, column=j).value or "") for j in range(1, ws.max_column+1)]
        up = [str(v).strip().upper() for v in vals]
        if ("PART#" in up or "NAME" in up) and ("QTY" in up or "QT" in up) and any(t in up for t in targets):
            return i
    return None

def ensure_workbook() -> Tuple[Workbook, int, Dict[str,int]]:
    if not EXCEL_PATH.exists():
        # Create a minimal workbook if missing
        wb = Workbook()
        ws = wb.active
        ws.title = SHEET_NAME
        ws.append(["STOCK ROOM BOOTS", "", datetime.now().date()])
        ws.append([])
        ws.append(["PART#", "QTY", "BOX#"])
        wb.save(EXCEL_PATH)
    wb = load_workbook(EXCEL_PATH)
    if SHEET_NAME not in wb.sheetnames:
        wb.create_sheet(SHEET_NAME)
        wb[SHEET_NAME].append(["PART#", "QTY", "BOX#"])
        wb.save(EXCEL_PATH)
    ws = wb[SHEET_NAME]
    header_row = find_header_row(ws) or 3  # default to row 3 like the uploaded file
    # Map headers to columns
    header_vals = [str(ws.cell(row=header_row, column=j).value or "").strip().upper() for j in range(1, ws.max_column+1)]
    col_map = {}
    for j, name in enumerate(header_vals, start=1):
        if name in ("PART#", "NAME"):
            col_map["PART#"] = j
        elif name in ("QTY", "QT"):
            col_map["QTY"] = j
        elif name in ("BOX#", "BOX", "BOX #", "BOX NUMBER"):
            col_map["BOX#"] = j
    # Ensure required columns exist
    for need in ("PART#", "QTY", "BOX#"):
        if need not in col_map:
            # append missing column name at end
            new_col = ws.max_column + 1
            ws.cell(row=header_row, column=new_col, value=need)
            col_map[need] = new_col
    # Ensure Log sheet
    if LOG_SHEET not in wb.sheetnames:
        log = wb.create_sheet(LOG_SHEET)
        log.append(["Time", "User", "Action", "Part#", "QtyChange", "FromBox", "ToBox", "Notes"])
    return wb, header_row, col_map

def load_inventory(ws, header_row, col_map) -> List[Dict]:
    rows = []
    for r in range(header_row+1, ws.max_row+1):
        part = ws.cell(row=r, column=col_map["PART#"]).value
        qty  = ws.cell(row=r, column=col_map["QTY"]).value
        box  = ws.cell(row=r, column=col_map["BOX#"]).value
        if part is None and qty is None and box is None:
            continue
        # Skip any repeated header lines
        if str(part).strip().upper() in ("PART#", "NAME"):
            continue
        try:
            qty = int(qty) if qty is not None and str(qty).strip() != "" else 0
        except:
            qty = 0
        rows.append({
            "row": r,
            "PART#": None if part is None else str(part).strip(),
            "QTY": qty,
            "BOX#": box
        })
    return rows

def write_log(wb, user: str, action: str, part: str, qty_delta=0, from_box="", to_box="", notes=""):
    log = wb[LOG_SHEET]
    log.append([datetime.now().isoformat(timespec="seconds"), user, action, part, qty_delta, from_box, to_box, notes])

def normalize_all(user: str="") -> int:
    wb, header_row, col_map = ensure_workbook()
    ws = wb[SHEET_NAME]
    changed = 0
    for r in range(header_row+1, ws.max_row+1):
        part = ws.cell(row=r, column=col_map["PART#"]).value
        if part:
            canon = normalize(str(part))
            if str(part) != canon:
                ws.cell(row=r, column=col_map["PART#"]).value = canon
                changed += 1
    if changed:
        log = wb[LOG_SHEET]
        log.append([datetime.now().isoformat(timespec="seconds"), user, "NORMALIZE_ALL", "", 0, "", "", f"{changed} items"])
        wb.save(EXCEL_PATH)
    return changed


# === Inventory Ops (Add/Remove) ===
def add_mode(part_input: str, qty: int, box: Optional[int], user: str) -> str:
    wb, header_row, col_map = ensure_workbook()
    ws = wb[SHEET_NAME]
    items = load_inventory(ws, header_row, col_map)

    target_norm = normalize(part_input)
    match = next((it for it in items if normalize(it["PART#"]) == target_norm), None)
    canon = normalize(part_input)  # canonical Part# we will STORE

    if match:
        # increment and optionally move box; also overwrite stored Part# to canonical
        new_qty = match["QTY"] + qty
        ws.cell(row=match["row"], column=col_map["QTY"]).value = new_qty
        ws.cell(row=match["row"], column=col_map["PART#"]).value = canon
        if box is not None:
            ws.cell(row=match["row"], column=col_map["BOX#"]).value = box
        write_log(wb, user, "ADD/INCR", canon, qty_delta=qty, to_box=box or match["BOX#"])
        wb.save(EXCEL_PATH)
        return f"Updated {canon}: {match['QTY']} → {new_qty}  (Box: {box or match['BOX#']})"
    else:
        # new row with canonical Part#
        last_row = ws.max_row + 1
        ws.cell(row=last_row, column=col_map["PART#"]).value = canon
        ws.cell(row=last_row, column=col_map["QTY"]).value = qty
        ws.cell(row=last_row, column=col_map["BOX#"]).value = box
        write_log(wb, user, "ADD_NEW", canon, qty_delta=qty, to_box=box, notes="created")
        wb.save(EXCEL_PATH)
        return f"Added {canon}  (qty {qty}, box {box})."


def remove_mode(part_input: str, qty: int, user: str) -> str:
    wb, header_row, col_map = ensure_workbook()
    ws = wb[SHEET_NAME]
    items = load_inventory(ws, header_row, col_map)
    target_norm = normalize(part_input)
    match = next((it for it in items if normalize(it["PART#"]) == target_norm), None)

    if not match:
        return "Code not found."
    if qty <= 0:
        return "Quantity must be positive."
    if match["QTY"] - qty < 0:
        return f"Cannot remove {qty}; only {match['QTY']} available."
    new_qty = match["QTY"] - qty
    ws.cell(row=match["row"], column=col_map["QTY"]).value = new_qty
    write_log(wb, user, "REMOVE/DECR", match["PART#"], qty_delta=-qty, from_box=match["BOX#"])
    wb.save(EXCEL_PATH)
    return f"Updated {match['PART#']}: {match['QTY']} → {new_qty}"

# === GUI ===
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Boots Inventory (Add / Remove)")
        self.geometry("920x560")

        # State
        self.user_var = tk.StringVar()
        self.mode_var = tk.StringVar(value="ADD")
        self.search_var = tk.StringVar()
        self.code_var = tk.StringVar()
        self.qty_var = tk.StringVar(value="1")
        self.box_var = tk.StringVar()  # only used in ADD mode
        self.msg_var = tk.StringVar(value="Ready.")

        self._build_top()
        self._build_table()
        self._bind_events()
        self.refresh_table()

    def _build_top(self):
        top = ttk.Frame(self, padding=10)
        top.pack(fill="x")

        # Row 0
        ttk.Label(top, text="User:").grid(row=0, column=0, sticky="e", padx=(0,6))
        ttk.Entry(top, textvariable=self.user_var, width=12).grid(row=0, column=1, sticky="w", padx=(0,14))

        ttk.Radiobutton(top, text="Add mode", value="ADD", variable=self.mode_var).grid(row=0, column=2, sticky="w")
        ttk.Radiobutton(top, text="Remove mode", value="REMOVE", variable=self.mode_var).grid(row=0, column=3, sticky="w", padx=(14,0))

        ttk.Label(top, text="Search:").grid(row=0, column=4, sticky="e", padx=(14,6))
        ttk.Entry(top, textvariable=self.search_var, width=28).grid(row=0, column=5, sticky="w")

        ttk.Button(top, text="Refresh", command=self.refresh_table).grid(row=0, column=6, sticky="w", padx=(10,0))

        # Row 1
        ttk.Label(top, text="Part#:").grid(row=1, column=0, sticky="e", padx=(0,6), pady=(10,0))
        ttk.Entry(top, textvariable=self.code_var, width=28).grid(row=1, column=1, columnspan=2, sticky="w", pady=(10,0))

        ttk.Label(top, text="Qty:").grid(row=1, column=3, sticky="e", padx=(10,6), pady=(10,0))
        ttk.Entry(top, textvariable=self.qty_var, width=8).grid(row=1, column=4, sticky="w", pady=(10,0))

        ttk.Label(top, text="Box# (Add mode):").grid(row=1, column=5, sticky="e", padx=(10,6), pady=(10,0))
        ttk.Entry(top, textvariable=self.box_var, width=10).grid(row=1, column=6, sticky="w", pady=(10,0))

        # Row 2 - Buttons
        btns = ttk.Frame(top)
        btns.grid(row=2, column=0, columnspan=7, sticky="w", pady=(12,0))
        ttk.Button(btns, text="Apply", command=self.on_apply).grid(row=0, column=0, padx=(0,10))
        ttk.Button(btns, text="Normalize All", command=self.on_normalize_all).grid(row=0, column=1, padx=(0,10))
        ttk.Button(btns, text="Open in Excel", command=self.open_excel).grid(row=0, column=2, padx=(0,10))



        # Status
        status = ttk.Frame(self, padding=(10,0,10,10))
        status.pack(fill="x")
        ttk.Label(status, textvariable=self.msg_var).pack(side="left")

        for c in range(7):
            top.grid_columnconfigure(c, weight=1)

    def _build_table(self):
        mid = ttk.Frame(self, padding=(10,0,10,10))
        mid.pack(fill="both", expand=True)

        cols = ("PART#", "QTY", "BOX#")
        self.tree = ttk.Treeview(mid, columns=cols, show="headings", selectmode="browse")
        for c in cols:
            self.tree.heading(c, text=c)
            anchor = "w" if c == "PART#" else "center"
            width = 360 if c == "PART#" else 90
            self.tree.column(c, anchor=anchor, width=width)
        self.tree.pack(side="left", fill="both", expand=True)

        vsb = ttk.Scrollbar(mid, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")


    def _bind_events(self):
        self.search_var.trace_add("write", lambda *args: self.refresh_table())
        self.tree.bind("<<TreeviewSelect>>", self.on_select_row)

    def on_select_row(self, _evt=None):
        sel = self.tree.selection()
        if not sel:
            return
        vals = self.tree.item(sel[0], "values")
        if not vals:
            return
        part, qty, box = vals
        self.code_var.set(part)
        self.qty_var.set("1")  # default change
        self.box_var.set(str(box))

    def _parse_int(self, s: str, allow_none=False) -> Optional[int]:
        s = (s or "").strip()
        if s == "" and allow_none:
            return None
        try:
            return int(s)
        except:
            return None

    def refresh_table(self):
        # Clear rows
        for iid in self.tree.get_children():
            self.tree.delete(iid)

        wb, header_row, col_map = ensure_workbook()
        ws = wb[SHEET_NAME]
        items = load_inventory(ws, header_row, col_map)

        # Filter by search (normalized contains)
        term = normalize(self.search_var.get())
        if term:
            items = [it for it in items if term in normalize(it["PART#"] or "")]

        # Sort by part#, then box
        items.sort(key=lambda it: (str(it["PART#"] or ""), it["BOX#"] or 0))
        for it in items:
            self.tree.insert("", "end", values=(it["PART#"], it["QTY"], it["BOX#"]))

        self.msg_var.set(f"{len(items)} items shown.")

    def on_apply(self):
        user = self.user_var.get().strip()
        code = self.code_var.get().strip()
        qty = self._parse_int(self.qty_var.get())
        if not code or qty is None or qty <= 0:
            messagebox.showerror("Input error", "Enter a Part# and a positive Qty.")
            return

        try:
            if self.mode_var.get() == "ADD":
                box = self._parse_int(self.box_var.get(), allow_none=True)
                msg = add_mode(code, qty, box, user)
            else:
                msg = remove_mode(code, qty, user)

            if msg == "Code not found.":
                messagebox.showwarning("Not found", msg)
            elif msg.startswith("Cannot remove"):
                messagebox.showerror("Quantity error", msg)
            else:
                messagebox.showinfo("Done", msg)

            self.msg_var.set(msg)
            self.refresh_table()
        except PermissionError:
            messagebox.showerror("Locked", "Close the Excel file and try again.")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def on_normalize_all(self):
        try:
            changed = normalize_all(self.user_var.get().strip())
            self.msg_var.set(f"Normalized {changed} item(s).")
            self.refresh_table()
            if changed:
                messagebox.showinfo("Normalization complete", f"Normalized {changed} item(s).")
        except PermissionError:
            messagebox.showerror("Locked", "Close the Excel file and try again.")
        except Exception as e:
            messagebox.showerror("Error", str(e))


    def open_excel(self):
        # Attempt to open with OS default app
        try:
            path = EXCEL_PATH.resolve()
            import os, platform, subprocess
            if platform.system() == "Windows":
                os.startfile(str(path))
            elif platform.system() == "Darwin":
                subprocess.run(["open", str(path)])
            else:
                subprocess.run(["xdg-open", str(path)])
        except Exception as e:
            messagebox.showerror("Open failed", str(e))
     

if __name__ == "__main__":
    app = App()
    app.mainloop()

