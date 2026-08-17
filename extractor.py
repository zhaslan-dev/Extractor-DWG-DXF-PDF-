"""
Generic DXF Table Parser
========================
- Extracts structured table data from DXF files
- Groups rows, detects headers, extracts cell values
- Exports results to Excel
"""

import os, sys, re, logging, traceback
from pathlib import Path
from collections import defaultdict
from dataclasses import dataclass, field
from typing import List, Dict, Optional, Tuple, Set, Any

import ezdxf
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ─── Configuration ─────────────────────────────────────────────────────────────
USE_Y_BUCKETS = True
Y_BUCKET_SIZE = 10.0
VERSION = "generic_1"
MIN_ID_LEN = 4

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    handlers=[
        logging.FileHandler("dxf_parser.log", encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)
log = logging.getLogger(__name__)

# ══════════════════════════════════════════════════════════════════════
#  GENERIC PATTERNS
# ══════════════════════════════════════════════════════════════════════
ANCHOR_PATTERN = re.compile(
    r'^'
    r'(\d+/'
    r'[\d]+'
    r'(?:\.\d+)*'
    r'(?:\.[A-Z]\d*)?'
    r'(?:[\._]\w+)*'
    r')'
    r'\s*:\s*'
    r'([A-Za-z][A-Za-z0-9_]*)'
    r'$',
    re.I
)

ID_PATTERN = re.compile(
    r'^'
    r'\d+/'
    r'\d+'
    r'(?:\.\d+)*'
    r'(?:\.[A-Z]\d*)?'
    r'(?:[\._]\w+)*'
    r'$',
    re.I
)

# ─── Generic column offsets (example) ────────────────────────────────────────
COLUMN_OFFSETS = {
    "ColB": 45.6,
    "ColB_Width": 10.5,
    "ColC": 76.0,
    "ColC_Width": 12.34,
    "Tolerance": 5.0,
}

# ══════════════════════════════════════════════════════════════════════
#  STRING CLEANING
# ══════════════════════════════════════════════════════════════════════
def decode_unicode_escapes(text: str) -> str:
    def repl(match):
        code = match.group(1)
        try:
            int_code = int(code, 16)
            if int_code > 0x10FFFF:
                return match.group(0)
            return chr(int_code)
        except (ValueError, OverflowError):
            return match.group(0)
    text = re.sub(r'\\U\+([0-9A-Fa-f]{4,8})', repl, text)
    text = re.sub(r'\\u([0-9A-Fa-f]{4})', repl, text)
    return text

def canonical_id(s: str) -> str:
    if not s:
        return ""
    return re.sub(r'\s+', '', s).strip().upper()

# ══════════════════════════════════════════════════════════════════════
#  BASIC FILTERS (simplified, no domain knowledge)
# ══════════════════════════════════════════════════════════════════════
RE_CHINESE = re.compile(r'[\u4e00-\u9fff]')
_SIZE_RE = re.compile(r'^\d+(\.\d+)?[*xX]\d+(\.\d+)?$')

# Minimal blacklist – just technical labels that appear in any table header
_FORBIDDEN_IDS = {
    "Col1","Col2","Col3","Col4","Col5","Col6","Col7","Col8",
    "Col9","Col10","Col11","Col12","Col13",
}

_FORBIDDEN_PATTERNS = [
    re.compile(r'^BWG\d', re.I),
    re.compile(r'^SCS\d', re.I),
    re.compile(r'^\d+/\d+'),
    re.compile(r'^T2_\d+_\d+'),
    re.compile(r'^MOAS_'),
    re.compile(r'^\d{2,3}$'),
    re.compile(r'^PVC$', re.I),
    re.compile(r'^\d+mm$', re.I),
]

def is_forbidden_id(pn: str) -> bool:
    pn = pn.strip()
    if not pn or pn in _FORBIDDEN_IDS:
        return True
    if len(pn) < MIN_ID_LEN:
        return True
    for pat in _FORBIDDEN_PATTERNS:
        if pat.match(pn):
            return True
    return False

# Minimal extended checks – just length and digit presence
def is_strict_id_v2(text: str) -> bool:
    t = text.strip()
    if len(t) < 5:
        return False
    if not any(c.isdigit() for c in t):
        return False
    if RE_CHINESE.search(t) or "&" in t or "mm" in t.lower():
        return False
    if not re.match(r'^[A-Z0-9\-_/.]{5,}$', t, re.I):
        return False
    return True

def is_strict_id(text: str) -> bool:
    t = text.strip()
    if len(t) < 4: return False
    if RE_CHINESE.search(t) or any(c in t for c in "\\|"): return False
    return is_strict_id_v2(t)

def is_clean_label(text: str) -> bool:
    t = text.strip()
    if len(t) < 3 or RE_CHINESE.search(t): return False
    if re.search(r'^[A-Z]:\\', t, re.I) or t.lower().endswith(('.dxf', '.dwg')): return False
    if t.upper().startswith(("CLIP", "TIE", "PAGE", "PVC")): return False
    if re.match(r'^[V][A-Z\d]+[D]$', t, re.I) or "*" in t: return False
    return True

# ══════════════════════════════════════════════════════════════════════
#  GENERIC VALIDATION FUNCTIONS (replace domain-specific ones)
# ══════════════════════════════════════════════════════════════════════
_ITEM_PATTERNS_A = [   # previously terminal patterns
    re.compile(r'^\d-\d{6,7}-\d$'),
    re.compile(r'^\d{6,7}-\d{1,2}$'),
    re.compile(r'^\d{4,6}-\d{1,2}$'),
    re.compile(r'^\d-\d{7}-\d$'),
    re.compile(r'^\d{10}$'),
    re.compile(r'^\d{10,12}$'),
    re.compile(r'^\d{8}$'),
    re.compile(r'^\d{4}-\d{4}-\d{2}$'),
    re.compile(r'^\d{4}-\d{4}$'),
    re.compile(r'^\d{6,8}[A-Z]{2,4}$'),
    re.compile(r'^\d{7,8}$'),
    re.compile(r'^[A-Z]{2}\d{6,}(-\d+)?$'),
    re.compile(r'^PP\d{7}$'),
    re.compile(r'^\d{3}-\d{3}-\d{3}$'),
    re.compile(r'^\d{5,}[A-Z]{2,4}$'),
]

_ITEM_PATTERNS_B = [   # previously seal patterns
    re.compile(r'^\d{5,7}-\d$'),
    re.compile(r'^\d{4}-\d{4}-\d{2}$'),
    re.compile(r'^\d{4}-\d{4}$'),
    re.compile(r'^JD\d{6,}(-\d+)?$'),
    re.compile(r'^JD\d{1,2}-\d{6}-\d$'),
    re.compile(r'^JD\d{4}-\d{4}-\d{2}$'),
    re.compile(r'^JD\d{9,}$'),
    re.compile(r'^HDF\d{3}$'),
    re.compile(r'^GM\d{7}$'),
    re.compile(r'^\d{9,10}$'),
]

def is_item_type_a(text: str) -> bool:
    t = text.strip()
    if not t: return False
    for pat in _ITEM_PATTERNS_A:
        if pat.match(t):
            return True
    return False

def is_item_type_b(text: str) -> bool:
    t = text.strip()
    if not t: return False
    for pat in _ITEM_PATTERNS_B:
        if pat.match(t):
            return True
    return False

_GENERIC_ID_PATTERNS = _ITEM_PATTERNS_A + _ITEM_PATTERNS_B

def _build_strict_patterns():
    patterns = list(_GENERIC_ID_PATTERNS)
    seen = set()
    unique = []
    for p in patterns:
        if p.pattern not in seen:
            seen.add(p.pattern)
            unique.append(p)
    loose = []
    for p in unique:
        s = p.pattern
        if s.startswith('^'): s = s[1:]
        if s.endswith('$'): s = s[:-1]
        s = r'\b' + s + r'\b'
        loose.append(re.compile(s))
    return loose

_STRICT_ID_PATTERNS = _build_strict_patterns()

def extract_strict_id(text: str) -> Optional[str]:
    for pat in _STRICT_ID_PATTERNS:
        m = pat.search(text)
        if m:
            return m.group(0)
    return None

def is_valid_id(pn: str) -> bool:
    if not pn: return False
    pn = pn.strip()
    if len(pn) < MIN_ID_LEN: return False
    if is_forbidden_id(pn): return False
    if is_strict_id_v2(pn): return True
    return False

def normalize_accessory(acc: str) -> Tuple[str, str, str]:
    # simplified – just pass through
    return ("", "", "")

# ══════════════════════════════════════════════════════════════════════
#  DATA STRUCTURES
# ══════════════════════════════════════════════════════════════════════
@dataclass
class RowData:
    col1: str = ""   # formerly pin
    col2: str = ""   # formerly terminal
    col3: str = ""   # formerly seal
    col4: str = ""   # supplier A
    col5: str = ""   # supplier B
    status: str = "Active"

@dataclass
class TableBlock:
    position: str = ""
    label: str = ""
    raw_id: str = ""
    manufacturer: str = ""
    pn: str = ""
    tpa: str = ""
    rows: List[RowData] = field(default_factory=list)
    id_x: float = 0.0
    id_y: float = 0.0
    components: Dict[str, Tuple[str, str, str, int]] = field(default_factory=dict)

_TYPE_MAP = {}  # removed domain classification
def classify(name: str) -> str:
    return "Other"

_MTEXT_RE = re.compile(
    r"\{\\f[^;]*;([^}]*)\}"
    r"|\{\\[^}]*\}"
    r"|\\[A-Za-z][0-9.]*;"
    r"|\\[Pp](?=[^\w]|$)"
)
def clean(raw: str) -> str:
    t = _MTEXT_RE.sub(lambda m: m.group(1) or "", raw)
    t = re.sub(r'\\f[^;]*;', '', t)
    t = t.replace("\\~", " ").replace("\r\n", " ").replace("\r", " ").strip()
    t = decode_unicode_escapes(t)
    return t

# ══════════════════════════════════════════════════════════════════════
#  PARSER
# ══════════════════════════════════════════════════════════════════════
class DXFParser:
    Y_TOL = 2
    RED_RADIUS = 14.0

    def __init__(self, path: str):
        self.path = str(path)
        self.name = Path(path).stem
        self._texts: List[dict] = []
        self._red_centers: List[Tuple[float, float]] = []
        self._y_buckets: Optional[Dict[int, List[dict]]] = None
        self._anchor_cache: Dict[str, List[Tuple[float, float]]] = {}
        self._connector_anchor_cache: Dict[str, List[Tuple[float, float]]] = {}
        self._table_regions: List[Tuple[float, float, float, float]] = []

    def _build_y_buckets(self):
        self._y_buckets = defaultdict(list)
        for t in self._texts:
            key = int(t["y"] // Y_BUCKET_SIZE)
            self._y_buckets[key].append(t)

    def _nearby(self, x: float, y: float, dy: float, dx: float) -> List[dict]:
        if self._y_buckets is not None:
            y_min, y_max = y - dy, y + dy
            start = int(y_min // Y_BUCKET_SIZE) - 1
            end   = int(y_max // Y_BUCKET_SIZE) + 2
            res = []
            for b in range(start, end):
                for t in self._y_buckets.get(b, []):
                    if abs(t["y"] - y) <= dy and abs(t["x"] - x) <= dx:
                        res.append(t)
            return res
        else:
            return [t for t in self._texts if abs(t["y"] - y) <= dy and abs(t["x"] - x) <= dx]

    def _load(self):
        for enc in ["gbk", "cp936", "utf-8", "latin1"]:
            try:
                doc = ezdxf.readfile(self.path, encoding=enc)
                log.info(f"  encoding: {enc}, DXF {doc.dxfversion}")
                return doc
            except Exception:
                pass
        raise RuntimeError(f"Cannot open: {self.path}")

    def _collect_texts(self, msp):
        items = []
        for e in msp.query("MTEXT"):
            try:
                t = clean(e.text)
                if t:
                    if RE_CHINESE.search(t) or _SIZE_RE.match(t):
                        continue
                    items.append({"x": e.dxf.insert.x, "y": e.dxf.insert.y,
                                  "text": t, "h": e.dxf.char_height, "color": e.dxf.color})
            except Exception:
                pass
        for e in msp.query("TEXT"):
            try:
                t = e.dxf.text.strip()
                if t:
                    t = decode_unicode_escapes(t)
                    if RE_CHINESE.search(t) or _SIZE_RE.match(t):
                        continue
                    items.append({"x": e.dxf.insert.x, "y": e.dxf.insert.y,
                                  "text": t, "h": e.dxf.height, "color": e.dxf.color})
            except Exception:
                pass
        return items

    def _collect_red(self, msp):
        pts = []
        for e in msp.query("LWPOLYLINE"):
            try:
                if e.dxf.color == 1:
                    pp = list(e.get_points("xy"))
                    xs = [p[0] for p in pp]
                    ys = [p[1] for p in pp]
                    pts.append((sum(xs) / len(xs), sum(ys) / len(ys)))
            except Exception:
                pass
        return pts

    def _is_red_flag_near(self, x: float, y: float, radius: float = None) -> bool:
        if radius is None:
            radius = self.RED_RADIUS
        return any(abs(x - fx) <= radius and abs(y - fy) <= radius
                   for fx, fy in self._red_centers)

    def _cluster_rows(self, items: List[dict]) -> List[List[dict]]:
        if not items: return []
        s = sorted(items, key=lambda t: -t["y"])
        rows = [[s[0]]]
        for t in s[1:]:
            if abs(t["y"] - rows[-1][0]["y"]) <= self.Y_TOL:
                rows[-1].append(t)
            else:
                rows.append([t])
        return rows

    def _inside_table(self, x: float, y: float) -> bool:
        for x1, x2, y1, y2 in self._table_regions:
            if x1 <= x <= x2 and y1 <= y <= y2:
                return True
        return False

    def _build_anchor_cache(self):
        self._anchor_cache = defaultdict(list)
        for t in self._texts:
            m = ANCHOR_PATTERN.match(t["text"])
            if m:
                conn_id = re.sub(r'\s', '', m.group(1))
                self._anchor_cache[conn_id].append((t["x"], t["y"]))

    def _build_connector_anchors(self):
        self._connector_anchor_cache.clear()
        for t in self._texts:
            if not ID_PATTERN.match(t["text"]):
                continue
            if self._inside_table(t["x"], t["y"]):
                continue
            if not self._has_connector_nearby(t["x"], t["y"]):
                continue
            conn_id = t["text"]
            self._connector_anchor_cache.setdefault(conn_id, []).append((t["x"], t["y"]))

    def _has_connector_nearby(self, x: float, y: float, radius: float = 15.0) -> bool:
        nearby = self._nearby(x, y, radius, radius)
        for t in nearby:
            txt = t["text"].strip()
            # simplified: just look for colon pattern
            if re.match(r'^[A-Z][A-Z0-9_]{1,12}:\d', txt):
                return True
            if re.match(r"TPA\s*:\s*(.+)", txt, re.I):
                return True
        return False

    # ─── TABLE LOCATION ──────────────────────────────────────────
    def _detect_col_groups(self, hdr_row: List[dict]) -> List[List[dict]]:
        if not hdr_row:
            return []
        sorted_hdrs = sorted(hdr_row, key=lambda t: t['x'])
        if len(sorted_hdrs) <= 1:
            return []
        gaps = [sorted_hdrs[i]['x'] - sorted_hdrs[i-1]['x'] for i in range(1, len(sorted_hdrs))]
        mean_gap = sum(gaps) / len(gaps)
        variance = sum((g - mean_gap) ** 2 for g in gaps) / len(gaps)
        std_dev = variance ** 0.5
        gap_threshold = max(20.0, mean_gap + 2.5 * std_dev)

        groups = []
        current_group = [sorted_hdrs[0]]
        for i in range(1, len(sorted_hdrs)):
            gap = sorted_hdrs[i]['x'] - sorted_hdrs[i-1]['x']
            if gap > gap_threshold:
                groups.append(current_group)
                current_group = [sorted_hdrs[i]]
            else:
                current_group.append(sorted_hdrs[i])
        if current_group:
            groups.append(current_group)
        if len(groups) > 1:
            log.debug("MCT groups: %d (threshold=%.1f)", len(groups), gap_threshold)
        return groups

    def _compute_right_bound(self, rows: List[List[dict]], x_pin: float, header_x_map: Dict[str, float] = None) -> float:
        # simplified: just take max X in first few rows + 10
        all_x = []
        for row in rows[:10]:
            for t in row:
                if len(t["text"].strip()) >= 3:
                    all_x.append(t["x"])
        if all_x:
            return max(all_x) + 10.0
        return x_pin + 200.0

    def _locate_table(self, anchor_x: float, anchor_y: float) -> Optional[Dict]:
        texts_below = [t for t in self._texts if t["y"] < anchor_y - 2.0]
        if not texts_below:
            log.info(f"Anchor ({anchor_x:.1f}, {anchor_y:.1f}): no texts below")
            return None

        temp_right = max(t['x'] for t in texts_below) + 50.0
        y_bottom = min(t["y"] for t in self._texts) - 10.0
        table_texts = [t for t in self._texts
                       if anchor_x - 2.0 <= t["x"] <= temp_right
                       and y_bottom <= t["y"] < anchor_y - 2.0]
        if len(table_texts) < 3:
            log.info(f"Anchor ({anchor_x:.1f}, {anchor_y:.1f}): too few texts ({len(table_texts)})")
            return None

        rows = self._cluster_rows(table_texts)
        rows.sort(key=lambda r: r[0]["y"], reverse=True)

        # Generic header names (13 generic column labels)
        HEADER_NAMES = {f"Col{i}" for i in range(1, 14)}

        def count_header_matches(text: str) -> int:
            t_clean = text.upper().replace(" ", "")
            found = set()
            for name in HEADER_NAMES:
                if name.upper().replace(" ", "") in t_clean:
                    found.add(name)
            return len(found)

        header_row = None
        header_x_map = {}
        separate_headers = False
        col_groups = []

        first_row = rows[0] if rows else []
        merged_text = " ".join(t["text"] for t in first_row)
        header_match_count = count_header_matches(merged_text)
        if header_match_count >= 13:   # all 13 generic headers present
            header_row = first_row
            for name in HEADER_NAMES:
                for t in first_row:
                    if name.upper().replace(" ", "") in t["text"].upper().replace(" ", ""):
                        header_x_map[name] = t["x"]
                        break
            if len(header_x_map) >= 12:
                separate_headers = True
                col_groups = self._detect_col_groups(first_row)
                if len(col_groups) > 1:
                    log.info(f"  detected multi-column table ({len(col_groups)} columns)")
        else:
            # search by "1" and "2"
            for i, row in enumerate(rows):
                for t in row:
                    if t["text"].strip() == "1":
                        if i+1 < len(rows) and any(t2["text"].strip() == "2" and abs(t2["x"] - t["x"]) < 6.0
                                                  for t2 in rows[i+1]):
                            if i > 0:
                                prev_row = rows[i-1]
                                merged_prev = "".join(tx["text"].strip().replace(" ", "") for tx in prev_row)
                                if count_header_matches(merged_prev) >= 13:
                                    header_row = row
                                    break
                            if not header_row and i > 1:
                                prev_prev_row = rows[i-2]
                                for pt in prev_prev_row:
                                    if ANCHOR_PATTERN.match(pt["text"]) and abs(pt["x"] - t["x"]) < 10.0:
                                        header_row = row
                                        break
                if header_row:
                    break
        if not header_row:
            first_row_text = " ".join(t["text"] for t in first_row) if first_row else ""
            matches = count_header_matches(first_row_text)
            log.info(f"Anchor ({anchor_x:.1f}, {anchor_y:.1f}): header not found (matches: {matches}/{len(HEADER_NAMES)})")
            return None

        x_pin = None
        if separate_headers and "Col1" in header_x_map:
            x_pin = header_x_map["Col1"]
        else:
            for t in header_row:
                if t["text"].strip() == "1":
                    x_pin = t["x"]
                    break
            if x_pin is None and len(rows) > 1:
                for t in rows[1]:
                    if t["text"].strip() == "1":
                        x_pin = t["x"]
                        break
        if x_pin is None:
            log.info(f"Anchor ({anchor_x:.1f}, {anchor_y:.1f}): could not determine Col1 X")
            return None

        data_start_idx = rows.index(header_row) + 1
        if data_start_idx >= len(rows):
            log.info(f"Anchor ({anchor_x:.1f}, {anchor_y:.1f}): no data rows after header")
            return None

        has_contacts = False
        last_contact_idx = data_start_idx
        for i in range(data_start_idx, len(rows)):
            row = rows[i]
            pin_texts = [t for t in row if abs(t["x"] - x_pin) < 6.0]
            if not pin_texts:
                break
            num_text = pin_texts[0]["text"].strip()
            if not num_text.isdigit():
                break
            if num_text == "1":
                has_contacts = True
            last_contact_idx = i

        if not has_contacts:
            log.info(f"Anchor ({anchor_x:.1f}, {anchor_y:.1f}): no numbered rows in Col1")
            return None

        data_rows = rows[data_start_idx:last_contact_idx+1]
        rightmost_x = self._compute_right_bound(data_rows[:10], x_pin, header_x_map if separate_headers else None)
        left_bound = x_pin - 2.0

        result = {
            'hdr_row': header_row,
            'x_pin': x_pin,
            'header_cols': header_x_map if separate_headers else None,
            'rows': data_rows,
            'left_bound': left_bound,
            'right_bound': rightmost_x,
            'y_bottom_actual': data_rows[-1][0]["y"] - 2.0 if data_rows else anchor_y - 200
        }
        if col_groups:
            groups_info = []
            for grp in col_groups:
                grp_map = {}
                pin_seen = False
                for t in grp:
                    name = t['text']
                    if name == 'Col1':
                        name = 'Col13' if pin_seen else 'Col1'
                        if name == 'Col1': pin_seen = True
                    if name not in grp_map:
                        grp_map[name] = t['x']
                x_vals = [t['x'] for t in grp]
                grp_x_pin = grp_map.get('Col1', x_pin)
                groups_info.append({
                    'cols': grp_map,
                    'x_min': min(x_vals) - 5,
                    'x_max': max(x_vals) + 5,
                    'x_pin': grp_x_pin
                })
            result['col_groups'] = groups_info
        return result

    # ─── ROW PARSING ────────────────────────────────────────────
    def _parse_rows(self, block: TableBlock, table_info: Dict):
        x_pin = table_info['x_pin']
        header_cols = table_info.get('header_cols')
        data_rows = table_info.get('rows', [])
        col_groups = table_info.get('col_groups')
        right_bound = table_info.get('right_bound', x_pin + 200.0)

        if col_groups:
            for grp_info in col_groups:
                cols = grp_info['cols']
                x_min = grp_info['x_min']
                x_max = grp_info['x_max']
                grp_x_pin = grp_info.get('x_pin', x_pin)

                has_col2 = 'Col2' in cols
                has_col3 = 'Col3' in cols

                filled_rows = 0
                for row in data_rows:
                    col1_text = ""
                    col2_text = ""
                    col3_text = ""
                    col4_text = ""
                    col5_text = ""

                    for t in row:
                        if not (x_min - 5 <= t["x"] <= x_max + 5):
                            continue
                        txt = t["text"].strip()
                        if abs(t["x"] - grp_x_pin) < 6.0:
                            col1_text = txt
                        elif has_col2 and abs(t["x"] - cols['Col2']) <= 10.0:
                            col2_text = self._extract_item_a(txt)
                            if col2_text:
                                col4_text = self._extract_prefix(txt)
                        elif has_col3 and abs(t["x"] - cols['Col3']) <= 10.0:
                            col3_text = self._extract_item_b(txt)
                            if col3_text:
                                col5_text = self._extract_prefix(txt)

                    if col2_text or col3_text:
                        block.rows.append(RowData(col1=col1_text, col2=col2_text, col3=col3_text,
                                                  col4=col4_text, col5=col5_text))
                        filled_rows += 1

                if not has_col2 or not has_col3 or filled_rows < 0.5 * len(data_rows):
                    for row in data_rows:
                        col1_text = ""
                        candidates = []
                        for t in row:
                            if not (x_min - 5 <= t["x"] <= x_max + 5):
                                continue
                            txt = t["text"].strip()
                            if abs(t["x"] - grp_x_pin) < 6.0:
                                col1_text = txt
                            else:
                                a = self._extract_item_a(txt)
                                if a:
                                    candidates.append((t['x'], 'A', a, self._extract_prefix(txt)))
                                b = self._extract_item_b(txt)
                                if b:
                                    candidates.append((t['x'], 'B', b, self._extract_prefix(txt)))
                        if candidates:
                            found_a = None
                            found_b = None
                            for cx, typ, pn, supp in candidates:
                                if typ == 'A' and not found_a:
                                    found_a = (pn, supp)
                                elif typ == 'B' and not found_b:
                                    found_b = (pn, supp)
                                if found_a and found_b:
                                    break
                            if found_a or found_b:
                                block.rows.append(RowData(col1=col1_text,
                                                         col2=found_a[0] if found_a else "",
                                                         col3=found_b[0] if found_b else "",
                                                         col4=found_a[1] if found_a else "",
                                                         col5=found_b[1] if found_b else ""))
            self._log_statistics(block, data_rows)
            return

        if header_cols and len(header_cols) >= 12 and all(k in header_cols for k in ["Col2", "Col3"]):
            col2_x = header_cols.get("Col2")
            col3_x = header_cols.get("Col3")
            tol = 10.0

            filled_rows = 0
            for row in data_rows:
                col1_text = ""
                col2_text = ""
                col3_text = ""
                col4_text = ""
                col5_text = ""

                for t in row:
                    txt = t["text"].strip()
                    if abs(t["x"] - x_pin) < 6.0:
                        col1_text = txt
                    elif col2_x and abs(t["x"] - col2_x) <= tol:
                        col2_text = self._extract_item_a(txt)
                        if col2_text:
                            col4_text = self._extract_prefix(txt)
                    elif col3_x and abs(t["x"] - col3_x) <= tol:
                        col3_text = self._extract_item_b(txt)
                        if col3_text:
                            col5_text = self._extract_prefix(txt)

                if col2_text or col3_text:
                    block.rows.append(RowData(col1=col1_text, col2=col2_text, col3=col3_text,
                                              col4=col4_text, col5=col5_text))
                    filled_rows += 1

            if filled_rows < 0.5 * len(data_rows):
                block.rows.clear()
                for row in data_rows:
                    col1_text = ""
                    candidates = []
                    for t in row:
                        rel_x = t["x"] - x_pin
                        txt = t["text"].strip()
                        if abs(rel_x) < 6.0:
                            col1_text = txt
                        elif rel_x > 6 and t["x"] <= right_bound + 2.0:
                            a = self._extract_item_a(txt)
                            if a:
                                candidates.append((t['x'], 'A', a, self._extract_prefix(txt)))
                            b = self._extract_item_b(txt)
                            if b:
                                candidates.append((t['x'], 'B', b, self._extract_prefix(txt)))
                    if candidates:
                        found_a = None
                        found_b = None
                        for cx, typ, pn, supp in candidates:
                            if typ == 'A' and not found_a:
                                found_a = (pn, supp)
                            elif typ == 'B' and not found_b:
                                found_b = (pn, supp)
                            if found_a and found_b:
                                break
                        if found_a or found_b:
                            block.rows.append(RowData(col1=col1_text,
                                                     col2=found_a[0] if found_a else "",
                                                     col3=found_b[0] if found_b else "",
                                                     col4=found_a[1] if found_a else "",
                                                     col5=found_b[1] if found_b else ""))
            self._log_statistics(block, data_rows)
            return

        # Fallback: merged headers
        for row in data_rows:
            col1_text = ""
            candidates = []
            for t in row:
                rel_x = t["x"] - x_pin
                txt = t["text"].strip()
                if abs(rel_x) < 6.0:
                    col1_text = txt
                elif rel_x > 6 and t["x"] <= right_bound + 2.0:
                    a = self._extract_item_a(txt)
                    if a:
                        candidates.append((t['x'], 'A', a, self._extract_prefix(txt)))
                    b = self._extract_item_b(txt)
                    if b:
                        candidates.append((t['x'], 'B', b, self._extract_prefix(txt)))
            if candidates:
                found_a = None
                found_b = None
                for cx, typ, pn, supp in candidates:
                    if typ == 'A' and not found_a:
                        found_a = (pn, supp)
                    elif typ == 'B' and not found_b:
                        found_b = (pn, supp)
                    if found_a and found_b:
                        break
                if found_a or found_b:
                    block.rows.append(RowData(col1=col1_text,
                                             col2=found_a[0] if found_a else "",
                                             col3=found_b[0] if found_b else "",
                                             col4=found_a[1] if found_a else "",
                                             col5=found_b[1] if found_b else ""))
        self._log_statistics(block, data_rows)

    def _extract_item_a(self, txt: str) -> str:
        txt = txt.strip()
        if is_item_type_a(txt):
            return txt
        extracted = extract_strict_id(txt)
        if extracted and is_item_type_a(extracted):
            return extracted
        return ""

    def _extract_item_b(self, txt: str) -> str:
        txt = txt.strip()
        if is_item_type_b(txt):
            return txt
        extracted = extract_strict_id(txt)
        if extracted and is_item_type_b(extracted):
            return extracted
        return ""

    def _extract_prefix(self, txt: str) -> str:
        return ""

    def _log_statistics(self, block: TableBlock, data_rows: List[List[dict]]):
        set_a = set()
        set_b = set()
        for r in block.rows:
            if r.col2:
                set_a.add(r.col2)
            if r.col3:
                set_b.add(r.col3)
        log.info(f"  table {block.position}: processed {len(block.rows)} rows "
                 f"(unique A: {len(set_a)}, unique B: {len(set_b)})")

    def _parse_header(self, id_x: float, id_y: float, block: TableBlock):
        area = [t for t in self._texts if abs(t["x"] - id_x) <= 7.0
                and -14.0 <= t["y"] - id_y <= 35.0]
        area.sort(key=lambda t: -t["y"])
        valid_labels = []
        processed = set()
        for t in area:
            txt = t["text"].strip()
            if "TPA" in txt.upper() and ":" in txt:
                m = re.match(r"TPA\s*:\s*(.+)", txt, re.I)
                if m and is_strict_id(m.group(1).strip()):
                    block.tpa = m.group(1).strip()
                    processed.add(txt)
                continue
            if ";" in txt:
                parts = txt.split(";", 1)
                desc, pn = parts[0].strip(), parts[1].strip()
                if is_strict_id(pn):
                    block.components[f"acc_{pn}"] = (desc, "", pn, 1)
                    processed.add(txt)
                continue
            m = re.match(r'^([A-Z][A-Z0-9_]{1,12}):(.{3,})$', txt)
            if m and not block.raw_id:
                p_pn = m.group(2).strip()
                if is_strict_id(p_pn):
                    block.raw_id = txt
                    block.manufacturer, block.pn = m.group(1), p_pn
                    processed.add(txt)
                continue
            if is_clean_label(txt) and ":" not in txt:
                valid_labels.append(txt)
        if valid_labels:
            block.label = max(valid_labels, key=len)
        block.conn_type = classify(block.label)

    def _aggregate_components(self, b: TableBlock):
        comps = {}
        if b.pn:
            desc = b.label if b.label else f"{b.manufacturer}:{b.pn}"
            comps["housing"] = (desc, b.manufacturer, b.pn, 1)
        if b.tpa:
            comps["tpa"] = (f"TPA {b.tpa}", "", b.tpa, 1)

        counter_a = {}
        counter_b = {}
        for r in b.rows:
            if r.col2:
                key = canonical_id(r.col2)
                if key not in counter_a:
                    counter_a[key] = (r.col4, r.col2, 1)
                else:
                    mfr, pn, qty = counter_a[key]
                    counter_a[key] = (mfr, pn, qty + 1)
            if r.col3:
                key = canonical_id(r.col3)
                if key not in counter_b:
                    counter_b[key] = (r.col5, r.col3, 1)
                else:
                    mfr, pn, qty = counter_b[key]
                    counter_b[key] = (mfr, pn, qty + 1)
        for key, (mfr, pn, qty) in counter_a.items():
            comps[f"itemA_{key}"] = ("ItemA", mfr, pn, qty)
        for key, (mfr, pn, qty) in counter_b.items():
            comps[f"itemB_{key}"] = ("ItemB", mfr, pn, qty)

        b.components = comps

    def _extract_loose_components(self) -> List[dict]:
        return []  # simplified

    def parse(self) -> Tuple[List[TableBlock], List[dict]]:
        log.info(f"▶ {self.name}")
        try:
            doc = self._load()
            msp = doc.modelspace()
            self._texts = self._collect_texts(msp)
            self._red_centers = self._collect_red(msp)
            if USE_Y_BUCKETS:
                self._build_y_buckets()
            log.info(f"  texts: {len(self._texts)}, red flags: {len(self._red_centers)}")
        except Exception as e:
            log.error(f"  Load error: {e}")
            return [], []

        self._build_anchor_cache()
        log.info(f"  anchors: {sum(len(v) for v in self._anchor_cache.values())}")

        id_pins: Dict[str, List[RowData]] = defaultdict(list)
        table_bottoms = {}
        for conn_id, positions in self._anchor_cache.items():
            for ax, ay in positions:
                table_info = self._locate_table(ax, ay)
                if not table_info:
                    continue
                block = TableBlock(position=conn_id)
                self._parse_rows(block, table_info)
                if block.rows:
                    id_pins[conn_id].extend(block.rows)
                if conn_id not in table_bottoms or table_info['y_bottom_actual'] < table_bottoms[conn_id]:
                    table_bottoms[conn_id] = table_info['y_bottom_actual']

        self._table_regions.clear()
        for conn_id, positions in self._anchor_cache.items():
            for ax, ay in positions:
                bottom_y = table_bottoms.get(conn_id, ay - 200)
                self._table_regions.append((ax - 40, ax + 120, bottom_y - 5, ay + 5))

        self._build_connector_anchors()
        log.info(f"  clean IDs (outside tables): {len(self._connector_anchor_cache)}")

        blocks = []
        used_ids = set()
        for conn_id, positions in self._connector_anchor_cache.items():
            if conn_id in used_ids:
                continue
            used_ids.add(conn_id)
            cx, cy = positions[0]
            block = TableBlock(position=conn_id, id_x=cx, id_y=cy)
            self._parse_header(cx, cy, block)
            block.rows = id_pins.get(conn_id, [])
            self._aggregate_components(block)
            if block.pn or block.rows:
                blocks.append(block)

        loose = self._extract_loose_components()
        log.info(f"  blocks: {len(blocks)}, loose: {len(loose)}")
        return blocks, loose


# ══════════════════════════════════════════════════════════════════════
#  EXCEL STYLES (unchanged)
# ══════════════════════════════════════════════════════════════════════
def clean_xml_string(s):
    if not isinstance(s, str):
        s = str(s)
    s = re.sub(r'[\x00-\x08\x0B\x0C\x0E-\x1F]', '', s)
    s = re.sub(r'[^\x09\x0A\x0D\x20-\uD7FF\uE000-\uFFFD]', '', s, flags=re.UNICODE)
    if len(s) > 32000:
        s = s[:32000] + "..."
    return s

def _s(color="B8CCE4"): return Side(border_style="thin", color=color)
_BDR  = Border(left=_s(), right=_s(), top=_s(), bottom=_s())
_BDR2 = Border(left=_s("FFFFFF"), right=_s("FFFFFF"), top=_s("FFFFFF"), bottom=_s("AAAAAA"))
_F = {
    "dark":    PatternFill("solid", fgColor="17375E"),
    "blue":    PatternFill("solid", fgColor="1F497D"),
    "changed": PatternFill("solid", fgColor="FFEB9C"),
    "deleted": PatternFill("solid", fgColor="FFC7CE"),
    "alt":     PatternFill("solid", fgColor="EEF3FF"),
    "white":   PatternFill("solid", fgColor="FFFFFF"),
    "green":   PatternFill("solid", fgColor="E2EFDA"),
    "lblue2":  PatternFill("solid", fgColor="DEEAF1"),
}
_FN = {
    "title": Font(name="Arial", bold=True, color="FFFFFF", size=13),
    "hdr":   Font(name="Arial", bold=True, color="FFFFFF", size=9),
    "body":  Font(name="Arial", size=9),
    "bold":  Font(name="Arial", bold=True, size=9),
}
C = Alignment(horizontal="center", vertical="center", wrap_text=True)
L = Alignment(horizontal="left",   vertical="center", wrap_text=True)

def hc(ws, row, col, val, w=None):
    if isinstance(val, str): val = clean_xml_string(val)
    c = ws.cell(row, col, val)
    c.font = _FN["hdr"]; c.fill = _F["blue"]; c.alignment = C; c.border = _BDR2
    if w: ws.column_dimensions[get_column_letter(col)].width = w
    ws.row_dimensions[row].height = 18
    return c

def dc(ws, row, col, val, fill=None, align=L, bold=False):
    if isinstance(val, str): val = clean_xml_string(val)
    c = ws.cell(row, col, val)
    c.font = _FN["bold"] if bold else _FN["body"]
    c.alignment = align; c.border = _BDR
    if fill: c.fill = fill
    return c


# ══════════════════════════════════════════════════════════════════════
#  EXPORT (generic names)
# ══════════════════════════════════════════════════════════════════════
def export_single(blocks: List[TableBlock], loose_parts: List[dict], file_label: str, out_path: str):
    wb = openpyxl.Workbook()
    ws_sum = wb.active
    ws_sum.title = "Summary"
    ws_sum.sheet_view.showGridLines = False
    ws_sum.merge_cells("A1:E1")
    ws_sum["A1"].value = f"Summary · {file_label}"
    ws_sum["A1"].font = _FN["title"]
    ws_sum["A1"].fill = _F["dark"]
    ws_sum["A1"].alignment = C
    ws_sum.row_dimensions[1].height = 24
    headers_sum = [("ID", 22), ("Description", 42), ("Mfr", 16), ("Qty", 10), ("Type", 14)]
    for ci, (h, w) in enumerate(headers_sum, 1):
        hc(ws_sum, 2, ci, h, w)
    pn_counter = {}
    for b in blocks:
        for key, (desc, mfr, pn, qty) in b.components.items():
            if key == "housing":
                pn_counter[pn] = {"desc": desc, "mfr": mfr, "type": "Housing",
                                  "qty": pn_counter.get(pn, {}).get("qty", 0) + qty}
            elif key == "tpa":
                pn_counter[pn] = {"desc": desc, "mfr": mfr, "type": "TPA",
                                  "qty": pn_counter.get(pn, {}).get("qty", 0) + qty}
            else:
                pn_counter[pn] = {"desc": desc, "mfr": mfr, "type": "Accessory",
                                  "qty": pn_counter.get(pn, {}).get("qty", 0) + qty}
    for lp in loose_parts:
        pn = lp.get("pn", "")
        if pn not in pn_counter:
            pn_counter[pn] = {"desc": lp.get("desc", ""), "mfr": "", "type": "Loose", "qty": 0}
        pn_counter[pn]["qty"] += lp.get("qty", 0)
    row = 3
    for pn, d in sorted(pn_counter.items()):
        f = _F["alt"] if row % 2 == 0 else _F["white"]
        dc(ws_sum, row, 1, pn, fill=f)
        dc(ws_sum, row, 2, d["desc"], fill=f)
        dc(ws_sum, row, 3, d["mfr"], fill=f)
        dc(ws_sum, row, 4, d["qty"], fill=f, align=C)
        dc(ws_sum, row, 5, d["type"], fill=f)
        row += 1
    ws_sum.freeze_panes = "A3"
    ws_sum.auto_filter.ref = ws_sum.dimensions

    ws_conn = wb.create_sheet("Details")
    ws_conn.sheet_view.showGridLines = False
    headers_conn = [("Position", 16), ("Category", 18), ("Description", 42),
                    ("Item ID", 22), ("Mfr", 16), ("Qty", 10)]
    for ci, (h, w) in enumerate(headers_conn, 1):
        hc(ws_conn, 1, ci, h, w)
    conn_rows = []
    for b in blocks:
        housing_pn = b.pn if b.pn else ""
        block_data = []
        if housing_pn:
            block_data.append(("Housing", f"Body {b.label if b.label else housing_pn}", housing_pn, b.manufacturer, 1))
        if b.tpa:
            block_data.append(("TPA", f"TPA:{b.tpa}", b.tpa, "", 1))
        for key, (desc, mfr, pn, qty) in b.components.items():
            if desc == "ItemA":
                block_data.append(("TypeA", "ItemA", pn, mfr, qty))
            elif desc == "ItemB":
                block_data.append(("TypeB", "ItemB", pn, mfr, qty))
            else:
                block_data.append(("Other", desc, pn, mfr, qty))
        for cat, desc, pn, mfr, qty in block_data:
            conn_rows.append((b.position, cat, desc, pn, mfr, qty))
    for lp in loose_parts:
        conn_rows.append(("", lp.get("type", ""), lp.get("desc", ""), lp.get("pn", ""), "", lp.get("qty", 0)))
    row = 2
    for pos, cat, desc, pn, mfr, qty in conn_rows:
        fill = _F["alt"] if row % 2 == 0 else _F["white"]
        dc(ws_conn, row, 1, pos, fill=fill, bold=True if pos else False)
        dc(ws_conn, row, 2, cat, fill=fill)
        dc(ws_conn, row, 3, desc, fill=fill)
        dc(ws_conn, row, 4, pn, fill=fill)
        dc(ws_conn, row, 5, mfr, fill=fill)
        dc(ws_conn, row, 6, qty, fill=fill, align=C)
        row += 1
    ws_conn.freeze_panes = "A2"
    ws_conn.auto_filter.ref = ws_conn.dimensions
    wb.save(out_path)
    log.info(f"  ✓ saved: {Path(out_path).name}")


def export_summary(all_blocks: Dict[str, List[TableBlock]], loose_parts: Dict[str, List[dict]], out_path: str):
    wb = openpyxl.Workbook()
    global_cnt = {}
    for file_label, blocks in all_blocks.items():
        for b in blocks:
            for key, (desc, mfr, pn, qty) in b.components.items():
                if key == "housing":
                    if pn not in global_cnt:
                        global_cnt[pn] = {"desc": desc, "mfr": mfr, "type": "Housing", "qty": 0, "files": set()}
                    global_cnt[pn]["qty"] += qty
                    global_cnt[pn]["files"].add(file_label)
                elif key == "tpa":
                    if pn not in global_cnt:
                        global_cnt[pn] = {"desc": desc, "mfr": mfr, "type": "TPA", "qty": 0, "files": set()}
                    global_cnt[pn]["qty"] += qty
                    global_cnt[pn]["files"].add(file_label)
                else:
                    if pn not in global_cnt:
                        global_cnt[pn] = {"desc": desc, "mfr": mfr, "type": "Accessory", "qty": 0, "files": set()}
                    global_cnt[pn]["qty"] += qty
                    global_cnt[pn]["files"].add(file_label)
        for lp in loose_parts.get(file_label, []):
            pn = lp.get("pn", "")
            if pn not in global_cnt:
                global_cnt[pn] = {"desc": lp.get("desc", ""), "mfr": "", "type": "Loose", "qty": 0, "files": set()}
            global_cnt[pn]["qty"] += lp.get("qty", 0)
            global_cnt[pn]["files"].add(file_label)

    ws_s = wb.active
    ws_s.title = "Summary"
    ws_s.sheet_view.showGridLines = False
    ws_s.merge_cells("A1:F1")
    ws_s["A1"].value = "Combined Summary"
    ws_s["A1"].font = _FN["title"]
    ws_s["A1"].fill = _F["dark"]
    ws_s["A1"].alignment = C
    ws_s.row_dimensions[1].height = 24
    headers_s = [("Item ID", 22), ("Description", 42), ("Mfr", 16),
                 ("Total Qty", 10), ("Type", 14), ("Files", 50)]
    for ci, (h, w) in enumerate(headers_s, 1):
        hc(ws_s, 2, ci, h, w)
    row = 3
    for pn, d in sorted(global_cnt.items()):
        f = _F["alt"] if row % 2 == 0 else _F["white"]
        dc(ws_s, row, 1, pn, fill=f)
        dc(ws_s, row, 2, d["desc"], fill=f)
        dc(ws_s, row, 3, d["mfr"], fill=f)
        dc(ws_s, row, 4, d["qty"], fill=f, align=C)
        dc(ws_s, row, 5, d["type"], fill=f)
        dc(ws_s, row, 6, ", ".join(sorted(d["files"])), fill=f)
        row += 1
    ws_s.freeze_panes = "A3"
    ws_s.auto_filter.ref = ws_s.dimensions

    ws_h = wb.create_sheet("By File")
    ws_h.sheet_view.showGridLines = False
    headers_h = [("File", 32), ("Item ID", 22), ("Description", 42),
                 ("Mfr", 16), ("Qty", 10), ("Type", 14)]
    for ci, (h, w) in enumerate(headers_h, 1):
        hc(ws_h, 1, ci, h, w)
    cur_row = 2
    for file_label in sorted(all_blocks.keys()):
        local = {}
        for b in all_blocks[file_label]:
            for key, (desc, mfr, pn, qty) in b.components.items():
                if key == "housing":
                    local[pn] = {"desc": desc, "mfr": mfr, "type": "Housing",
                                 "qty": local.get(pn, {}).get("qty", 0) + qty}
                elif key == "tpa":
                    local[pn] = {"desc": desc, "mfr": mfr, "type": "TPA", "qty": local.get(pn, {}).get("qty", 0) + qty}
                else:
                    local[pn] = {"desc": desc, "mfr": mfr, "type": "Accessory",
                                 "qty": local.get(pn, {}).get("qty", 0) + qty}
        for lp in loose_parts.get(file_label, []):
            pn = lp.get("pn", "")
            local[pn] = {"desc": lp.get("desc", ""), "mfr": "", "type": "Loose",
                         "qty": local.get(pn, {}).get("qty", 0) + lp.get("qty", 0)}
        for pn, d in sorted(local.items()):
            f = _F["alt"] if cur_row % 2 == 0 else _F["white"]
            dc(ws_h, cur_row, 1, file_label, fill=f, bold=True)
            dc(ws_h, cur_row, 2, pn, fill=f)
            dc(ws_h, cur_row, 3, d["desc"], fill=f)
            dc(ws_h, cur_row, 4, d["mfr"], fill=f)
            dc(ws_h, cur_row, 5, d["qty"], fill=f, align=C)
            dc(ws_h, cur_row, 6, d["type"], fill=f)
            cur_row += 1
    ws_h.freeze_panes = "A2"
    ws_h.auto_filter.ref = ws_h.dimensions

    wb.save(out_path)
    log.info(f"  ✓ summary report: {Path(out_path).name}")


# ══════════════════════════════════════════════════════════════════════
#  LAUNCH
# ══════════════════════════════════════════════════════════════════════
def process_folder(dxf_dir: str, output_dir: str = "./output"):
    out = Path(output_dir)
    out.mkdir(parents=True, exist_ok=True)
    files_set = set()
    for ext in ("*.dxf", "*.DXF"):
        for f in Path(dxf_dir).glob(ext):
            files_set.add(f.resolve())
    files = sorted(files_set)
    if not files:
        log.warning(f"No DXF files found: {dxf_dir}")
        return
    log.info(f"Found {len(files)} DXF files")
    all_blocks: Dict[str, List[TableBlock]] = {}
    loose_parts: Dict[str, List[dict]] = {}
    ok = fail = 0
    for f in files:
        try:
            p = DXFParser(str(f))
            blocks, loose = p.parse()
            if blocks or loose:
                all_blocks[p.name] = blocks
                loose_parts[p.name] = loose
                export_single(blocks, loose, p.name, str(out / f"{p.name}_parsed_{VERSION}.xlsx"))
                log.info(f"  {p.name}: blocks={len(blocks)}, loose={len(loose)}")
                ok += 1
            else:
                log.warning(f"  {f.name}: no data")
                fail += 1
        except Exception as e:
            log.error(f"  {f.name}: {e}\n{traceback.format_exc()}")
            fail += 1
    log.info(f"\nResult: {ok} ✓, {fail} ✗")
    if all_blocks:
        export_summary(all_blocks, loose_parts, str(out / f"Summary_Parsed_{VERSION}.xlsx"))
        log.info(f"Done → {out.resolve()}")

if __name__ == "__main__":
    input_folder = r"C:\path\to\dxf\files"
    output_folder = r"./output"
    process_folder(input_folder, output_folder)
```
