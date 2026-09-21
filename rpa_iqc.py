"""
IQC RPA — EHLLAPI (no OCR)
Polls iqc_queue from MSSQL, drives DEQ05161 → DEQ05162 → DEQ05171 → DEQ05172

⚠  Needs Python 32-bit — ehlapi32.dll is 32-bit only
   Run: C:\path\to\python32\python.exe rpa_iqc.py

Flow per record:
  1. DEQ05161  — type INVOICE_NO at INVOICE NO field → Enter
  2. DEQ05162  — find row with ITEM_NO → type X → Enter  (SELECT TO ENTRY list)
  3. DEQ05171  — verify, Enter
  4. DEQ05172  — type VISUAL_RESULT, DIM_RESULT → Enter (CHECK)
"""
import os, sys, struct, ctypes, time, logging, json, pyodbc, urllib.request
from ctypes import c_int, byref, create_string_buffer
from dotenv import load_dotenv

_base = os.path.dirname(sys.executable if getattr(sys, "frozen", False) else os.path.abspath(__file__))
load_dotenv(os.path.join(_base, ".env"))

logging.basicConfig(level=logging.DEBUG, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger("rpa-iqc")

# ── Config ────────────────────────────────────────────────────────────────────
DB_SERVER      = os.getenv("DB_SERVER",     "localhost")
DB_NAME        = os.getenv("DB_NAME",       "IQC_DB")
DB_USER        = os.getenv("DB_USER",       "")
DB_PASSWORD    = os.getenv("DB_PASSWORD",   "")
WH_DB_NAME     = os.getenv("WH_DB_NAME",    "Warehouse_F5")
POLL_INTERVAL  = int(os.getenv("POLL_INTERVAL", "5"))
EHLLAPI_SESSION = os.getenv("EHLLAPI_SESSION", "A")
DLL_DIR        = os.getenv("DLL_DIR", r"C:\Program Files (x86)\IBM\Client Access\Emulator")

AS400_USER     = os.getenv("AS400_USER",    "")
AS400_PASS     = os.getenv("AS400_PASS",    "")
PBASMT_USER    = os.getenv("PBASMT_USER",   "")
PBASMT_PASS    = os.getenv("PBASMT_PASS",   "")
AS400_MENU_OPTION = os.getenv("AS400_MENU_OPTION", "")   # e.g. "16-MP-14"
PBASMT_CHECK_PASS   = os.getenv("PBASMT_CHECK_PASS",   "")
PBASMT_SECTION_CODE = os.getenv("PBASMT_SECTION_CODE", "")
SKIP_LOGIN     = os.getenv("SKIP_LOGIN", "false").lower() == "true"

CONFIRM_MODE        = os.getenv("CONFIRM_MODE", "false").lower() == "true"
CONFIRM_HOST        = os.getenv("CONFIRM_HOST", "")          # e.g. "192.168.1.10:3070"
CONFIRM_TIMEOUT_SEC = int(os.getenv("CONFIRM_TIMEOUT_SEC", "120"))

# ponytail: calibrate all _ROW/_COL constants against the real terminal by calling
# read_screen() and counting chars before first production run. Positions are 1-based.

# DEQ05161 — observed from cursor "08/041" after typing "JPD2609-033" (11 chars):
#   INVOICE NO field starts at row 8, col 30  (41 - 11 = 30)
#   P/O NO field is at row 6, col 30 (cursor "06/030" on blank screen)
INV_FIELD_ROW = int(os.getenv("INV_FIELD_ROW", "8"))
INV_FIELD_COL = int(os.getenv("INV_FIELD_COL", "30"))

# List screen — OPT input column; ponytail: verify with read_screen() dump
LIST_OPT_COL  = int(os.getenv("LIST_OPT_COL",  "1"))

# DEQ05172 — ponytail: cursor shows "11/012" on entry, likely VISUAL RESULT field
VISUAL_ROW    = int(os.getenv("VISUAL_ROW",    "11"))
VISUAL_COL    = int(os.getenv("VISUAL_COL",    "24"))
DIM_RESULT_ROW = int(os.getenv("DIM_RESULT_ROW", "13"))
DIM_RESULT_COL = int(os.getenv("DIM_RESULT_COL", "24"))

# ── EHLLAPI constants ────────────────────────────────────────────────────────
CONNECT_PS, DISCONNECT_PS, SEND_KEY = 1, 2, 3
COPY_PS              = 5
WAIT_PS              = 4
QUERY_CURSOR         = 7   # returns 1-based PS position in the length param
QUERY_SESSION_STATUS = 22
# ponytail: SET_CURSOR (func 40) returns rc=7 on this ehlapi32.dll — never use it.
# Navigation via Home+Tab (field_stops/goto_field) is the only working approach.

ENTER_KEY = "@E"   # AID key mnemonics for SEND_KEY
F1_KEY    = "@1"   # CMD1 — BACK / END
F3_KEY    = "@3"
F7_KEY    = "@7"   # CMD.7 — REFER IQC NO.
HOME_KEY  = "@0"   # move cursor to first input field
ERASE_EOF = "@F"   # erase from cursor to end of field
RESET_KEY = "@R"   # reset / clear input-inhibit

# ── EHLLAPI session ──────────────────────────────────────────────────────────
_call_fn      = None
_session_id   = None
_screen_dims  = (24, 80)   # default; queried on connect
_stops_cache: list | None = None   # field_stops() cache; cleared on every Enter


def _load_hllapi():
    if 8 * struct.calcsize("P") != 32:
        sys.exit("ERROR: needs Python 32-bit — ehlapi32.dll is 32-bit. "
                 "Run: <python32>\\python.exe rpa_iqc.py")
    if hasattr(os, "add_dll_directory") and os.path.isdir(DLL_DIR):
        os.add_dll_directory(DLL_DIR)
    for name in ("ehlapi32", "pcshll32"):
        for path in (os.path.join(DLL_DIR, name), name):
            try:
                dll = ctypes.WinDLL(path)
                if hasattr(dll, "hllapi"):
                    log.info("EHLLAPI loaded: %s", path)
                    return dll.hllapi
            except OSError:
                pass
    sys.exit("ERROR: cannot load ehlapi32/pcshll32 — is IBM PCOMM installed?")


def _make_call(hllapi_fn):
    def _call(func, data=b"", length=None):
        if length is None:
            length = len(data)
        buf = create_string_buffer(data if data else b"\x00", max(length, 1))
        f, ln, rc = c_int(func), c_int(length), c_int(0)
        hllapi_fn(byref(f), buf, byref(ln), byref(rc))
        return buf.raw[:ln.value], ln.value, rc.value
    return _call


def init_ehllapi():
    global _call_fn, _session_id, _screen_dims
    if _call_fn is None:
        _call_fn = _make_call(_load_hllapi())
    for s in ([EHLLAPI_SESSION] + [c for c in "ABCDEFGH" if c != EHLLAPI_SESSION]):
        _, _, rc = _call_fn(CONNECT_PS, s.encode(), 1)
        if rc == 0:
            _session_id = s
            data = s.encode() + b" " * 19
            raw, _, rc2 = _call_fn(QUERY_SESSION_STATUS, data, len(data))
            if rc2 == 0 and len(raw) >= 15:
                rows = struct.unpack_from("<H", raw, 11)[0]
                cols = struct.unpack_from("<H", raw, 13)[0]
                if 20 <= rows <= 27 and 80 <= cols <= 132:
                    _screen_dims = (rows, cols)
            log.info("EHLLAPI session=%s dims=%s", s, _screen_dims)
            return
        _call_fn(DISCONNECT_PS, s.encode(), 1)
    raise RuntimeError("cannot connect to any EHLLAPI session — open IBM PCOMM first")


def _c(func, data=b"", length=None):
    """Low-level call with lazy init."""
    if _call_fn is None:
        init_ehllapi()
    return _call_fn(func, data, length)


def read_screen() -> str:
    rows, cols = _screen_dims
    size = rows * cols
    raw, _, rc = _c(COPY_PS, b"\x00" * size, size)
    if rc not in (0, 1):
        log.warning("COPY_PS rc=%d", rc)
    return raw.decode("latin-1", "replace")


def send_key(text: str, _retries: int = 3) -> int:
    """Send text + AID mnemonics. Auto-recovers from rc=4/5 (busy/input-inhibit)."""
    enc = text.encode("latin-1", "replace")
    _, _, rc = _c(SEND_KEY, enc, len(enc))
    if rc == 0:
        return 0
    if rc in (4, 5) and _retries > 0:
        _c(SEND_KEY, RESET_KEY.encode(), len(RESET_KEY))   # clear inhibit
        time.sleep(0.3)
        _c(WAIT_PS, b"", 0)
        time.sleep(0.3)
        return send_key(text, _retries - 1)
    log.warning("SEND_KEY %r rc=%d", text, rc)
    return rc


def query_cursor() -> tuple[int, int]:
    """Return (row, col) 1-based of current host cursor via func 7 (works on this DLL)."""
    if _call_fn is None:
        init_ehllapi()
    buf = create_string_buffer(8)
    f, ln, rc = c_int(QUERY_CURSOR), c_int(0), c_int(0)
    _call_fn(byref(f), buf, byref(ln), byref(rc))
    _, cols = _screen_dims
    p = ln.value or 1
    return (p - 1) // cols + 1, (p - 1) % cols + 1


def field_stops(max_fields: int = 120) -> list[tuple[int, int]]:
    """[(row,col)] of every input field on the current screen via Home+Tab enumeration.
    Cached per screen — cleared on press_enter() when screen transitions."""
    global _stops_cache
    if _stops_cache is not None:
        return _stops_cache
    send_key(HOME_KEY)
    first = query_cursor()
    stops = [first]
    for _ in range(max_fields):
        send_key("@T")
        cur = query_cursor()
        if cur == first:
            break
        stops.append(cur)
    _stops_cache = stops
    log.debug("field_stops: %d fields — %s", len(stops), stops)
    return stops


def goto_field(row: int, col: int) -> tuple[int, int]:
    """Navigate to the input field nearest (row, col) using Home+Tab.
    Returns actual (row, col) landed — may differ if no field exists at exact target."""
    stops = field_stops()
    same_row = [s for s in stops if s[0] == row] or stops
    target = min(same_row, key=lambda s: (abs(s[0] - row), abs(s[1] - col)))
    idx = stops.index(target)
    send_key(HOME_KEY)
    for _ in range(idx):
        send_key("@T")
    return query_cursor()


def type_at(row: int, col: int, value: str):
    """Navigate to field at (row, col) via Home+Tab, erase existing value, type new value."""
    landed = goto_field(row, col)
    log.debug("[TYPE_AT] target=(%d,%d) landed=(%d,%d) value=%r", row, col, *landed, value)
    send_key(ERASE_EOF)
    time.sleep(0.05)
    send_key(str(value))
    time.sleep(0.1)


def press_enter(wait: float = 1.0):
    global _stops_cache
    _stops_cache = None   # screen changes after Enter — next type_at re-enumerates fields
    send_key(ENTER_KEY)
    time.sleep(wait)


def find_row_with(ps: str, keyword: str):
    """Return 1-based row containing keyword, or None."""
    _, cols = _screen_dims
    pos = ps.upper().find(keyword.upper())
    return (pos // cols + 1) if pos >= 0 else None


def wait_screen(keyword: str, timeout: int = 15, label: str = "") -> str:
    """Poll until keyword appears in PS. Returns PS text."""
    for _ in range(timeout * 2):
        ps = read_screen()
        if keyword.upper() in ps.upper():
            return ps
        time.sleep(0.5)
    ps = read_screen()
    raise TimeoutError(f"timeout waiting for {label or keyword!r}")


# ── DB ────────────────────────────────────────────────────────────────────────
_CONN_STR = (
    "DRIVER={ODBC Driver 17 for SQL Server};"
    f"SERVER={DB_SERVER};DATABASE={DB_NAME};"
    f"UID={DB_USER};PWD={DB_PASSWORD}"
)


def _db():
    return pyodbc.connect(_CONN_STR, autocommit=False)


_WH_CONN_STR = (
    "DRIVER={ODBC Driver 17 for SQL Server};"
    f"SERVER={DB_SERVER};DATABASE={WH_DB_NAME};"
    f"UID={DB_USER};PWD={DB_PASSWORD}"
)


def _wh_db():
    return pyodbc.connect(_WH_CONN_STR, autocommit=False)


def update_warehouse_inbound(invoice_no: str):
    """Best-effort: update inbound_task.status in Warehouse_F5. Logs on failure, never raises."""
    try:
        with _wh_db() as conn:
            rows = conn.execute(
                "UPDATE inbound_task SET status='IQC_COMPLETE' WHERE INVOICE_NO=?",
                invoice_no
            ).rowcount
            conn.commit()
        log.info("[WH] inbound_task updated %d row(s) for invoice %s", rows, invoice_no)
    except Exception as e:
        log.warning("[WH] warehouse update failed for invoice %s — %s", invoice_no, e)


def ensure_table():
    with _db() as conn:
        conn.execute("""
            IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='iqc_queue' AND xtype='U')
            CREATE TABLE iqc_queue (
                id                  INT IDENTITY PRIMARY KEY,
                INVOICE_NO          NVARCHAR(50)  NOT NULL,
                ITEM_NO             NVARCHAR(50)  NULL,        -- OSA No. (optional, special request)
                MODEL_NAME          NVARCHAR(100) NULL,
                REV                 NVARCHAR(20)  NULL,
                VISUAL_QTY          INT           NULL,
                VISUAL_GOOD_QTY     INT           NULL,
                VISUAL_NG_QTY       INT           NULL,
                VISUAL_RESULT       NVARCHAR(1)   NULL DEFAULT 'A',
                DIM_QTY             INT           NULL,
                DIM_GOOD_QTY        INT           NULL,
                DIM_NG_QTY          INT           NULL,
                DIM_RESULT          NVARCHAR(1)   NULL DEFAULT 'A',
                SKIP_LOT_NO         NVARCHAR(50)  NULL,
                INSPECTION_TIME     DATETIME2     NULL,
                INSPECTION_OPERATOR NVARCHAR(100) NULL,
                AQL_LEVEL           NVARCHAR(20)  NULL,
                OSA_NO              NVARCHAR(50)  NULL,
                REMARK              NVARCHAR(MAX) NULL,
                screen_text         NVARCHAR(MAX) NULL,
                status              NVARCHAR(30)  NOT NULL DEFAULT 'IQC_WAITING',
                error               NVARCHAR(MAX) NULL,
                created_at          DATETIME2     NOT NULL DEFAULT GETUTCDATE(),
                updated_at          DATETIME2     NOT NULL DEFAULT GETUTCDATE()
            )
        """)
        conn.execute("""
            IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='iqc_result' AND xtype='U')
            CREATE TABLE iqc_result (
                id                  INT IDENTITY PRIMARY KEY,
                queue_id            INT           NOT NULL,
                INVOICE_NO          NVARCHAR(50)  NOT NULL,
                ITEM_NO             NVARCHAR(50)  NULL,
                MODEL_NAME          NVARCHAR(100) NULL,
                REV                 NVARCHAR(20)  NULL,
                VISUAL_QTY          INT           NULL,
                VISUAL_GOOD_QTY     INT           NULL,
                VISUAL_NG_QTY       INT           NULL,
                VISUAL_RESULT       NVARCHAR(1)   NULL,
                DIM_QTY             INT           NULL,
                DIM_GOOD_QTY        INT           NULL,
                DIM_NG_QTY          INT           NULL,
                DIM_RESULT          NVARCHAR(1)   NULL,
                SKIP_LOT_NO         NVARCHAR(50)  NULL,
                INSPECTION_TIME     DATETIME2     NULL,
                INSPECTION_OPERATOR NVARCHAR(100) NULL,
                AQL_LEVEL           NVARCHAR(20)  NULL,
                OSA_NO              NVARCHAR(50)  NULL,
                REMARK              NVARCHAR(MAX) NULL,
                completed_at        DATETIME2     NOT NULL DEFAULT GETUTCDATE()
            )
        """)
        conn.commit()
    log.info("DB tables ready (IQC_DB)")


def fetch_pending():
    with _db() as conn:
        rows = conn.execute(
            "SELECT id, INVOICE_NO, ITEM_NO, MODEL_NAME, REV, "
            "       VISUAL_QTY, VISUAL_GOOD_QTY, VISUAL_NG_QTY, VISUAL_RESULT, "
            "       DIM_QTY, DIM_GOOD_QTY, DIM_NG_QTY, DIM_RESULT, "
            "       SKIP_LOT_NO, INSPECTION_TIME, INSPECTION_OPERATOR, AQL_LEVEL, OSA_NO, REMARK "
            "FROM iqc_queue WHERE status='IQC_WAITING' ORDER BY created_at"
        ).fetchall()
    return [{
        "id":                 r[0],
        "INVOICE_NO":         r[1] or "",
        "ITEM_NO":            r[2] or "",
        "MODEL_NAME":         r[3] or "",
        "REV":                r[4] or "",
        "VISUAL_QTY":         r[5],
        "VISUAL_GOOD_QTY":    r[6],
        "VISUAL_NG_QTY":      r[7],
        "VISUAL_RESULT":      r[8] or "A",
        "DIM_QTY":            r[9],
        "DIM_GOOD_QTY":       r[10],
        "DIM_NG_QTY":         r[11],
        "DIM_RESULT":         r[12] or "A",
        "SKIP_LOT_NO":        r[13] or "",
        "INSPECTION_TIME":    r[14],
        "INSPECTION_OPERATOR":r[15] or "",
        "AQL_LEVEL":          r[16] or "",
        "OSA_NO":             r[17] or "",
        "REMARK":             r[18] or "",
    } for r in rows]


def mark_processing(row_id: int):
    with _db() as conn:
        conn.execute(
            "UPDATE iqc_queue SET status='PROCESSING', updated_at=GETUTCDATE() WHERE id=?",
            row_id
        )
        conn.commit()


def insert_iqc_result(record: dict):
    with _db() as conn:
        conn.execute(
            "INSERT INTO iqc_result "
            "(queue_id, INVOICE_NO, ITEM_NO, MODEL_NAME, REV, "
            " VISUAL_QTY, VISUAL_GOOD_QTY, VISUAL_NG_QTY, VISUAL_RESULT, "
            " DIM_QTY, DIM_GOOD_QTY, DIM_NG_QTY, DIM_RESULT, "
            " SKIP_LOT_NO, INSPECTION_TIME, INSPECTION_OPERATOR, AQL_LEVEL, OSA_NO, REMARK) "
            "VALUES (?,?,?,?,?, ?,?,?,?, ?,?,?,?, ?,?,?,?,?,?)",
            record["id"], record["INVOICE_NO"], record["ITEM_NO"] or None,
            record["MODEL_NAME"] or None, record["REV"] or None,
            record["VISUAL_QTY"], record["VISUAL_GOOD_QTY"], record["VISUAL_NG_QTY"],
            record["VISUAL_RESULT"],
            record["DIM_QTY"], record["DIM_GOOD_QTY"], record["DIM_NG_QTY"],
            record["DIM_RESULT"],
            record["SKIP_LOT_NO"] or None, record["INSPECTION_TIME"],
            record["INSPECTION_OPERATOR"] or None, record["AQL_LEVEL"] or None,
            record["OSA_NO"] or None, record["REMARK"] or None,
        )
        conn.commit()


def mark_done(row_id: int, status: str, error: str = None, invoice_no: str = None):
    with _db() as conn:
        conn.execute(
            "UPDATE iqc_queue SET status=?, error=?, updated_at=GETUTCDATE() WHERE id=?",
            status, error, row_id
        )
        if status == "IQC_COMPLETE" and invoice_no:
            # Mark all remaining IQC_WAITING rows for the same invoice complete
            conn.execute(
                "UPDATE iqc_queue SET status='IQC_COMPLETE', updated_at=GETUTCDATE() "
                "WHERE INVOICE_NO=? AND status='IQC_WAITING'",
                invoice_no
            )
        conn.commit()


def mark_awaiting_confirm(row_id: int, screen_text: str):
    with _db() as conn:
        conn.execute(
            "UPDATE iqc_queue SET status='AWAITING_CONFIRM', screen_text=?, updated_at=GETUTCDATE() WHERE id=?",
            screen_text, row_id
        )
        conn.commit()
    log.info("AWAITING_CONFIRM id=%d", row_id)


def get_queue_status(row_id: int) -> str:
    with _db() as conn:
        row = conn.execute("SELECT status FROM iqc_queue WHERE id=?", row_id).fetchone()
    return row[0] if row else "FAILED"


def send_confirm_request(record: dict, screen_text: str):
    """POST record + screen text to the IQC web site for operator confirmation modal."""
    if not CONFIRM_HOST:
        log.warning("CONFIRM_HOST not set — web notification skipped")
        return
    url = f"http://{CONFIRM_HOST}/api/iqc/confirm-request"
    body = json.dumps({
        "id":            record["id"],
        "invoice_no":    record["INVOICE_NO"],
        "item_no":       record["ITEM_NO"],
        "visual_result": record["VISUAL_RESULT"],
        "dim_result":    record["DIM_RESULT"],
        "screen_text":   screen_text,
    }).encode()
    req = urllib.request.Request(url, data=body, headers={"Content-Type": "application/json"})
    try:
        with urllib.request.urlopen(req, timeout=10) as resp:
            log.info("confirm request sent → HTTP %d", resp.status)
    except Exception as e:
        log.warning("confirm request failed: %s — continuing with DB poll", e)


# ── Login screens (same PBASMT flow as original) ────────────────────────────

def ensure_signed_in():
    if SKIP_LOGIN:
        log.info("SKIP_LOGIN=true — assuming already at DEQ05161")
        return
    ps = read_screen()
    if "Sign On" in ps or "SIGN ON" in ps.upper():
        _do_signon()
    ps = read_screen()
    if "USER ID" in ps.upper():
        _do_pbasmt_login()
    if AS400_MENU_OPTION:
        _navigate_menu()
    ps = read_screen()
    if "PASSWORD CHECK" in ps.upper():
        _do_password_check()


def _do_signon():
    if not AS400_USER or not AS400_PASS:
        log.warning("AS400_USER/PASS not set — skipping signon")
        return
    log.info("[SIGNON] logging in as %s", AS400_USER)
    send_key(AS400_USER + "@T" + AS400_PASS + ENTER_KEY)
    time.sleep(1.5)
    press_enter(0.5)


def _do_pbasmt_login():
    if not PBASMT_USER or not PBASMT_PASS:
        log.warning("PBASMT_USER/PASS not set — skipping app login")
        return
    log.info("[PBASMT] app login as %s", PBASMT_USER)
    send_key(PBASMT_USER + "@T" + PBASMT_PASS + ENTER_KEY)
    time.sleep(1.5)
    press_enter(0.5)


def _navigate_menu():
    for step in AS400_MENU_OPTION.split("-"):
        log.info("[MENU] %s", step)
        send_key(step + ENTER_KEY)
        time.sleep(0.5)
    time.sleep(1.0)


def _do_password_check():
    if not PBASMT_CHECK_PASS:
        log.warning("PBASMT_CHECK_PASS not set — skipping password check screen")
        return
    log.info("[PWCHECK] entering password")
    val = PBASMT_CHECK_PASS
    if PBASMT_SECTION_CODE:
        val += "@T" + PBASMT_SECTION_CODE
    send_key(val + ENTER_KEY)
    time.sleep(1.5)


# ── IQC Screen handlers ──────────────────────────────────────────────────────

def screen_deq05161_input_invoice(invoice_no: str):
    """DEQ05161 — type INVOICE_NO at the INVOICE NO. field → Enter."""
    wait_screen("MATERIAL FOR IQC", label="DEQ05161")
    log.info("[DEQ05161] invoice=%s", invoice_no)
    type_at(INV_FIELD_ROW, INV_FIELD_COL, invoice_no)
    press_enter(1.0)

    # Dismiss "TIME STARTING AT" timestamp popup if it appears
    ps = read_screen()
    if "TIME STARTING AT" in ps.upper():
        log.info("[DEQ05161] dismissing TIME STARTING AT popup")
        press_enter(1.0)


def screen_deq05162_select_item(item_no: str, invoice_no: str):
    """DEQ05162 (SELECT TO ENTRY list) — find row with both ITEM_NO and INVOICE_NO, type X → Enter.
    Same ITEM_NO can appear on multiple rows (different invoices) so both must match.
    """
    ps = wait_screen("SELECT TO ENTRY", label="DEQ05162")
    _, cols = _screen_dims
    rows_text = [ps[i * cols:(i + 1) * cols] for i in range(_screen_dims[0])]

    # Find rows containing item_no, then pick the one that also has invoice_no
    inv_suffix = invoice_no.upper().split("-")[-1]  # match trailing part e.g. "0397" from "NMB-0397"
    matches = [
        i + 1 for i, row in enumerate(rows_text)
        if item_no.upper() in row.upper() and inv_suffix in row.upper()
    ]
    if not matches:
        # Fallback: item_no only (warn — may pick wrong row if duplicates)
        matches = [i + 1 for i, row in enumerate(rows_text) if item_no.upper() in row.upper()]
        if not matches:
            log.error("[DEQ05162] screen dump:\n%s", "\n".join(rows_text))
            raise ValueError(f"ITEM_NO {item_no!r} / INVOICE_NO {invoice_no!r} not found on list screen")
        log.warning("[DEQ05162] invoice suffix %r not found on same row — using item_no match only (row %d)", inv_suffix, matches[0])
    if len(matches) > 1:
        log.warning("[DEQ05162] %d rows match item+invoice — picking first (row %d)", len(matches), matches[0])

    row_num = matches[0]
    log.info("[DEQ05162] item %s invoice %s at row %d — selecting X", item_no, invoice_no, row_num)
    type_at(row_num, LIST_OPT_COL, "X")
    press_enter(1.5)


def screen_deq05171_verify(record: dict):
    """DEQ05171 — verify screen, press Enter to proceed to DEQ05172."""
    ps = wait_screen("INSP.DATE", label="DEQ05171")
    item = record["ITEM_NO"]
    if item and item.upper() not in ps.upper():
        log.warning("[DEQ05171] ITEM_NO %s not visible — proceeding anyway", item)
    log.info("[DEQ05171] verified — Enter")
    press_enter(1.0)


def screen_deq05172_enter_result(record: dict):
    """DEQ05172 — type VISUAL_RESULT and DIM_RESULT.
    CONFIRM_MODE: send modal to web site, poll DB for operator decision, then Enter.
    """
    wait_screen("VISUAL RESULT", label="DEQ05172")
    log.info("[DEQ05172] visual=%s dim=%s", record["VISUAL_RESULT"], record["DIM_RESULT"])
    type_at(VISUAL_ROW,     VISUAL_COL,     record["VISUAL_RESULT"])
    type_at(DIM_RESULT_ROW, DIM_RESULT_COL, record["DIM_RESULT"])

    if not CONFIRM_MODE:
        press_enter(1.5)
        return

    # Read filled screen, notify web site, wait for operator
    screen_text = read_screen()
    mark_awaiting_confirm(record["id"], screen_text)   # DB first — web polls after
    send_confirm_request(record, screen_text)

    deadline = time.time() + CONFIRM_TIMEOUT_SEC
    log.info("[DEQ05172] CONFIRM_MODE — waiting for operator (timeout %ds)", CONFIRM_TIMEOUT_SEC)
    while time.time() < deadline:
        status = get_queue_status(record["id"])
        if status == "CONFIRMED":
            press_enter(1.5)
            log.info("[DEQ05172] operator confirmed → Enter (CHECK)")
            return
        if status == "REJECTED":
            send_key(F1_KEY)   # CMD1 - BACK
            time.sleep(1.0)
            raise ValueError("operator rejected — record stays FAILED")
        time.sleep(3)

    send_key(F1_KEY)
    time.sleep(1.0)
    raise TimeoutError(f"confirm timeout after {CONFIRM_TIMEOUT_SEC}s — no operator response")


def input_to_as400(record: dict):
    """Full 4-screen flow for one IQC record."""
    if not record["INVOICE_NO"]:
        raise ValueError("record missing INVOICE_NO")
    if not record["ITEM_NO"]:
        raise ValueError("record missing ITEM_NO")
    screen_deq05161_input_invoice(record["INVOICE_NO"])
    screen_deq05162_select_item(record["ITEM_NO"], record["INVOICE_NO"])
    screen_deq05171_verify(record)
    screen_deq05172_enter_result(record)
    log.info("done: invoice=%s item=%s", record["INVOICE_NO"], record["ITEM_NO"])


# ── Main loop ────────────────────────────────────────────────────────────────

def process_record(record: dict):
    mark_processing(record["id"])
    try:
        input_to_as400(record)
        insert_iqc_result(record)
        mark_done(record["id"], "IQC_COMPLETE", invoice_no=record["INVOICE_NO"])
        update_warehouse_inbound(record["INVOICE_NO"])
        log.info("IQC_COMPLETE: %s / %s", record["INVOICE_NO"], record["ITEM_NO"])
    except Exception as e:
        log.error("ERROR: %s / %s — %s", record["INVOICE_NO"], record["ITEM_NO"], e)
        mark_done(record["id"], "FAILED", str(e))
        ctypes.windll.user32.MessageBoxW(
            0,
            f"INVOICE: {record['INVOICE_NO']}  ITEM: {record['ITEM_NO']}\n\n{e}\n\nRecord marked FAILED.",
            "IQC RPA — Error",
            0x10,
        )


def poll():
    pending = fetch_pending()
    if not pending:
        log.debug("no pending records")
        return
    log.info("%d pending", len(pending))
    for record in pending:
        process_record(record)


def main():
    init_ehllapi()
    ensure_table()
    ensure_signed_in()
    log.info("polling every %ds — Ctrl+C to stop", POLL_INTERVAL)
    while True:
        try:
            poll()
        except KeyboardInterrupt:
            log.info("stopped")
            break
        except Exception as e:
            log.error("poll error: %s", e)
        time.sleep(POLL_INTERVAL)


if __name__ == "__main__":
    main()
