"""
Calibration helper — run with Python 32-bit while DEQ05172 is on screen.

Usage:
  python32 _calibrate.py           # dump full screen with row/col grid
  python32 _calibrate.py 12 15     # test-click row 12 col 15 (no typing)
  python32 _calibrate.py 12 15 A   # move cursor to (12,15) and type "A"

Use the grid dump to read exact row/col for VISUAL RESULT and DIMENSION RESULT,
then update .env (VISUAL_ROW, VISUAL_COL, DIM_RESULT_ROW, DIM_RESULT_COL).
"""
import sys, os, struct, ctypes
from ctypes import c_int, byref, create_string_buffer

DLL_DIR = os.getenv("DLL_DIR", r"C:\Program Files (x86)\IBM\Client Access\Emulator")

CONNECT_PS, DISCONNECT_PS = 1, 2
SEND_KEY, COPY_PS = 3, 5
QUERY_SESSION_STATUS = 22
SET_CURSOR = 40

if 8 * struct.calcsize("P") != 32:
    sys.exit("Need Python 32-bit")

if hasattr(os, "add_dll_directory") and os.path.isdir(DLL_DIR):
    os.add_dll_directory(DLL_DIR)

hllapi = None
for name in ("ehlapi32", "pcshll32"):
    for path in (os.path.join(DLL_DIR, name), name):
        try:
            dll = ctypes.WinDLL(path)
            if hasattr(dll, "hllapi"):
                hllapi = dll.hllapi
                break
        except OSError:
            pass
    if hllapi:
        break
if not hllapi:
    sys.exit("Cannot load ehlapi32/pcshll32")


def call(func, data=b"", length=None):
    if length is None: length = len(data)
    buf = create_string_buffer(data if data else b"\x00", max(length, 1))
    f, ln, rc = c_int(func), c_int(length), c_int(0)
    hllapi(byref(f), buf, byref(ln), byref(rc))
    return buf.raw[:ln.value], ln.value, rc.value


# Connect
for s in "ABCDEFGH":
    _, _, rc = call(CONNECT_PS, s.encode(), 1)
    if rc == 0:
        session = s
        break
    call(DISCONNECT_PS, s.encode(), 1)
else:
    sys.exit("No session found")

# Query dims
data = session.encode() + b" " * 19
raw, _, _ = call(QUERY_SESSION_STATUS, data, len(data))
rows = struct.unpack_from("<H", raw, 11)[0] if len(raw) >= 15 else 24
cols = struct.unpack_from("<H", raw, 13)[0] if len(raw) >= 15 else 80
if not (20 <= rows <= 27 and 80 <= cols <= 132):
    rows, cols = 24, 80
print(f"Session: {session}  Dims: {rows}x{cols}\n")

# Read screen
raw, _, _ = call(COPY_PS, b"\x00" * rows * cols, rows * cols)
ps = raw.decode("latin-1", "replace")

argv = sys.argv[1:]

if not argv:
    # Dump with row/col ruler
    ruler_tens = "          " + "".join(str(i // 10) if i % 10 == 0 else " " for i in range(1, cols + 1))
    ruler_ones = "     " + "".join(str(i % 10) for i in range(1, cols + 1))
    print(ruler_tens)
    print(ruler_ones)
    print("     " + "-" * cols)
    for r in range(rows):
        line = ps[r * cols:(r + 1) * cols]
        print(f"{r+1:3d}  |{line}|")
    print("\nLook for VISUAL RESULT and DIMENSION RESULT rows above.")
    print("Count the column where the input field (blank space after label) starts.")
    print("Then run:  python32 _calibrate.py <row> <col> [value]")

elif len(argv) >= 2:
    row, col = int(argv[0]), int(argv[1])
    pos = (row - 1) * cols + col
    _, _, rc = call(SET_CURSOR, b"", pos)
    print(f"SET_CURSOR ({row},{col}) -> rc={rc}")
    if len(argv) >= 3:
        text = argv[2]
        enc = text.encode("latin-1", "replace")
        _, _, rc = call(SEND_KEY, enc, len(enc))
        print(f"SEND_KEY {text!r} -> rc={rc}")
    # Re-read to show cursor context
    raw, _, _ = call(COPY_PS, b"\x00" * rows * cols, rows * cols)
    ps2 = raw.decode("latin-1", "replace")
    print(f"\nRow {row-1}: {ps2[(row-2)*cols:(row-1)*cols]!r}")
    print(f"Row {row}:   {ps2[(row-1)*cols:row*cols]!r}  ← cursor here col {col}")
    print(f"Row {row+1}: {ps2[row*cols:(row+1)*cols]!r}")

call(DISCONNECT_PS, session.encode(), 1)
