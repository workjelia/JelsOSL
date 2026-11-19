"""
Noon Report Converter (versi Streamlit)
Jalankan:  pip install streamlit openpyxl pandas
           streamlit run noon_report_converter.py
"""
import hashlib
import io
import math
import re
from datetime import date, datetime, time, timedelta

import pandas as pd
import streamlit as st
from openpyxl import load_workbook

# ----------------------------------------------------------------------------
# PARSER (port 1:1 dari versi website)
# ----------------------------------------------------------------------------
GROUPS = [
    ["NAME OF MASTER"], ["SATELLITE EMAIL"], ["NOON POSITION REPORT OF"], ["VOYAGE NUMBER"], ["DATE"],
    ["FROM"], ["TO"], ["REMARKS"], ["LATITUDE"], ["LONGITUDE"], ["COURSE"], ["DISTANCE TO RUN"],
    ["TOTAL DISTANCE TO RUN"], ["STEAMING TIME"], ["TOTAL STEAMING TIME"], ["AVERAGE SPEED"],
    ["GENERAL AVERAGE SPEED"], ["DISTANCE TO GO"], ["ETA DATE"], ["BY SPEED"], ["RPM (REV PER MINUTE)"],
    ["WEATHER CONDITION"],
    ["CURRENT ROB (REMAIN ON BOARD)", ["BUNKER", "LUBRICATING OIL", "OILY WATER/SLOP TANK", "FRESH WATER"]],
    ["SUPPLY ACTUAL RECEIVED (IN LAST 24 HOUR)", ["BUNKER", "LUBRICATING OIL", "FRESH WATER"]],
    ["CURRENT CONSUMPTION RATE", ["BUNKER", "LUBRICATING OIL", "FRESH WATER"]],
    ["CARGO NAME", ["A", "B", "C", "D"]],
    ["CARGO CURRENT ROB", ["A", "B", "C", "D"]],
    ["PUMP RATE", ["A", "B", "C", "D"]],
    ["EST COMPLETED DATE", ["A", "B", "C", "D"]],
    ["EST COMPLETED HOUR", ["A", "B", "C", "D"]],
]
COLS = []
for _g in GROUPS:
    if len(_g) > 1:
        for _s in _g[1]:
            COLS.append({"g": _g[0], "s": _s})
    else:
        COLS.append({"g": _g[0], "s": None})

DATE_IDX = {4, 18, 44, 45, 46, 47}  # E, S, AS..AV
TIME_IDX = {48, 49, 50, 51}         # AW..AZ


def col_name(c):
    return f'{c["g"]} - {c["s"]}' if c["s"] else c["g"]


COLNAMES = [col_name(c) for c in COLS]


def norm(s):
    return re.sub(r"\s+", " ", "" if s is None else str(s)).strip().upper()


def parse_input(get):
    """get(r, c) -> nilai sel (1-based) atau None. Return (vals, notes, hemi) atau None."""
    L0 = 0
    for r in range(1, 16):
        if L0:
            break
        for c in range(1, 7):
            if norm(get(r, c)) == "NAME OF MASTER":
                L0 = c
                break
    if not L0:
        return None
    V0 = L0 + 2
    notes, rows, bunker = [], {}, []
    for r in range(1, 161):
        t = norm(get(r, L0))
        if not t:
            continue
        if t not in rows:
            rows[t] = r
        if t == "BUNKER":
            bunker.append(r)
    if not rows.get("VOYAGE NUMBER"):
        return None

    def find(*ls):
        for l in ls:
            if l in rows:
                return rows[l]
        for k in rows:
            if any(l in k for l in ls):
                return rows[k]
        return 0

    out = [None] * len(COLS)

    def put(i, r, c):
        if r:
            out[i] = get(r, c)
        else:
            notes.append("Label tidak ditemukan: " + COLS[i]["g"] + (" / " + COLS[i]["s"] if COLS[i]["s"] else ""))

    put(0, find("NAME OF MASTER"), V0)
    put(1, find("EMAIL"), V0)
    rc = find("NOON POSITION REPORT OF")
    if rc:
        v = get(rc, L0 + 3)
        out[2] = v if v is not None else get(rc, V0)
    else:
        notes.append("Label tidak ditemukan: NOON POSITION REPORT OF")
    rv = find("VOYAGE NUMBER")
    put(3, rv, V0)
    put(4, rv, L0 + 5)
    rf = find("FROM")
    put(5, rf, V0)
    put(6, rf, L0 + 5)
    put(7, find("CURRENT LOCATION", "REMARK") or (rf + 1 if rf and not norm(get(rf + 1, L0)) else 0), V0)
    for i, l in enumerate(["LATITUDE", "LONGITUDE", "COURSE", "DISTANCE TO RUN", "TOTAL DISTANCE TO RUN",
                           "STEAMING TIME", "TOTAL STEAMING TIME", "AVERAGE SPEED", "GENERAL AVERAGE SPEED",
                           "DISTANCE TO GO", "ETA DATE", "BY SPEED"]):
        put(8 + i, rows.get(l), V0)
    put(20, find("RPM"), V0)
    put(21, rows.get("WEATHER CONDITION"), V0)
    # petunjuk hemisphere (kolom satuan)
    hemi = [
        norm(get(rows["LATITUDE"], V0 + 8)) if rows.get("LATITUDE") else "",
        norm(get(rows["LONGITUDE"], V0 + 8)) if rows.get("LONGITUDE") else "",
    ]
    # section dibagi oleh 3 baris BUNKER (ROB, SUPPLY, CONSUMPTION)
    end_all = rows.get("CARGO") or rows.get("NOON CARGO OPERATION REPORT") or ((bunker[2] if len(bunker) > 2 else 0) + 10)
    b = lambda i: bunker[i] if i < len(bunker) else None
    sec = [[b(0), b(1)], [b(1), b(2)], [b(2), end_all]]
    LB = {"LO": "LUBRICATING OIL", "FW": "FRESH WATER", "OW/ST": "OILY WATER/SLOP TANK"}

    def code(s, cd, lbl):
        if not s[0]:
            return None
        if s[1] is None:
            return 0
        for r in range(s[0], s[1]):
            if norm(get(r, L0 + 1)) == cd:
                return r
        if lbl:
            for r in range(s[0], s[1]):
                if norm(get(r, L0)) == lbl:
                    return r
        return 0

    for st_i, s, cs in [(22, sec[0], ["HSD", "LO", "OW/ST", "FW"]),
                        (26, sec[1], ["HSD", "LO", "FW"]),
                        (29, sec[2], ["HSD", "LO", "FW"])]:
        for k, cd in enumerate(cs):
            r = code(s, cd, LB.get(cd))
            if r:
                out[st_i + k] = get(r, V0)
            elif cd != "OW/ST":
                notes.append("Baris tidak ditemukan: " + COLS[st_i + k]["g"] + " / " + COLS[st_i + k]["s"])
    # cargo
    cr = rows.get("CARGO")
    for k, g in enumerate("ABCD"):
        r = rows.get("GRADE " + g)
        if not r or not cr:
            notes.append("Baris tidak ditemukan: GRADE " + g)
            continue
        for j, o in enumerate([2, 4, 6, 9, 10]):
            out[32 + j * 4 + k] = get(r, L0 + o)
    return out, notes, hemi


def fmt(i, v, zero):
    if v is None or (isinstance(v, str) and not v.strip()):
        return 0 if (zero and i not in DATE_IDX and i not in TIME_IDX) else ""
    if isinstance(v, str):
        return v.strip()
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        if i in DATE_IDX:
            ms = math.floor((v - 25569) * 864e5 + 0.5)
            d = datetime(1970, 1, 1) + timedelta(milliseconds=ms)
            return f"{d.day:02d}/{d.month:02d}/{d.year}"
        if i in TIME_IDX:
            m = math.floor((v % 1) * 1440 + 0.5)
            return f"{m // 60}:{m % 60:02d}"
    return v


def display(vals, hemi, zero, add_hemi):
    res = []
    for i, v in enumerate(vals):
        f = fmt(i, v, zero)
        if add_hemi and i in (8, 9) and isinstance(f, str) and f and not re.search(r"[NSEW]\s*$", f, re.I):
            u = hemi[i - 8]
            if re.fullmatch(r"[NSEW]", u):
                f += " " + u
        res.append(f)
    return res


# ----------------------------------------------------------------------------
# BACA EXCEL
# ----------------------------------------------------------------------------
_EPOCH = datetime(1899, 12, 30)


def to_serial(v):
    """Samakan dengan nilai mentah SheetJS: tanggal -> serial, jam -> pecahan hari."""
    if isinstance(v, datetime):
        return (v - _EPOCH).total_seconds() / 86400
    if isinstance(v, date):
        return (datetime(v.year, v.month, v.day) - _EPOCH).total_seconds() / 86400
    if isinstance(v, time):
        return (v.hour * 3600 + v.minute * 60 + v.second + v.microsecond / 1e6) / 86400
    if isinstance(v, timedelta):
        return v.total_seconds() / 86400
    return v


@st.cache_data(show_spinner="Membaca file...")
def read_workbook(data: bytes):
    wb = load_workbook(io.BytesIO(data), data_only=True)
    sheets = []
    for n in wb.sheetnames:
        ws = wb[n]
        grid = {}
        if hasattr(ws, "iter_rows"):
            for row in ws.iter_rows(min_row=1, max_row=200, max_col=40):
                for c in row:
                    if c.value is not None:
                        grid[(c.row, c.column)] = to_serial(c.value)
        res = parse_input(lambda r, c, g=grid: g.get((r, c)))
        sheets.append({"n": n, "ok": res is not None,
                       "vals": res[0] if res else None,
                       "notes": res[1] if res else [],
                       "hemi": res[2] if res else []})
    return sheets


def default_on(sheets):
    inp = next((s for s in sheets if norm(s["n"]) == "INPUT" and s["ok"]), None)
    if inp:
        return [s is inp for s in sheets]
    return [s["ok"] and not re.search("CONVERT", s["n"], re.I) for s in sheets]


# ----------------------------------------------------------------------------
# CALLBACK TOMBOL
# ----------------------------------------------------------------------------
def cb_all(key, oks):
    for si, ok in enumerate(oks):
        st.session_state[f"cb_{key}_{si}"] = ok


def cb_none(key, oks):
    for si, _ in enumerate(oks):
        st.session_state[f"cb_{key}_{si}"] = False


def cb_match(key, names, oks):
    q = norm(st.session_state.get(f"q_{key}", ""))
    for si, (n, ok) in enumerate(zip(names, oks)):
        st.session_state[f"cb_{key}_{si}"] = bool(ok and q in norm(n))


def cb_remove(key):
    st.session_state.removed.add(key)


def cb_clear():
    st.session_state.removed = set()
    st.session_state.uploader_n += 1


def tsv(rows, header):
    lines = []
    if header:
        lines.append("\t".join(COLNAMES))
    for r in rows:
        lines.append("\t".join(re.sub(r"[\t\r\n]+", " ", str(v)) for v in r))
    return "\n".join(lines)


# ----------------------------------------------------------------------------
# UI
# ----------------------------------------------------------------------------
def render():
    """Dipanggil dari app.py di dalam tab Converter Noon Report."""
    ss = st.session_state
    ss.setdefault("removed", set())
    ss.setdefault("uploader_n", 0)
    # ✕ pada baris tabel -> hilangkan centang tab (diproses sebelum widget dibuat)
    for k in ss.pop("_uncheck", []):
        ss[k] = False

    st.header("🚢 Converter Noon Report")
    st.write("Upload file Excel, lalu pilih tab mana saja yang mau dikonversi.")

    uploaded = st.file_uploader("",
                                type=["xlsx", "xlsm"], accept_multiple_files=True,
                                key=f"uploader_{ss.uploader_n}",
                                help="Satu tab = satu baris. Layout template boleh berbeda antar kapal.")

    # ---- daftar file & pemilihan tab
    files, seen = [], {}
    for f in uploaded or []:
        data = f.getvalue()
        base = f"{f.name}_{hashlib.md5(data).hexdigest()[:8]}"
        seen[base] = seen.get(base, 0) + 1
        key = f"{base}_{seen[base]}"
        if key in ss.removed:
            continue
        try:
            sheets = read_workbook(data)
        except Exception as e:  # noqa: BLE001
            st.warning(f"{f.name}: gagal dibaca ({e})")
            continue
        if f"init_{key}" not in ss:
            for si, on in enumerate(default_on(sheets)):
                ss[f"cb_{key}_{si}"] = on
            ss[f"init_{key}"] = True
        files.append({"key": key, "name": f.name, "sheets": sheets})

    for F in files:
        key, sheets = F["key"], F["sheets"]
        names = [s["n"] for s in sheets]
        oks = [s["ok"] for s in sheets]
        n_on = sum(1 for si, ok in enumerate(oks) if ok and ss.get(f"cb_{key}_{si}"))
        with st.container(border=True):
            st.markdown(f"**{F['name']}** — {n_on} dari {len(sheets)} tab dipilih")
            c0, c1, c2, c3, c4 = st.columns([3, 1.3, 1, 1, 1])
            c0.text_input("filter nama tab", key=f"q_{key}", placeholder="filter nama tab, mis. JUNE",
                          label_visibility="collapsed")
            c1.button("Pilih yang cocok", key=f"m_{key}", on_click=cb_match, args=(key, names, oks))
            c2.button("Pilih semua", key=f"a_{key}", on_click=cb_all, args=(key, oks))
            c3.button("Kosongkan", key=f"n_{key}", on_click=cb_none, args=(key, oks))
            c4.button("Hapus file", key=f"r_{key}", on_click=cb_remove, args=(key,))
            ncol = 5
            cols = st.columns(ncol)
            for si, s in enumerate(sheets):
                if not s["ok"]:
                    ss[f"cb_{key}_{si}"] = False
                cols[si % ncol].checkbox(s["n"], key=f"cb_{key}_{si}", disabled=not s["ok"],
                                         help=None if s["ok"] else "Layout tidak dikenali (tidak ada label NAME OF MASTER)")

    # ---- opsi
    o1, o2, o3, o4 = st.columns([1, 1.4, 1.6, 1])
    opt_head = o1.checkbox("Sertakan header", value=False)
    opt_zero = o2.checkbox("NA jadi 0", value=False)
    opt_hemi = o3.checkbox("Tambahkan N/S/E/W (Latitude/Longitude)", value=False)
    o4.button("Reset", key="clear_all", on_click=cb_clear)

    # ---- baris hasil
    rows = []
    for F in files:
        for si, s in enumerate(F["sheets"]):
            if s["ok"] and ss.get(f"cb_{F['key']}_{si}"):
                rows.append({"cb": f"cb_{F['key']}_{si}", "name": F["name"], "sheet": s["n"],
                             "disp": display(s["vals"], s["hemi"], opt_zero, opt_hemi), "notes": s["notes"]})

    if not rows:
        st.info("Belum ada tab yang dipilih." if files else "Belum ada file.")
        return

    df = pd.DataFrame([[False, f"{r['sheet']}  ({r['name']})"] + [str(v) for v in r["disp"]] for r in rows],
                      columns=["✕", "File / Tab"] + COLNAMES)
    sig = hashlib.md5(repr(([r["cb"] for r in rows], opt_zero, opt_hemi)).encode()).hexdigest()[:10]
    edited = st.data_editor(
        df, key=f"editor_{sig}", hide_index=True, use_container_width=True, num_rows="fixed",
        disabled=["File / Tab"],
        column_config={"✕": st.column_config.CheckboxColumn("✕", help="Lepas tab ini dari pilihan", width="small")},
    )

    removed = [rows[i]["cb"] for i in range(len(rows)) if bool(edited.iloc[i]["✕"])]
    if removed:
        ss["_uncheck"] = removed
        st.rerun()

    # ---- catatan
    bad = [r for r in rows if r["notes"]]
    if bad:
        for r in bad:
            st.warning(f"**{r['sheet']}** ({r['name']}): " + " · ".join(r["notes"]))
    else:
        st.success(f"{len(rows)} baris, semua kolom terbaca ✓")

    # ---- salin
    data_rows = edited[COLNAMES].values.tolist()
    st.subheader("Salin semua baris")
    st.caption("Klik ikon salin di pojok kanan atas kotak, lalu tempel (Ctrl+V) ke Excel.")
    st.code(tsv(data_rows, opt_head), language=None)

    st.subheader("Salin satu baris")
    idx = st.selectbox("Pilih baris", range(len(rows)), format_func=lambda i: f"{rows[i]['sheet']} ({rows[i]['name']})")
    st.code(tsv([data_rows[idx]], opt_head), language=None)

    st.caption("Tips: klik sel di tabel untuk mengoreksi nilai, lalu salin dan tempel (Ctrl+V) ke Excel.")


if __name__ == "__main__":  # bisa juga dijalankan sendiri
    st.set_page_config(page_title="Noon Report Converter", page_icon="🚢", layout="wide")
    render()
