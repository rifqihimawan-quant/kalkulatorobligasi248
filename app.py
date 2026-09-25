"""
bond_price_list.py — "List Harga Obligasi" tab for Kalkulator Obligasi.

Self-contained add-on. Nothing in app.py is modified except two lines
(an import and one extra st.tabs entry) — see the README block at the end.

What it does
------------
1. Scrapes the BCA secondary-market bond table (server-rendered HTML table,
   parsed with the stdlib HTMLParser — no bs4/lxml/pandas.read_html needed).
2. Renders it with a filter widget for every column.
3. Exports the (filtered) list to a multi-page PDF laid out like the
   "INDIKASI HARGA OBLIGASI" sheet, named HARGA INDIKASI OBLIGASI_YYYYMMDD.pdf.
4. Small refresh button that busts the cache and re-fetches.

Dependencies: streamlit, matplotlib, requests, pandas — all already pulled in
by the existing app / by Streamlit itself. No new requirements.txt entries.
"""

from __future__ import annotations

import datetime as dt
import io
import re
import textwrap
from html.parser import HTMLParser

import streamlit as st
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages

BCA_URL = ("https://www.bca.co.id/id/individu/produk/investasi-dan-asuransi/"
           "obligasi/pilihan-produk-obligasi")

_UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
       "(KHTML, like Gecko) Chrome/125.0 Safari/537.36")

# Indonesian month abbreviations as the BCA table prints them.
_ID_MON = {"jan": 1, "feb": 2, "mar": 3, "apr": 4, "mei": 5, "jun": 6,
           "jul": 7, "agt": 8, "agu": 8, "ags": 8, "sep": 9, "okt": 10,
           "nov": 11, "des": 12,
           # tolerate English spellings if BCA ever switches the locale
           "may": 5, "aug": 8, "oct": 10, "dec": 12}

_EN_MON = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
           "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

# Group prefixes, longest / most specific first.
_GROUPS = ("USDFR", "INDOIS", "INDON", "FR", "ORI", "PBS", "SR", "ST")
_USD_GROUPS = {"USDFR", "INDOIS", "INDON"}


# ===========================================================================
# HTML table extraction
# ===========================================================================
class _TableParser(HTMLParser):
    """Minimal table scraper. Collects every <table> as a list of row-lists."""

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.tables: list[list[list[str]]] = []
        self._rows = None
        self._row = None
        self._cell = None

    def handle_starttag(self, tag, attrs):
        if tag == "table":
            # A nested table would clobber the outer one; push the outer aside.
            if self._rows is not None:
                self.tables.append(self._rows)
            self._rows = []
        elif tag == "tr" and self._rows is not None:
            self._row = []
        elif tag in ("td", "th") and self._row is not None:
            self._cell = []
        elif tag == "br" and self._cell is not None:
            self._cell.append(" ")

    def handle_endtag(self, tag):
        if tag == "table" and self._rows is not None:
            self.tables.append(self._rows)
            self._rows = None
        elif tag == "tr" and self._row is not None:
            if any(c for c in self._row):
                self._rows.append(self._row)
            self._row = None
        elif tag in ("td", "th") and self._cell is not None:
            self._row.append(" ".join("".join(self._cell).split()))
            self._cell = None

    def handle_data(self, data):
        if self._cell is not None:
            self._cell.append(data)


def _pick_price_table(tables):
    """The price table is the one whose header mentions 'Nama Produk'."""
    for tbl in tables:
        for probe in tbl[:3]:
            joined = " | ".join(probe).lower()
            if "nama produk" in joined and "jatuh tempo" in joined:
                return tbl, probe
    return None, None


def _column_map(header):
    """Map logical fields to column indices. Order of tests matters:
    'Imbal Hasil Hingga Jatuh Tempo' also contains 'jatuh tempo'."""
    idx = {}
    for i, raw in enumerate(header):
        h = raw.lower()
        if "imbal" in h or "ytm" in h:
            idx["ytm"] = i
        elif "nama produk" in h:
            idx["name"] = i
        elif "jatuh tempo" in h:
            idx["maturity"] = i
        elif "kupon" in h:
            idx["coupon"] = i
        elif "beli" in h:
            idx["buy"] = i
        elif "jual" in h:
            idx["sell"] = i
        elif "update" in h:
            idx["updated"] = i
    return idx


def _parse_num(text):
    """
    Decimal-tolerant number parser scoped to this table.

    Every value here (coupon, price, yield) sits between 0 and ~200, so a lone
    separator is always a decimal point — '8,375' is 8.375, not 8375. Mixed
    separators fall back to 'rightmost one is the decimal'.
    """
    if text is None:
        return None
    s = str(text).strip().replace("%", "").replace("\u00a0", "").replace(" ", "")
    if not s or s in ("-", "--", "N/A", "n/a"):
        return None
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


def _parse_date(text):
    """'15 Sep 2026' / '15 Agt 2026' -> datetime.date."""
    if not text:
        return None
    m = re.search(r"(\d{1,2})\s*[-/\s]\s*([A-Za-z]{3,})\s*[-/\s]\s*(\d{4})", str(text))
    if not m:
        return None
    day, mon, year = int(m.group(1)), m.group(2)[:3].lower(), int(m.group(3))
    if mon not in _ID_MON:
        return None
    try:
        return dt.date(year, _ID_MON[mon], day)
    except ValueError:
        return None


def _fmt_en(d):
    return "" if d is None else f"{d.day:02d} {_EN_MON[d.month - 1]} {d.year}"


_SERI_RE = re.compile(r"\bseri\b", re.IGNORECASE)


def _extract_code(name):
    """'Obligasi Negara Valas Seri INDON27NEWNEW-Callable' -> the trailing code."""
    parts = _SERI_RE.split(name)
    tail = parts[-1].strip() if len(parts) > 1 else name.strip()
    tail = tail.split()[-1] if tail.split() else name.strip()
    return tail.upper()


def _group_of(code):
    c = code.upper()
    for g in _GROUPS:
        if c.startswith(g):
            return g
    return "LAIN"


def _build_record(cells, cmap, today):
    def cell(key):
        i = cmap.get(key)
        return cells[i] if i is not None and i < len(cells) else ""

    name = cell("name")
    if not name:
        return None
    code = _extract_code(name)
    if not re.search(r"\d", code):          # a code always carries digits
        return None

    maturity = _parse_date(cell("maturity"))
    updated = _parse_date(cell("updated"))
    group = _group_of(code)
    duration = None if maturity is None else round((maturity - today).days / 365.25, 2)

    return {
        "code": code,
        "name": name,
        "group": group,
        "currency": "USD" if group in _USD_GROUPS else "IDR",
        "maturity": maturity,
        "maturity_en": _fmt_en(maturity),
        "coupon": _parse_num(cell("coupon")),
        "buy": _parse_num(cell("buy")),
        "sell": _parse_num(cell("sell")),
        "ytm": _parse_num(cell("ytm")),
        "updated": updated,
        "updated_en": _fmt_en(updated),
        "duration": duration,
    }


def parse_bca_html(html, today=None):
    today = today or dt.date.today()
    p = _TableParser()
    p.feed(html)
    tbl, header = _pick_price_table(p.tables)
    if tbl is None:
        raise ValueError(
            "Tabel harga tidak ditemukan di halaman BCA. Struktur halaman mungkin "
            "berubah, atau isi tabel dimuat lewat JavaScript sehingga tidak ada di HTML."
        )
    cmap = _column_map(header)
    missing = {"name", "maturity", "buy", "sell"} - set(cmap)
    if missing:
        raise ValueError(f"Kolom wajib tidak terbaca: {', '.join(sorted(missing))}.")

    out, seen = [], set()
    for row in tbl:
        if row is header:
            continue
        rec = _build_record(row, cmap, today)
        if rec and rec["code"] not in seen:
            seen.add(rec["code"])
            out.append(rec)
    if not out:
        raise ValueError("Tabel ditemukan tetapi tidak ada baris data yang terbaca.")
    return out


# ===========================================================================
# Fetch (cached; the refresh button clears the cache)
# ===========================================================================
@st.cache_data(ttl=900, show_spinner=False)
def fetch_bond_prices():
    """Returns {'rows': [...], 'fetched_at': datetime}. Raises on failure."""
    import requests
    resp = requests.get(
        BCA_URL,
        headers={"User-Agent": _UA,
                 "Accept": "text/html,application/xhtml+xml",
                 "Accept-Language": "id-ID,id;q=0.9,en;q=0.8"},
        timeout=12,
    )
    resp.raise_for_status()
    resp.encoding = resp.encoding or "utf-8"
    rows = parse_bca_html(resp.text)
    return {"rows": rows, "fetched_at": dt.datetime.now()}


# ===========================================================================
# PDF export — laid out to match the "INDIKASI HARGA OBLIGASI" sheet
# ===========================================================================
PDF_TITLE = "INDIKASI HARGA OBLIGASI"

PDF_NOTES = [
    "Harga saat transaksi dapat berubah dari indikasi di tabel sesuai kondisi pasar.",
    "Untuk harga ter-update, hubungi Personal Banker/Relationship Manager anda maupun "
    "akses langsung melalui aplikasi Welma@myBCA",
]

PDF_DISCLAIMER = (
    "DISCLAIMER: This report is for information only, and is not intended as an offer or "
    "solicitation with respect to the purchase or sale of any commodities, securities, or "
    "currencies. We deem that the information contained in this report has been taken from "
    "sources which we deem reliable. However, we do not guarantee their accuracy, and any such "
    "information may be incomplete or condensed. None of PT. Bank Central Asia Tbk (\u201cBCA\u201d), "
    "and/or its affiliated companies, and/or their respective employees and/or agents makes any "
    "representation or warranty (express or implied) or accepts any responsibility or liability "
    "as to, or in relation to, the accuracy or completeness of the information and opinions "
    "contained in this report or as to any information contained in this report or any other such "
    "information or opinions remaining unchanged after the issue thereof. BCA, or any of its "
    "related companies or any individuals connected with BCA or BCA group accepts no liability "
    "for any direct, special, indirect, consequential, incidental damages or any other loss or "
    "damages of any kind arising from any use of the information herein (including any error, "
    "omission or misstatement herein, negligent or otherwise) or further communication thereof, "
    "even if the BCA or any other person has been advised of the possibility thereof. Opinion "
    "expressed is the analysts\u2019 current personal views as of the date appearing on this "
    "material only, and subject to change without notice."
)


def _t(v, fmt):
    return "-" if v is None else format(v, fmt)


# (label, width fraction, alignment, value getter)
_PDF_COLS = [
    ("Obligasi Name",              0.190, "left",   lambda r: r["code"]),
    ("Currency",                   0.070, "center", lambda r: r["currency"]),
    ("Due Date",                   0.110, "center", lambda r: r["maturity_en"] or "-"),
    ("Coupon Tier\n(% p.a.)",      0.100, "right",  lambda r: _t(r["coupon"], ".3f")),
    ("Price Buy (%)",              0.100, "right",  lambda r: _t(r["buy"], ".2f")),
    ("Price Sell (%)",             0.100, "right",  lambda r: _t(r["sell"], ".2f")),
    ("Yield to Maturity/YTM\n(%)", 0.120, "right",  lambda r: _t(r["ytm"], ".2f")),
    ("Last Update",                0.110, "center", lambda r: r["updated_en"] or "-"),
    ("Duration to\nMaturity (yr)", 0.100, "right",  lambda r: _t(r["duration"], ".2f")),
]

_INK = "#111111"
_MUTED = "#5c7085"
_HAIR = "#d9e0e6"

_W, _H = 11.69, 8.27                      # A4 landscape
_ML = _MR = 0.42
_MT, _MB = 0.40, 0.42
_ROW_H = 0.153
_HDR_H = 0.40


def _new_page(pdf_fig_list):
    fig = plt.figure(figsize=(_W, _H), dpi=150)
    fig.patch.set_facecolor("white")
    ax = fig.add_axes([0, 0, 1, 1])
    ax.set_xlim(0, _W)
    ax.set_ylim(0, _H)
    ax.axis("off")
    pdf_fig_list.append(fig)
    return fig, ax


def _col_positions():
    """Returns [(x_left, width), ...] in inches."""
    table_w = _W - _ML - _MR
    xs, x = [], _ML
    for _, frac, _a, _g in _PDF_COLS:
        w = table_w * frac
        xs.append((x, w))
        x += w
    return xs


def _place(ax, x, w, y, text, align, size, color=_INK, weight="normal", pad=0.07):
    if align == "left":
        ax.text(x + pad, y, text, fontsize=size, color=color, weight=weight,
                ha="left", va="center")
    elif align == "right":
        ax.text(x + w - pad, y, text, fontsize=size, color=color, weight=weight,
                ha="right", va="center")
    else:
        ax.text(x + w / 2, y, text, fontsize=size, color=color, weight=weight,
                ha="center", va="center")


def _draw_header(ax, as_of, page_no):
    y = _H - _MT
    ax.text(_ML, y, PDF_TITLE, fontsize=13.5, color=_INK, weight="bold",
            ha="left", va="top")
    ax.text(_W - _MR, y, f"{as_of.day:02d}-{_EN_MON[as_of.month - 1]}-{as_of.year}",
            fontsize=10.5, color=_INK, ha="right", va="top")
    ax.text(_W - _MR, _MB * 0.55, f"Halaman {page_no}", fontsize=6.5,
            color=_MUTED, ha="right", va="center")
    return y - 0.34


def _draw_table_header(ax, y_top):
    xs = _col_positions()
    y_bot = y_top - _HDR_H
    ax.plot([_ML, _W - _MR], [y_top, y_top], color=_INK, lw=1.0)
    ax.plot([_ML, _W - _MR], [y_bot, y_bot], color=_INK, lw=1.0)
    for (x, w), (label, _f, align, _g) in zip(xs, _PDF_COLS):
        _place(ax, x, w, (y_top + y_bot) / 2, label, align, 7.0, weight="bold")
    return y_bot


def _draw_rows(ax, y_top, chunk):
    xs = _col_positions()
    y = y_top
    for rec in chunk:
        yc = y - _ROW_H / 2
        for (x, w), (_l, _f, align, getter) in zip(xs, _PDF_COLS):
            _place(ax, x, w, yc, getter(rec), align, 7.0)
        y -= _ROW_H
        ax.plot([_ML, _W - _MR], [y, y], color=_HAIR, lw=0.4)
    return y


def _footer_height():
    lines = textwrap.wrap(PDF_DISCLAIMER, width=196)
    return 0.16 + len(PDF_NOTES) * 0.155 + 0.14 + len(lines) * 0.118 + 0.10


def _draw_footer(ax, y_top):
    y = y_top - 0.16
    for note in PDF_NOTES:
        ax.text(_ML, y, note, fontsize=7.0, color=_INK, ha="left", va="center")
        y -= 0.155
    y -= 0.14
    for line in textwrap.wrap(PDF_DISCLAIMER, width=196):
        ax.text(_ML, y, line, fontsize=5.1, color=_INK, ha="left", va="center")
        y -= 0.118
    return y


def render_price_list_pdf(rows, as_of=None, sort_by_code=True):
    """Multi-page PDF of the price list. Returns bytes."""
    if not rows:
        raise ValueError("Tidak ada baris untuk diekspor.")
    as_of = as_of or dt.date.today()
    data = sorted(rows, key=lambda r: r["code"]) if sort_by_code else list(rows)

    # rows that fit on a page once the title band and table header are drawn
    usable = (_H - _MT - 0.34 - _HDR_H) - _MB
    per_page = max(1, int(usable / _ROW_H))

    chunks = [data[i:i + per_page] for i in range(0, len(data), per_page)]
    foot_h = _footer_height()
    figs = []

    for page_no, chunk in enumerate(chunks, 1):
        _fig, ax = _new_page(figs)
        y = _draw_header(ax, as_of, page_no)
        y = _draw_table_header(ax, y)
        y = _draw_rows(ax, y, chunk)
        last = page_no == len(chunks)
        if last and (y - _MB) >= foot_h:
            _draw_footer(ax, y)
        elif last:
            _fig2, ax2 = _new_page(figs)
            y2 = _draw_header(ax2, as_of, page_no + 1)
            _draw_footer(ax2, y2)

    buf = io.BytesIO()
    with PdfPages(buf) as pdf:
        for fig in figs:
            pdf.savefig(fig)
            plt.close(fig)
    return buf.getvalue()


def pdf_filename(when=None):
    when = when or dt.date.today()
    return f"HARGA INDIKASI OBLIGASI_{when:%Y%m%d}.pdf"


# ===========================================================================
# Filtering
# ===========================================================================
def _span(rows, key):
    vals = [r[key] for r in rows if r.get(key) is not None]
    return (min(vals), max(vals)) if vals else (None, None)


def _range_slider(label, key, lo, hi, step=0.01, fmt="%.2f"):
    if lo is None or hi is None:
        return None
    lo, hi = float(lo), float(hi)
    if hi - lo < step:
        hi = lo + step
    return st.slider(label, min_value=lo, max_value=hi, value=(lo, hi),
                     step=step, format=fmt, key=key)


def _in_range(value, rng):
    if rng is None:
        return True
    if value is None:
        return False
    return rng[0] - 1e-9 <= value <= rng[1] + 1e-9


def _apply_filters(rows, f):
    out = []
    for r in rows:
        if f["text"] and f["text"].lower() not in (r["code"] + " " + r["name"]).lower():
            continue
        if f["groups"] and r["group"] not in f["groups"]:
            continue
        if f["currencies"] and r["currency"] not in f["currencies"]:
            continue
        if f["mat_from"] and (r["maturity"] is None or r["maturity"] < f["mat_from"]):
            continue
        if f["mat_to"] and (r["maturity"] is None or r["maturity"] > f["mat_to"]):
            continue
        if f["updates"] and (r["updated_en"] not in f["updates"]):
            continue
        if not _in_range(r["coupon"], f["coupon"]):
            continue
        if not _in_range(r["buy"], f["buy"]):
            continue
        if not _in_range(r["sell"], f["sell"]):
            continue
        if not _in_range(r["ytm"], f["ytm"]):
            continue
        if not _in_range(r["duration"], f["duration"]):
            continue
        out.append(r)
    return out


# ===========================================================================
# Streamlit tab
# ===========================================================================
_FILTER_KEYS = ["pl_text", "pl_group", "pl_cur", "pl_mat_from", "pl_mat_to",
                "pl_upd", "pl_coupon", "pl_buy", "pl_sell", "pl_ytm", "pl_dur"]


def tab_price_list():
    head, btn = st.columns([8, 1])
    with head:
        st.markdown('<div class="ko-cap">Daftar harga obligasi pasar sekunder (BCA)</div>',
                    unsafe_allow_html=True)
    with btn:
        refresh = st.button("🔄", key="pl_refresh", help="Ambil ulang harga dari situs BCA")

    if refresh:
        fetch_bond_prices.clear()
        st.session_state["pl_loaded"] = True

    # Streamlit re-runs the whole script on every interaction, and st.tabs
    # renders every tab body regardless of which one is on screen. An
    # unconditional network call here would therefore block the four
    # simulation tabs too, on every click. Fetch only when asked.
    if not st.session_state.get("pl_loaded"):
        st.info("Tekan tombol di bawah untuk mengambil daftar harga terbaru dari situs BCA.")
        if st.button("Muat daftar harga", type="primary", key="pl_load", **_full_width()):
            st.session_state["pl_loaded"] = True
            st.rerun()
        return

    try:
        with st.spinner("Mengambil daftar harga…"):
            payload = fetch_bond_prices()
    except Exception as exc:  # noqa: BLE001 — surfaced to the user verbatim
        st.error(f"Gagal mengambil data harga: {exc}")
        st.caption(
            "Jika pesan menyebut tabel tidak ditemukan, halaman BCA kemungkinan "
            "memuat tabel lewat JavaScript atau memblokir permintaan dari server "
            "Streamlit. Coba tombol 🔄, atau buka halaman sumber di bawah."
        )
        st.link_button("Buka halaman BCA", BCA_URL)
        return

    rows = payload["rows"]
    fetched = payload["fetched_at"]
    st.caption(f"{len(rows)} seri · diambil {fetched:%d %b %Y %H:%M} WIB · sumber: bca.co.id")

    # ---- filters, one per column ----
    with st.expander("Filter", expanded=True):
        c1, c2, c3 = st.columns(3)
        with c1:
            text = st.text_input("Nama / kode", "", key="pl_text",
                                 placeholder="mis. FR01, INDON, ORI027")
            groups = st.multiselect(
                "Jenis", sorted({r["group"] for r in rows}), key="pl_group")
            currencies = st.multiselect(
                "Mata uang", sorted({r["currency"] for r in rows}), key="pl_cur")
        with c2:
            mats = [r["maturity"] for r in rows if r["maturity"]]
            mat_from = st.date_input("Jatuh tempo dari", value=None,
                                     min_value=min(mats) if mats else None,
                                     max_value=max(mats) if mats else None,
                                     key="pl_mat_from")
            mat_to = st.date_input("Jatuh tempo s.d.", value=None,
                                   min_value=min(mats) if mats else None,
                                   max_value=max(mats) if mats else None,
                                   key="pl_mat_to")
            upd_options = sorted({r["updated_en"] for r in rows if r["updated_en"]},
                                 key=lambda s: _parse_date(s) or dt.date.min, reverse=True)
            updates = st.multiselect("Last update", upd_options, key="pl_upd")
        with c3:
            lo, hi = _span(rows, "coupon")
            coupon = _range_slider("Kupon (% p.a.)", "pl_coupon", lo, hi, 0.005, "%.3f")
            lo, hi = _span(rows, "ytm")
            ytm = _range_slider("YTM (%)", "pl_ytm", lo, hi, 0.01, "%.2f")
            lo, hi = _span(rows, "duration")
            duration = _range_slider("Sisa tenor (tahun)", "pl_dur", lo, hi, 0.1, "%.2f")

        c4, c5, c6 = st.columns(3)
        with c4:
            lo, hi = _span(rows, "buy")
            buy = _range_slider("Harga beli (%)", "pl_buy", lo, hi, 0.05, "%.2f")
        with c5:
            lo, hi = _span(rows, "sell")
            sell = _range_slider("Harga jual (%)", "pl_sell", lo, hi, 0.05, "%.2f")
        with c6:
            st.write("")
            st.write("")
            if st.button("Reset filter", key="pl_reset", **_full_width()):
                for k in _FILTER_KEYS:
                    st.session_state.pop(k, None)
                st.rerun()

    filtered = _apply_filters(rows, dict(
        text=text, groups=groups, currencies=currencies,
        mat_from=mat_from, mat_to=mat_to, updates=updates,
        coupon=coupon, buy=buy, sell=sell, ytm=ytm, duration=duration,
    ))

    st.caption(f"Menampilkan {len(filtered)} dari {len(rows)} seri.")
    if not filtered:
        st.info("Tidak ada seri yang cocok dengan filter.")
        return

    import pandas as pd
    df = pd.DataFrame([{
        "Obligasi": r["code"],
        "Mata Uang": r["currency"],
        "Jatuh Tempo": r["maturity"],
        "Kupon (% p.a.)": r["coupon"],
        "Harga Beli (%)": r["buy"],
        "Harga Jual (%)": r["sell"],
        "YTM (%)": r["ytm"],
        "Last Update": r["updated"],
        "Sisa Tenor (thn)": r["duration"],
    } for r in filtered])

    st.dataframe(
        df, hide_index=True, height=min(38 + len(df) * 35, 560),
        column_config={
            "Jatuh Tempo": st.column_config.DateColumn(format="DD MMM YYYY"),
            "Last Update": st.column_config.DateColumn(format="DD MMM YYYY"),
            "Kupon (% p.a.)": st.column_config.NumberColumn(format="%.3f"),
            "Harga Beli (%)": st.column_config.NumberColumn(format="%.2f"),
            "Harga Jual (%)": st.column_config.NumberColumn(format="%.2f"),
            "YTM (%)": st.column_config.NumberColumn(format="%.2f"),
            "Sisa Tenor (thn)": st.column_config.NumberColumn(format="%.2f"),
        },
        **_full_width(),
    )

    st.divider()
    st.markdown('<div class="ko-cap">Ekspor</div>', unsafe_allow_html=True)
    scope_all = st.checkbox("Ekspor seluruh daftar (abaikan filter)", value=False,
                            key="pl_scope")
    export_rows = rows if scope_all else filtered
    try:
        pdf_bytes = render_price_list_pdf(export_rows)
    except Exception as exc:  # noqa: BLE001
        st.warning(f"PDF tidak dapat dibuat: {exc}")
        return
    st.download_button(
        f"Unduh PDF ({len(export_rows)} seri)",
        data=pdf_bytes,
        file_name=pdf_filename(),
        mime="application/pdf",
        type="primary",
        **_full_width(),
    )
    st.caption(f"Nama berkas: {pdf_filename()}")


def _full_width():
    """Mirror of app.py's FULL_WIDTH probe, kept local so this module has no
    import dependency on app.py (which would create a circular import)."""
    import inspect
    try:
        if "width" in inspect.signature(st.button).parameters:
            return {"width": "stretch"}
    except (ValueError, TypeError):
        pass
    return {"use_container_width": True}


# ===========================================================================
# README — the only two edits needed in app.py
# ===========================================================================
#
#   1) near the top, with the other imports:
#
#        from bond_price_list import tab_price_list
#
#   2) inside main(), replace the tabs line and add one `with` block:
#
#        beli, jual, switching, kredit, harga = st.tabs(
#            ["Simulasi Beli", "Simulasi Jual", "Simulasi Switching",
#             "Simulasi Kredit", "List Harga Obligasi"])
#        ...
#        with harga:
#            tab_price_list()
#
# Nothing else in app.py changes.
