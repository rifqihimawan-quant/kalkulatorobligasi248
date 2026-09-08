"""
Kalkulator Obligasi — Streamlit edition.

Ported from Kalkulator_Obligasi_2_4_8.xls. Every formula below carries the
originating cell reference so the two can be diffed. The hidden calculation
blocks in the workbook are rows 202-330 (Simulasi Beli), 201-388 (Simulasi
Jual) and 208-500 (Simulasi Switching); the product master table is C333:R497.

Run locally:  streamlit run app.py
"""

from __future__ import annotations

import datetime as dt
import io
import json
import math
from dataclasses import dataclass

import streamlit as st
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.backends.backend_agg import FigureCanvasAgg

# ---------------------------------------------------------------------------
# Product master table — C333:R497 of "Simulasi Beli".
# Columns: code, issued, maturity, coupon, coupon day, first month, second
# month, first accrual date, first coupon date. Dates are Excel serials.
# ---------------------------------------------------------------------------
PRODUCTS_RAW = json.loads(r"""[["FR0056",40444,46280,0.08375,15,3,9,null,null],["FR0058",40745,48380,0.0825,15,12,6,null,null],["FR0059",40801,46522,0.07,15,11,5,null,null],["FR0064",41134,46888,0.06125,15,11,5,null,null],["FR0065",41151,48714,0.06625,15,11,5,null,null],["FR0068",41487,49018,0.08375,15,9,3,null,null],["FR0071",41529,47192,0.09,15,9,3,null,null],["FR0072",42194,49810,0.0825,15,11,5,null,null],["FR0073",42222,47983,0.0875,15,11,5,null,null],["FR0074",42684,48441,0.075,15,2,8,null,null],["FR0075",42957,50540,0.075,15,11,5,null,null],["FR0076",43000,54193,0.07375,15,11,5,null,null],["FR0078",43235,47253,0.0825,15,11,5,null,null],["FR0079",43388,50875,0.08375,15,4,10,null,null],["FR0080",43631,49475,0.075,15,12,6,null,null],["FR0082",43539,47741,0.07,15,9,3,null,null],["FR0083",43753,51241,0.075,15,4,10,null,null],["FR0084",43876,46068,0.0725,15,8,2,null,null],["FR0085",43936,47953,0.0775,15,10,4,null,null],["FR0086",43936,46127,0.055,15,10,4,null,null],["FR0087",43876,47894,0.065,15,8,2,null,null],["FR0088",44180,49841,0.0625,15,6,12,null,null],["FR0089",44058,55380,0.06875,15,2,8,null,null],["FR0090",44301,46492,0.05125,15,10,4,44301,44484],["FR0091",44301,48319,0.06375,15,10,4,44301,44484],["FR0092",44362,52032,0.07125,15,12,6,44362,44545],["FR0093",44392,50236,0.06375,15,1,7,null,null],["FR0094",44576,46767,0.056,15,7,1,null,null],["FR0095",44788,46980,0.06375,15,2,8,null,null],["FR0096",44788,48625,0.07,15,2,8,null,null],["FR0097",44727,52397,0.07125,15,12,6,null,null],["FR0098",44727,50571,0.07125,15,12,6,44727,44910],["FR0099",44941,47133,0.064,15,7,1,null,null],["FR0100",45153,48990,0.06625,15,2,8,45153,45337],["FR0101",45214,47223,0.06875,15,4,10,45214,45397],["FR0102",45306,56445,0.06875,15,1,7,null,null],["FR0103",45512,49505,0.0675,15,1,7,null,null],["FR0104",45525,47679,0.065,15,1,7,null,null],["FR0106",45298,51363,0.07125,15,2,8,null,null],["FR0107",45298,53189,0.07125,15,2,8,null,null],["FR0108",45867,49780,0.065,15,4,10,null,null],["FR0109",45881,47922,0.05875,15,9,3,null,null],["FR0110",46238,48349,0.0725,15,11,5,null,null],["FR0111",46266,50114,0.07,15,9,3,null,null],["USDFR0003",44576,48228,0.03,15,7,1,null,null],["INDOIS26",42458,46110,0.0455,29,9,3,42458,42642],["INDOIS26NEW",44356,46182,0.015,9,12,6,44356,44539],["INDOIS27",42823,46475,0.0415,29,9,3,42823,43007],["INDOIS27NEW",44718,46544,0.044,6,12,6,44718,44901],["INDOIS28",43160,46813,0.044000000000000004,1,9,3,43160,43344],["INDOIS28NEW",45245,47072,0.054,15,5,11,45245,45427],["INDOIS29",43516,47169,0.044500000000000005,20,8,2,43516,43697],["INDOIS29NEW",45475,47301,0.051,2,1,7,45475,45659],["INDOIS30",44005,47657,0.028,23,12,6,44005,44188],["INDOIS30NEW",45621,47628,0.05,25,5,11,45621,45802],["INDOIS30NEWNEW",45861,47687,0.0455,23,1,7,45861,46045],["INDOIS30NEWNEWNEW",45992,47818,0.045,1,6,12,45992,46174],["INDOIS31",44356,48008,0.0255,9,12,6,44356,44539],["INDOIS32",44719,48371,0.047,6,12,6,44718,44901],["INDOIS33",45245,48898,0.056,15,5,11,45245,45427],["INDOIS34",45475,49127,0.052,2,1,7,45475,45659],["INDOIS34NEW",45621,49273,0.0525,25,5,11,45621,45802],["INDOIS35",45861,49513,0.052,23,1,7,45861,46045],["INDOIS35NEW",45992,49644,0.05,1,6,12,45992,46174],["INDOIS50",44005,54962,0.038,23,12,6,44005,44188],["INDOIS51",44356,55313,0.0355,9,12,6,44356,44539],["INDOIS54",45475,56432,0.055,2,1,7,45475,45659],["INDOIS54NEW",45621,56578,0.0565,25,5,11,45621,45802],["INDON26",42346,46030,0.0475,8,7,1,42346,42559],["INDON27",42712,46395,0.0435,8,7,1,42712,42924],["INDON27NEW",42934,46586,0.0385,18,1,7,42934,43118],["INDON27NEWNEW",44824,46650,0.0415,20,3,9,44824,45005],["INDON28",43080,46763,0.035,11,7,1,43080,43292],["INDON28NEW",43214,46867,0.040999999999999995,24,10,4,43214,43397],["INDON28NEWNEW",44937,46763,0.0455,11,7,1,44937,45118],["INDON29",43445,47160,0.0475,11,8,2,43445,43688],["INDON29NEW",43634,47379,0.034,18,3,9,43634,43908],["INDON29NEWNEW",45301,47187,0.044,10,9,3,45301,45545],["INDON30",43844,47528,0.0285,14,8,2,43844,44057],["INDON30NEW",43936,47771,0.0385,15,10,4,43936,44119],["INDON30NEWNEW",45672,47498,0.0525,15,7,1,45672,45853],["INDON31",44208,47919,0.0185,12,9,3,44208,44451],["INDON31NEW",44405,48057,0.0215,28,1,7,44405,44589],["INDON31NEWNEW",45946,47954,0.043,16,4,10,45946,46128],["INDON31NEWNEWNEW",46043,47900,0.0435,21,8,2,46043,46255],["INDON31NEWNEWNEWNEW",46175,47997,0.0503,29,11,5,46171,46355],["INDON32",44651,48304,0.0355,30,9,1,44651,44834],["INDON32NEW",44824,48477,0.0465,20,3,9,44824,45005],["INDON33",44937,48590,0.0485,11,7,1,44937,45118],["INDON34",45301,48985,0.047,10,8,2,45301,45514],["INDON34NEW",45545,49197,0.0475,10,3,9,45545,45726],["INDON35",38637,49594,0.085,12,4,10,38637,38819],["INDON35NEW",45672,49324,0.056,15,7,1,45672,45853],["INDON36",45946,49781,0.049,16,4,10,45946,46128],["INDON36NEW",46043,49726,0.0495,21,8,2,46043,46255],["INDON36NEWNEW",46175,49824,0.0569,29,11,5,46171,46355],["INDON37",39127,50088,0.06625,17,8,2,39130,39311],["INDON38",39464,50422,0.0775,17,7,1,39464,39646],["INDON42",40925,51883,0.0525,17,7,1,40925,41107],["INDON43",41379,52336,0.04625,15,10,4,41379,41562],["INDON44",41654,52611,0.0675,15,7,1,41654,41835],["INDON45",42019,52977,0.05125,15,7,1,42019,42200],["INDON46",42346,53335,0.059500000000000004,8,7,1,42346,42559],["INDON47",42712,53700,0.0525,8,7,1,42712,42924],["INDON47NEW",42934,53891,0.0475,18,1,7,42934,43118],["INDON48",43080,54068,0.0435,11,7,1,43080,43292],["INDON49",43445,54465,0.0535,11,8,2,43445,43688],["INDON49NEW",43768,54726,0.037000000000000005,30,4,10,43768,43951],["INDON50",43844,54833,0.035,14,8,2,43844,44057],["INDON50NEW",43936,55076,0.042,15,10,4,43936,44119],["INDON51",44208,55224,0.0305,12,9,3,44208,44451],["INDON52",44651,55609,0.043,30,9,3,44651,44834],["INDON52NEW",44824,55782,0.0545,20,3,9,44824,45005],["INDON53",44937,55895,0.0565,11,7,1,44937,45118],["INDON54",45301,56290,0.051,10,8,2,45301,45514],["INDON54NEW",45545,56502,0.0515,10,3,9,45545,45726],["INDON56",46043,57031,0.05475,21,8,2,46043,46255],["INDON61",44462,59072,0.032,23,3,9,44462,44643],["INDON70",43936,62198,0.0445,15,10,4,43936,44119],["INDON71",44208,62529,0.0335,12,9,3,44208,44451],["ORI030T3",46239,47314,0.069,15,9,10,null,null],["ORI030T6",46239,48410,0.07,15,9,10,null,null],["ORI029T3",46078,47164,0.0545,15,4,5,null,null],["ORI029T6",46078,48259,0.058,15,4,5,null,null],["ORI028T3",45959,47041,0.0535,15,12,1,null,null],["ORI028T6",45959,48136,0.0565,15,12,1,null,null],["ORI027T3",45714,46798,0.0665,15,4,5,null,null],["ORI027T6",45714,47894,0.0675,15,4,5,null,null],["ORI026T3",45595,46675,0.063,15,12,1,null,null],["ORI026T6",45595,47771,0.064,15,12,1,null,null],["ORI025T3",45350,46433,0.0625,15,4,5,null,null],["ORI025T6",45350,47529,0.064,15,4,5,null,null],["ORI024T3",45238,46310,0.061,15,12,1,null,null],["ORI024T6",45238,47406,0.0635,15,12,1,null,null],["ORI023T3",45133,46218,0.059,15,9,10,null,null],["ORI023T6",45133,47314,0.061,15,9,10,null,null],["ORI022",44860,45945,0.0595,15,12,1,null,null],["PBS003",40923,46402,0.06,15,7,1,null,null],["PBS004",40955,50086,0.061,15,8,2,null,null],["PBS012",42397,48167,0.08875,15,5,11,null,null],["PBS029",44089,49018,0.06375,15,3,9,null,null],["PBS030",44392,46949,0.05875,15,1,7,null,null],["PBS032",44392,46218,0.04875,15,1,7,null,null],["PBS033",44545,53858,0.0675,15,6,12,null,null],["PBS034",44545,50936,0.065,15,6,12,null,null],["PBS035",44635,51940,0.0675,15,9,3,null,null],["PBS036",44788,45884,0.05375,15,2,8,null,null],["PBS037",44819,49749,0.06875,15,3,11,null,null],["PBS038",45267,54772,0.06875,15,12,6,null,null],["PBS040",45792,47802,0.05,15,11,5,null,null],["SR025T3",46288,47371,0.068,10,10,11,null,null],["SR025T5",46288,48101,0.069,10,10,11,null,null],["SR024T3",46134,47187,0.0555,10,5,6,null,null],["SR024T5",46134,47917,0.059,10,5,6,null,null],["SR023T3",45922,47036,0.058,10,11,12,null,null],["SR023T5",45922,47766,0.0595,10,11,12,null,null],["SR022T3",45833,46914,0.0645,10,8,9,null,null],["SR022T5",45833,47644,0.0655,10,8,9,null,null],["SR021T3",45560,46640,0.0635,10,11,12,null,null],["SR021T5",45560,47371,0.0645,10,11,12,null,null],["SR020T3",45385,46456,0.063,10,5,6,null,null],["SR020T5",45385,47187,0.064,10,5,6,null,null],["SR019T3",45196,46275,0.0595,10,11,12,null,null],["SR019T5",45196,47006,0.061,10,11,12,null,null],["SR018T3",45021,46091,0.0625,10,5,6,null,null],["SR018T5",45021,46822,0.064,10,5,6,null,null],["ST016T2",46183,46914,0.0605,10,7,8,null,null],["ST016T4",46183,47613,0.0625,10,7,8,null,null]]""")

TAX = 0.10
EPOCH = dt.date(1899, 12, 30)
MONTHS_ID = ["Januari", "Februari", "Maret", "April", "Mei", "Juni",
             "Juli", "Agustus", "September", "Oktober", "November", "Desember"]


# ===========================================================================
# Excel-compatible primitives
# ===========================================================================
def to_serial(d) -> int:
    if isinstance(d, (int, float)):
        return int(d)
    return (d - EPOCH).days


def to_date(serial) -> dt.date:
    return EPOCH + dt.timedelta(days=int(serial))


def fmt_date(serial) -> str:
    if serial is None:
        return "N/A"
    d = to_date(serial)
    return f"{d.day:02d} {MONTHS_ID[d.month - 1]} {d.year}"


def fmt_date_short(serial) -> str:
    d = to_date(serial)
    return f"{d.day:02d} {MONTHS_ID[d.month - 1][:3]} {d.year}"


_MON_EN = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


def fmt_date_en(serial) -> str:
    """Short English date (12-Aug-26), matching the original workbook export."""
    if serial is None:
        return "N/A"
    d = to_date(serial)
    return f"{d.day:02d}-{_MON_EN[d.month - 1]}-{str(d.year)[2:]}"


def _days_in_month(y: int, m: int) -> int:
    leap = (y % 4 == 0 and y % 100 != 0) or y % 400 == 0
    return [31, 29 if leap else 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31][m - 1]


def edate(serial, months: int) -> int:
    """Excel EDATE: same day of month, clamped to the month end."""
    d = to_date(serial)
    total = d.year * 12 + (d.month - 1) + months
    y, m = total // 12, total % 12 + 1
    return to_serial(dt.date(y, m, min(d.day, _days_in_month(y, m))))


def xdate(y: int, m: int, day: int) -> int:
    """Excel DATE(y, m, d) — month and day overflow roll forward."""
    y2, m2 = y + (m - 1) // 12, (m - 1) % 12 + 1
    return to_serial(dt.date(y2, m2, 1)) + (day - 1)


def datedif_m(a, b) -> int:
    """Excel DATEDIF unit "M" — complete calendar months."""
    if b < a:
        return -datedif_m(b, a)
    da, db = to_date(a), to_date(b)
    n = (db.year - da.year) * 12 + (db.month - da.month)
    if db.day < da.day:
        n -= 1
    return n


def datedif_ym(a, b) -> int:
    return abs(datedif_m(a, b)) % 12


def datedif_yd(a, b) -> int:
    if b < a:
        a, b = b, a
    da, db = to_date(a), to_date(b)
    try:
        anchor = dt.date(db.year, da.month, da.day)
    except ValueError:
        anchor = dt.date(db.year, da.month, _days_in_month(db.year, da.month))
    if anchor > db:
        anchor = dt.date(db.year - 1, da.month, min(da.day, _days_in_month(db.year - 1, da.month)))
    return (db - anchor).days


def days360_eu(a, b) -> int:
    """Excel DAYS360(start, end, TRUE) — European method."""
    da, db = to_date(a), to_date(b)
    d1 = 30 if da.day == 31 else da.day
    d2 = 30 if db.day == 31 else db.day
    return (db.year - da.year) * 360 + (db.month - da.month) * 30 + (d2 - d1)


def days360_us(a, b) -> int:
    """Excel DAYS360(start, end) — US (NASD) method, used inside PRICE."""
    da, db = to_date(a), to_date(b)
    d1, d2 = da.day, db.day
    if d1 == 31:
        d1 = 30
    if d2 == 31 and d1 == 30:
        d2 = 30
    return (db.year - da.year) * 360 + (db.month - da.month) * 30 + (d2 - d1)


def trunc(x, digits: int = 0):
    f = 10 ** digits
    return math.trunc(x * f) / f


def xround(x, digits: int = 0):
    """Excel ROUND — half away from zero, unlike Python's banker's rounding."""
    f = 10 ** digits
    v = x * f
    return (-math.floor(-v + 0.5) if v < 0 else math.floor(v + 0.5)) / f


def bbg_round(x):
    """The workbook's IF(x - TRUNC(x) > 0.5, ROUNDUP, ROUNDDOWN) idiom."""
    if x == 0:
        return 0.0
    return float(math.ceil(x)) if (x - math.trunc(x)) > 0.5 else float(math.floor(x))


def _coupon_dates(settle, maturity, freq):
    step = 12 // freq
    d, nxt = maturity, None
    guard = 0
    while d > settle and guard < 2000:
        nxt, d = d, edate(d, -step)
        guard += 1
    return d, (nxt if nxt is not None else edate(d, step))


def _coup_num(settle, maturity, freq) -> int:
    step, d, n, guard = 12 // freq, maturity, 0, 0
    while d > settle and guard < 2000:
        n += 1
        d = edate(d, -step)
        guard += 1
    return n


def price(settle, maturity, rate, yld, redemption, freq, basis):
    """Excel PRICE(). basis 0 = 30/360 US, basis 1 = actual/actual."""
    pcd, ncd = _coupon_dates(settle, maturity, freq)
    n = _coup_num(settle, maturity, freq)
    if basis == 0:
        e = 360.0 / freq
        a = days360_us(pcd, settle)
        dsc = e - a
    else:
        e = float(ncd - pcd)
        a = float(settle - pcd)
        dsc = float(ncd - settle)
    cpn = 100.0 * rate / freq
    y = 1.0 + yld / freq
    p = redemption / y ** (n - 1 + dsc / e)
    for k in range(1, n + 1):
        p += cpn / y ** (k - 1 + dsc / e)
    return p - cpn * a / e


def xyield(settle, maturity, rate, pr, redemption, freq, basis):
    """Excel YIELD(), solved by bisection."""
    if maturity <= settle or pr <= 0:
        return None
    lo, hi = -0.99, 10.0
    for _ in range(200):
        mid = (lo + hi) / 2
        if price(settle, maturity, rate, mid, redemption, freq, basis) > pr:
            lo = mid
        else:
            hi = mid
    return (lo + hi) / 2


# ===========================================================================
# Product metadata
# ===========================================================================
@dataclass
class Meta:
    code: str
    issued: int
    maturity: int
    coupon: float
    cday: int
    m1: int
    m2: int
    first_accrual: int | None
    first_coupon: int | None
    prefix: str
    frequency: str
    freq_months: int
    currency: str
    unit_value: int


def _build_index():
    idx = {}
    for code, issued, maturity, coupon, cday, m1, m2, fad, fcd in PRODUCTS_RAW:
        prefix = code[:2]
        semi = prefix in ("FR", "IN", "PB", "US")          # D208
        currency = "USD" if prefix in ("IN", "US") else "IDR"   # D8
        idx[code] = Meta(
            code=code, issued=issued, maturity=maturity, coupon=coupon,
            cday=cday, m1=m1, m2=m2, first_accrual=fad, first_coupon=fcd,
            prefix=prefix,
            frequency="Semi Annually" if semi else "Monthly",
            freq_months=6 if semi else 1,                   # D270
            currency=currency,
            unit_value=1_000_000 if currency == "IDR" else 1000,  # D217
        )
    return idx


PRODUCT_INDEX = _build_index()
PRODUCT_CODES = sorted(PRODUCT_INDEX)


def product_meta(code):
    return PRODUCT_INDEX.get((code or "").strip().upper())


class SimError(Exception):
    """Raised when the inputs cannot produce a meaningful simulation."""


# ===========================================================================
# Shared blocks
# ===========================================================================
def _bracket_coupons(meta: Meta, settlement: int):
    """
    Coupon dates bracketing a settlement date, walked from the anniversary
    month/day in the product table. Mirrors D306:D310, D346:D350, D471:D475.
    """
    fm = meta.freq_months
    d = to_date(settlement)
    anchor = xdate(d.year, d.month, meta.cday) if fm == 1 else xdate(d.year, meta.m1, meta.cday)
    nxt = edate(anchor, fm)
    if anchor > settlement:
        last = edate(anchor, -fm)
    else:
        last = anchor if nxt > settlement else nxt
    if fm != 1 and last > settlement:
        last = edate(last, -fm)
    return last, edate(last, fm)


def _first_coupon_window(meta: Meta, settlement: int) -> bool:
    """
    D211/D209/D365. Note the sheet tests only the prefix and the date window —
    it deliberately ignores the "New Issuance Obligasi Sekunder?" column,
    which is display-only.
    """
    return (meta.prefix == "IN"
            and meta.first_accrual is not None and meta.first_coupon is not None
            and meta.first_accrual < settlement < meta.first_coupon)


def _resolve_coupons(meta: Meta, settlement: int, is_perdana: bool):
    if is_perdana:
        anchor = xdate(to_date(settlement).year, meta.m1, meta.cday)
        return None, (edate(anchor, 12) if anchor < settlement else anchor)
    if _first_coupon_window(meta, settlement):
        return meta.first_accrual, meta.first_coupon
    return _bracket_coupons(meta, settlement)


def _accrued_block(meta: Meta, last_coupon, next_coupon, settlement,
                   nominal, px, is_perdana, round_accrual=True):
    """Accrued interest. Mirrors D219:D223 (Beli), D217:D221 & D243:D247 (Jual)."""
    units = nominal / meta.unit_value
    per_period = 12 if meta.freq_months == 1 else 2

    if is_perdana:                                                    # D219 / D243
        days = 0
    elif meta.currency == "IDR" or meta.prefix == "US":
        days = settlement - last_coupon
    else:
        days = days360_eu(last_coupon, settlement)

    if is_perdana:                                                    # D220 / D244
        per_unit = 0.0
    elif meta.currency == "IDR" or meta.prefix == "US":
        per_unit = (meta.unit_value * meta.coupon / per_period
                    * (settlement - last_coupon) / (next_coupon - last_coupon))
    else:
        raw = days360_eu(last_coupon, settlement) / 360 * nominal * meta.coupon
        per_unit = xround(raw, 2) if round_accrual else raw

    rounded = bbg_round(per_unit) if meta.currency == "IDR" else per_unit   # D221 / D245

    if meta.currency == "IDR":                                        # D222 / D246
        accrued = units * rounded
    elif meta.prefix == "US":
        accrued = units * (xround(per_unit, 2) if round_accrual else per_unit)
    else:
        accrued = per_unit

    gross = nominal * px + accrued
    total = trunc(gross) if meta.currency == "IDR" else gross         # D223 / D247
    return dict(units=units, days=days, per_unit=per_unit, rounded=rounded,
                accrued=accrued, total=total)


def _coupon_stream(meta: Meta, *, last_coupon, next_coupon, second_coupon,
                   settlement, nominal, units, accrued, horizon,
                   is_perdana, months_horizon=None, net_check=None):
    """Coupon economics. Mirrors D242:D273 (Beli) and D269:D299 (Jual)."""
    fm, cur, pref, cpn_rate = meta.freq_months, meta.currency, meta.prefix, meta.coupon
    per_period = 12 if fm == 1 else 2

    span = datedif_m(edate(next_coupon, -fm), horizon)                # D242 / D269
    periods = span if (cur == "IDR" and fm == 1) else int(xround(span / 6))
    months_invested = datedif_m(settlement, horizon if months_horizon is None else months_horizon)

    base = meta.unit_value * cpn_rate / per_period                    # D244 / D271
    if cur == "IDR":
        per_unit = bbg_round(base)
    elif pref == "US":
        per_unit = xround(base, 2)
    else:
        per_unit = nominal * cpn_rate / per_period

    gross_coupon = units * per_unit if cur == "IDR" or pref == "US" else per_unit  # D245
    coupon_tax = gross_coupon * TAX if cur == "IDR" or pref == "US" else 0.0       # D246
    coupon_tax_r = xround(coupon_tax)                                              # D247

    short_per_unit = 0.0                                              # D248 / D275
    if not is_perdana:
        if cur == "IDR":
            short_per_unit = bbg_round(meta.unit_value * cpn_rate / (12 / fm)
                                       * (next_coupon - settlement) / (next_coupon - last_coupon))
        elif pref == "US":
            short_per_unit = per_unit * (next_coupon - settlement) / (next_coupon - last_coupon)
    short_gross = units * (xround(short_per_unit, 2) if pref == "US" else short_per_unit)
    short_tax_r = xround(short_gross * TAX)                           # D251 / D278

    one_back = edate(next_coupon, -fm)                                # D260 / D287
    two_back = edate(one_back, -fm)                                   # D261 / D288

    first_gap = (next_coupon - settlement) if is_perdana else (next_coupon - last_coupon)
    ctype = "LONG COUPON" if first_gap > second_coupon - next_coupon else "SHORT COUPON"  # D265

    ipo_short = ipo_long = 0.0                                        # D253 / D257
    if is_perdana:
        if ctype == "SHORT COUPON":
            ipo_short = bbg_round(cpn_rate * meta.unit_value / 12
                                  * (next_coupon - settlement) / (next_coupon - one_back))
        else:
            ipo_long = bbg_round(cpn_rate * meta.unit_value / 12
                                 * (one_back - settlement) / (one_back - two_back))
    ipo_short_gross = units * ipo_short                               # D254 / D281
    ipo_short_tax_r = xround(ipo_short_gross * TAX)                   # D256 / D283

    long_per_unit = 0.0                                               # D258 / D285
    if not is_perdana and ctype == "LONG COUPON":
        if cur == "IDR":
            long_per_unit = bbg_round(meta.unit_value * cpn_rate / (12 / fm)
                                      * (one_back - last_coupon) / (one_back - two_back))
        elif pref == "US":
            long_per_unit = (meta.unit_value * cpn_rate / (12 / fm)
                             * (one_back - last_coupon) / (one_back - two_back))
        else:
            long_per_unit = xround(days360_eu(last_coupon, one_back) / 360 * nominal * cpn_rate, 2)

    long_tax_per_unit = 0.0                                           # D259 / D286
    if not is_perdana and long_per_unit != 0 and ctype == "LONG COUPON":
        if cur == "IDR":
            long_tax_per_unit = bbg_round(meta.unit_value * cpn_rate / (12 / fm)
                                          * (one_back - settlement) / (one_back - two_back))
        elif pref == "US":
            long_tax_per_unit = (meta.unit_value * cpn_rate / (12 / fm)
                                 * (one_back - settlement) / (one_back - two_back))

    long_gross = 0.0                                                  # D262 / D289
    if ctype == "LONG COUPON":
        if is_perdana:
            long_gross = units * (ipo_long + per_unit)
        elif cur == "IDR" or pref == "US":
            long_gross = units * (long_per_unit + per_unit)
        else:
            long_gross = long_per_unit + per_unit

    if is_perdana:                                                    # D263 / D290
        long_tax = long_gross * TAX
    elif long_tax_per_unit == 0:
        long_tax = short_gross * TAX
    elif cur == "IDR" or pref == "US":
        long_tax = units * (long_tax_per_unit + per_unit) * TAX
    else:
        long_tax = 0.0
    long_tax_r = xround(long_tax)                                     # D264 / D291

    if cur == "IDR":                                                  # D266 / D293
        if is_perdana:
            first_net = long_gross - long_tax_r if ctype == "LONG COUPON" else ipo_short_gross - ipo_short_tax_r
        else:
            first_net = long_gross - long_tax_r if ctype == "LONG COUPON" else gross_coupon - short_tax_r
    elif ctype == "LONG COUPON":
        first_net = long_gross - long_tax_r if pref == "US" else long_gross
    else:
        first_net = gross_coupon - short_tax_r if pref == "US" else gross_coupon

    first_gross = gross_coupon                                        # D267 / D294
    if cur == "IDR" and is_perdana:
        first_gross = long_gross if ctype == "LONG COUPON" else ipo_short_gross

    # D268 / D295 — note D295 tests the hold-to-maturity counts, not the sale ones.
    check_periods, check_months = net_check if net_check else (periods, months_invested)
    if check_periods == 1 and check_months < fm:
        per_period_net = first_net
    elif cur == "IDR" or pref == "US":
        per_period_net = gross_coupon - coupon_tax_r
    else:
        per_period_net = gross_coupon

    total_net = 0.0 if periods == 0 else (periods - 1) * per_period_net + first_net - accrued  # D271
    total_gross = (periods - 1) * gross_coupon + first_gross                                    # D273

    return dict(periods=periods, months_invested=months_invested, per_unit=per_unit,
                gross_coupon=gross_coupon, coupon_tax=coupon_tax, coupon_tax_r=coupon_tax_r,
                short_tax_r=short_tax_r, type=ctype, first_net=first_net,
                first_gross=first_gross, per_period_net=per_period_net,
                total_net=total_net, total_gross=total_gross)


def _schedule(start, freq_months, periods, first_amount, later_amount, stop):
    rows, d = [], start
    for i in range(min(periods, 400)):
        rows.append((i + 1, d, first_amount if i == 0 else later_amount))
        if d >= stop:
            break
        d = edate(d, freq_months)
    return rows


# ===========================================================================
# SIMULASI BELI
# ===========================================================================
def simulate_beli(code, market, settlement, nominal, px):
    meta = product_meta(code)
    if meta is None:
        raise SimError(f"Kode {code} tidak ada di database produk.")
    if nominal is None or px is None or nominal <= 0 or px <= 0:
        raise SimError("Nilai nominal dan harga harus lebih besar dari nol.")
    if settlement >= meta.maturity:
        raise SimError("Tanggal setelmen sudah melewati tanggal jatuh tempo.")

    is_perdana = market == "Pasar Perdana"
    last_c, next_c = _resolve_coupons(meta, settlement, is_perdana)
    second_c = edate(next_c, meta.freq_months)                        # D213

    acc = _accrued_block(meta, last_c, next_c, settlement, nominal, px, is_perdana)

    basis = 1 if meta.currency == "IDR" else 0                        # D236
    yfreq = 4 if meta.freq_months == 1 else 2                         # D237
    ytm = meta.coupon if is_perdana else xyield(
        settlement, meta.maturity, meta.coupon, px * 100, 100, yfreq, basis)  # D19

    s = _coupon_stream(meta, last_coupon=last_c, next_coupon=next_c, second_coupon=second_c,
                       settlement=settlement, nominal=nominal, units=acc["units"],
                       accrued=acc["accrued"], horizon=meta.maturity, is_perdana=is_perdana)

    capital_gain = xround((1 - px) * nominal)                         # D276
    cg_tax = xround(TAX * capital_gain) if meta.currency == "IDR" else 0.0   # D277
    last_cpn_tax = (s["short_tax_r"] if (s["periods"] == 1 and s["months_invested"] < meta.freq_months)
                    else s["coupon_tax_r"]) if meta.currency == "IDR" else s["coupon_tax"]  # D278
    total_tax = trunc(0 if cg_tax + last_cpn_tax < 0 else cg_tax + last_cpn_tax)             # D279
    final_cpn_net = s["gross_coupon"] - total_tax                                            # D280

    coupon_adj = 0.0 if s["periods"] == 0 else (                                             # D272
        (s["periods"] - 2) * s["per_period_net"] + s["first_net"] - acc["accrued"] + final_cpn_net)

    return dict(
        meta=meta, is_perdana=is_perdana, last_coupon=last_c, next_coupon=next_c,
        accrued_days=acc["days"], accrued_interest=acc["accrued"], amount_paid=acc["total"],
        ytm=ytm, units=acc["units"], periods=s["periods"], months=s["months_invested"],
        coupon_to_maturity=s["total_net"],                            # D271 / J19
        proceeds_at_maturity=nominal + coupon_adj,                    # D274 / J20
        received_at_maturity=final_cpn_net + nominal,                 # J21
        principal_back=nominal, final_coupon_gross=s["gross_coupon"],
        capital_gain=capital_gain, capital_gain_tax=cg_tax,
        last_coupon_tax=last_cpn_tax, total_tax=total_tax,
        cumulative_nominal=nominal - nominal * px + coupon_adj,       # D283
        cumulative_percent=(nominal - nominal * px + coupon_adj) / (nominal * px),  # D285
        schedule=_schedule(next_c, meta.freq_months, s["periods"],
                           s["first_net"], s["per_period_net"], meta.maturity),
    )


# ===========================================================================
# SIMULASI JUAL
# ===========================================================================
def simulate_jual(code, buy_settlement, nominal, buy_price, sell_settlement, sell_price):
    meta = product_meta(code)
    if meta is None:
        raise SimError(f"Kode {code} tidak ada di database produk.")
    if None in (nominal, buy_price, sell_price) or nominal <= 0 or buy_price <= 0 or sell_price <= 0:
        raise SimError("Nominal dan harga harus lebih besar dari nol.")
    if sell_settlement <= buy_settlement:
        raise SimError("Setelmen jual harus setelah setelmen beli.")
    if sell_settlement > meta.maturity:
        raise SimError("Setelmen jual melewati tanggal jatuh tempo.")

    # Buy leg. D209 infers the primary market from settlement == issue date.
    is_perdana = buy_settlement == meta.issued
    buy_last, buy_next = _resolve_coupons(meta, buy_settlement, is_perdana)
    buy_second = edate(buy_next, meta.freq_months)
    buy_acc = _accrued_block(meta, buy_last, buy_next, buy_settlement, nominal, buy_price, is_perdana)

    # Sell leg — D235:D247.
    if _first_coupon_window(meta, sell_settlement):
        sell_last, sell_next = meta.first_accrual, meta.first_coupon
    else:
        sell_last, sell_next = _bracket_coupons(meta, sell_settlement)
    sell_acc = _accrued_block(meta, sell_last, sell_next, sell_settlement, nominal, sell_price, False)

    holding = (sell_settlement - sell_last) if buy_settlement < sell_last else (sell_settlement - buy_settlement)  # D242
    capital_gain = xround((sell_price - buy_price) * nominal)         # D250
    cg_tax = xround(TAX * capital_gain)                               # D251
    accrual_for_tax = xround(                                         # D252
        (sell_acc["rounded"] if meta.currency == "IDR" else sell_acc["per_unit"])
        * (holding / sell_acc["days"] if sell_acc["days"] and holding < sell_acc["days"] else 1))
    accrual_tax = xround(TAX * (sell_acc["units"] * accrual_for_tax    # D253
                                if meta.currency == "IDR" else accrual_for_tax))
    total_tax = trunc(0 if cg_tax + accrual_tax < 0 else cg_tax + accrual_tax) if meta.currency == "IDR" else 0.0  # D254
    net_proceeds = (trunc(sell_acc["total"] if total_tax < 0 else sell_acc["total"] - total_tax)   # D255
                    if meta.currency == "IDR" else sell_acc["total"])

    basis = 1 if meta.currency == "IDR" else 0
    yfreq = 4 if meta.freq_months == 1 else 2
    ytm_hold = xyield(buy_settlement, meta.maturity, meta.coupon, buy_price * 100, 100, yfreq, basis)   # D231
    ytm_sell = xyield(buy_settlement, sell_settlement, meta.coupon,
                      buy_price * 100, sell_price * 100, yfreq, basis)                                  # D265

    common = dict(last_coupon=buy_last, next_coupon=buy_next, second_coupon=buy_second,
                  settlement=buy_settlement, nominal=nominal, units=buy_acc["units"],
                  accrued=buy_acc["accrued"], is_perdana=is_perdana)
    to_mat = _coupon_stream(meta, horizon=meta.maturity, **common)      # D319:D321
    to_sale = _coupon_stream(meta, horizon=sell_last, months_horizon=sell_settlement,
                             net_check=(to_mat["periods"], to_mat["months_invested"]), **common)  # D269:D299

    coupon_to_sale = (sell_acc["accrued"] - buy_acc["accrued"] - total_tax    # D298
                      if to_sale["periods"] == 0 else
                      (to_sale["periods"] - 1) * to_sale["per_period_net"] + to_sale["first_net"]
                      - buy_acc["accrued"] + sell_acc["accrued"])

    gain_sold = (net_proceeds - nominal * buy_price + coupon_to_sale     # D307
                 - sell_acc["accrued"] + total_tax)

    cg_mat = xround((1 - buy_price) * nominal)                          # D324
    cg_mat_tax = xround(TAX * cg_mat) if meta.currency == "IDR" else 0.0  # D326
    last_cpn_tax = ((to_mat["short_tax_r"] if (to_mat["periods"] == 1
                     and to_mat["months_invested"] < meta.freq_months) else to_mat["coupon_tax_r"])
                    if meta.currency == "IDR" else to_mat["coupon_tax"])  # D327
    total_tax_mat = trunc(0 if cg_mat_tax + last_cpn_tax < 0 else cg_mat_tax + last_cpn_tax)  # D328
    final_cpn_net = to_mat["gross_coupon"] - total_tax_mat                                     # D329
    coupon_to_mat = 0.0 if to_mat["periods"] == 0 else (                                       # D321
        (to_mat["periods"] - 2) * to_mat["per_period_net"] + to_mat["first_net"]
        - buy_acc["accrued"] + final_cpn_net)

    return dict(
        meta=meta, is_perdana=is_perdana,
        buy_last=buy_last, buy_next=buy_next, buy_accrued_days=buy_acc["days"],
        buy_accrued=buy_acc["accrued"], amount_paid=buy_acc["total"],
        sell_last=sell_last, sell_next=sell_next, sell_accrued_days=sell_acc["days"],
        sell_accrued=sell_acc["accrued"],
        capital_gain=capital_gain, total_tax=total_tax, net_proceeds=net_proceeds,
        ytm_hold=ytm_hold, ytm_sell=ytm_sell,
        coupon_to_sale=coupon_to_sale,
        proceeds_if_sold=net_proceeds + coupon_to_sale - sell_acc["accrued"],   # D305
        months_if_sold=to_sale["months_invested"],
        gain_if_sold=gain_sold, pct_if_sold=gain_sold / (nominal * buy_price),  # D309
        proceeds_if_held=nominal + coupon_to_mat,                               # D330
        months_if_held=to_mat["months_invested"],
        gain_if_held=nominal - nominal * buy_price + coupon_to_mat,             # D332
        pct_if_held=(nominal - nominal * buy_price + coupon_to_mat) / (nominal * buy_price),  # D334
        schedule=_schedule(buy_next, meta.freq_months, to_sale["periods"],
                           to_sale["first_net"], to_sale["per_period_net"], sell_last),
    )


# ===========================================================================
# SIMULASI SWITCHING
# ===========================================================================
def simulate_switching(code1, buy_settlement1, nominal1, buy_price1,
                       sell_settlement1, sell_price1,
                       code2, settlement2, maturity2, nominal2, price2):
    try:
        p1 = simulate_jual(code1, buy_settlement1, nominal1, buy_price1,
                           sell_settlement1, sell_price1)
    except SimError as exc:
        raise SimError(f"Produk 1: {exc}") from exc

    meta2 = product_meta(code2)
    if meta2 is None:
        raise SimError(f"Kode Produk 2 ({code2}) tidak ada di database produk.")
    if nominal2 is None or price2 is None or nominal2 <= 0 or price2 <= 0:
        raise SimError("Nominal dan harga Produk 2 harus lebih besar dari nol.")
    if maturity2 <= settlement2:
        raise SimError("Jatuh tempo Produk 2 harus setelah tanggal setelmennya.")
    if maturity2 <= buy_settlement1:
        raise SimError("Jatuh tempo Produk 2 harus setelah setelmen beli Produk 1 — "
                       "lama investasi tidak dapat dihitung.")

    is_perdana2 = settlement2 == meta2.issued                          # D365
    last2, next2 = _resolve_coupons(meta2, settlement2, is_perdana2)
    # D374 omits the 2-decimal rounding that D220 and D218 apply.
    acc2 = _accrued_block(meta2, last2, next2, settlement2, nominal2, price2,
                          is_perdana2, round_accrual=False)

    basis2 = 1 if meta2.currency == "IDR" else 0
    yfreq2 = 4 if meta2.freq_months == 1 else 2
    ytm2 = xyield(settlement2, maturity2, meta2.coupon, price2 * 100, 100, yfreq2, basis2)  # D388

    span2 = datedif_m(edate(next2, -meta2.freq_months), maturity2)      # D391
    periods2 = span2 if meta2.freq_months == 1 else span2 // 6
    months2 = datedif_m(settlement2, maturity2)                        # D392

    per_period2 = 12 if meta2.freq_months == 1 else 2
    base2 = meta2.unit_value * meta2.coupon / per_period2              # D401
    if meta2.currency == "IDR":
        cpu2 = bbg_round(base2)
    elif meta2.prefix == "US":
        cpu2 = xround(base2, 2)
    else:
        cpu2 = nominal2 * meta2.coupon / per_period2
    gross2 = acc2["units"] * cpu2 if meta2.currency == "IDR" else cpu2  # D402
    tax2_r = xround(gross2 * TAX) if meta2.currency == "IDR" else 0.0   # D404
    net2 = gross2 - tax2_r                                              # D425

    coupon_to_horizon2 = (acc2["accrued"] if periods2 == 0                # D428
                          else (periods2 - 1) * net2 + gross2 - acc2["accrued"])

    capital = nominal1 * buy_price1                                     # D453
    gain_sell1 = p1["net_proceeds"] - capital                           # D455
    gain_buy2 = p1["net_proceeds"] - acc2["total"]                      # D459
    gain_switch = gain_sell1 + p1["coupon_to_sale"] + gain_buy2 + coupon_to_horizon2  # D463

    return dict(
        p1=p1, meta2=meta2, is_perdana2=is_perdana2,
        currency_mismatch=p1["meta"].currency != meta2.currency,
        last_coupon2=last2, next_coupon2=next2,
        accrued_days2=acc2["days"], accrued2=acc2["accrued"], amount_paid2=acc2["total"],
        ytm2=ytm2, coupon_to_horizon2=coupon_to_horizon2,
        periods2=periods2, months2=months2,
        months_switched=datedif_m(buy_settlement1, maturity2),          # D452
        capital=capital, proceeds_sell1=p1["net_proceeds"], gain_sell1=gain_sell1,
        coupon_until_sale1=p1["coupon_to_sale"], paid2=acc2["total"], gain_buy2=gain_buy2,
        total_return_switch=gain_buy2 + nominal1 + coupon_to_horizon2,  # D461
        gain_switch=gain_switch, pct_switch=gain_switch / capital,      # D465
        top_up=p1["net_proceeds"] - acc2["total"],
        schedule2=_schedule(next2, meta2.freq_months, periods2, gross2, net2, maturity2),
    )


# ===========================================================================
# SIMULASI KREDIT (gadai obligasi)
# Compares the bond's annual coupon income against the annual cost of a loan
# taken against the bond as collateral. This is a fresh feature, not a port.
# ===========================================================================
def default_ltv(meta) -> float:
    """
    LTV tiers by currency and remaining maturity (years from today):
      IDR : <5yr 90% · 5-10yr 85% · >10yr 80%
      USD : <10yr 80% · >=10yr 70%
    """
    today = to_serial(dt.date.today())
    years = max(0.0, (meta.maturity - today) / 365.25)
    if meta.currency == "IDR":
        if years < 5:
            return 0.90
        if years <= 10:
            return 0.85
        return 0.80
    # USD
    return 0.80 if years < 10 else 0.70


def coupon_tax_rate(meta) -> float:
    """INDON/INDOIS (USD sovereign) coupons are tax-free here; IDR bonds 10%."""
    return 0.0 if meta.currency == "USD" else 0.10


def simulate_kredit(code, nominal, ltv, loan_rate, tenor_years,
                    provision_pct, admin_fee):
    meta = product_meta(code)
    if meta is None:
        raise SimError(f"Kode {code} tidak ada di database produk.")
    if nominal is None or nominal <= 0:
        raise SimError("Nominal investasi harus lebih besar dari nol.")
    if ltv is None or not (0 < ltv <= 1):
        raise SimError("LTV harus di antara 0% dan 100%.")
    if loan_rate is None or loan_rate < 0:
        raise SimError("Suku bunga kredit tidak boleh negatif.")
    if tenor_years is None or tenor_years <= 0:
        raise SimError("Periode angsuran harus lebih besar dari nol.")

    cur = meta.currency
    tax = coupon_tax_rate(meta)

    # --- collateral / loan sizing ---
    plafon = nominal * ltv                                    # Plafon Kredit

    # --- investment income (bond coupon, nett of tax) ---
    coupon_gross_yr = nominal * meta.coupon
    coupon_net_yr = coupon_gross_yr * (1 - tax)
    inv_month = coupon_net_yr / 12
    inv_day = coupon_net_yr / 365

    # --- loan cost (interest on the full plafon, worst case) ---
    loan_interest_yr = plafon * loan_rate
    loan_month = loan_interest_yr / 12
    loan_day = loan_interest_yr / 365

    # --- annuity installment on the plafon over the tenor ---
    n = int(round(tenor_years * 12))
    r_m = loan_rate / 12
    if r_m == 0:
        annuity = plafon / n
    else:
        annuity = plafon * r_m / (1 - (1 + r_m) ** (-n))
    total_repay = annuity * n
    total_loan_interest = total_repay - plafon

    # --- upfront fees ---
    provision = plafon * (provision_pct or 0)
    total_fees = provision + (admin_fee or 0)

    # --- net carry (the headline the RM is showing) ---
    net_carry_yr = coupon_net_yr - loan_interest_yr
    net_carry_month = net_carry_yr / 12

    # --- tenor vs maturity guard: facility must end >= 1 month before JT ---
    today = to_serial(dt.date.today())
    loan_end = edate(today, n)
    latest_allowed = edate(meta.maturity, -1)
    tenor_ok = loan_end <= latest_allowed
    max_tenor_months = max(0, datedif_m(today, latest_allowed))

    return dict(
        meta=meta, currency=cur, tax=tax,
        nominal=nominal, ltv=ltv, plafon=plafon, loan_rate=loan_rate,
        tenor_years=tenor_years, tenor_months=n,
        coupon_gross_yr=coupon_gross_yr, coupon_net_yr=coupon_net_yr,
        inv_month=inv_month, inv_day=inv_day,
        loan_interest_yr=loan_interest_yr, loan_month=loan_month, loan_day=loan_day,
        annuity=annuity, total_repay=total_repay, total_loan_interest=total_loan_interest,
        provision=provision, admin_fee=(admin_fee or 0), total_fees=total_fees,
        net_carry_yr=net_carry_yr, net_carry_month=net_carry_month,
        loan_end=loan_end, latest_allowed=latest_allowed,
        tenor_ok=tenor_ok, max_tenor_months=max_tenor_months,
    )


# ===========================================================================
# Presentation helpers
# ===========================================================================
def money(value, currency="IDR", decimals=None):
    if value is None or not isinstance(value, (int, float)) or not math.isfinite(value):
        return "—"
    d = decimals if decimals is not None else (0 if currency == "IDR" else 2)
    sign = "-" if value < 0 else ""
    body = f"{abs(value):,.{d}f}".replace(",", "\u00a0")
    symbol = "Rp" if currency == "IDR" else "$"
    return f"{sign}{symbol} {body}"


def pct(value, decimals=2):
    if value is None or not isinstance(value, (int, float)) or not math.isfinite(value):
        return "—"
    return f"{value * 100:,.{decimals}f}%".replace(",", "\u00a0")


def num(value, decimals=0):
    if value is None:
        return "—"
    return f"{value:,.{decimals}f}".replace(",", "\u00a0")


def parse_number(text):
    """Accept 2.000.000, 2,000,000, 2000000 and 1234.56 alike."""
    if text is None:
        return None
    s = str(text).strip().replace(" ", "").replace("\u00a0", "")
    if not s:
        return None
    if "," in s and "." in s:
        s = s.replace("." if s.rfind(",") > s.rfind(".") else ",", "")
        s = s.replace(",", ".")
    elif s.count(".") > 1:
        s = s.replace(".", "")
    elif "," in s:
        s = s.replace(",", ".") if len(s.split(",")[-1]) != 3 else s.replace(",", "")
    try:
        return float(s)
    except ValueError:
        return None


def _full_width():
    """
    Return the kwargs that make a widget span its container, tolerant of the
    Streamlit version. Newer builds prefer width='stretch'; older ones only
    understand use_container_width=True. Picking at runtime avoids a hard
    break whichever version Streamlit Cloud resolves.
    """
    import inspect
    try:
        params = inspect.signature(st.button).parameters
        if "width" in params:
            return {"width": "stretch"}
    except (ValueError, TypeError):
        pass
    return {"use_container_width": True}


FULL_WIDTH = _full_width()


def money_input(label, default, key, cur_hint="Rp"):
    """
    Text input for a nominal amount that echoes the parsed value back in
    grouped form (Rp 200.000.000) so the RM can confirm the zero count at a
    glance. Returns the raw string; caller runs parse_number.
    """
    raw = st.text_input(label, default, key=key)
    val = parse_number(raw)
    if val is not None:
        st.caption(f"= {cur_hint} {val:,.0f}".replace(",", "."))
    elif raw.strip():
        st.caption("⚠️ Angka tidak dikenali")
    return raw


# ---------------------------------------------------------------------------
# PNG export. Drawn server-side with matplotlib so the download works
# identically on every browser and needs no client-side capture library.
# Only the fonts bundled with matplotlib are used — nothing is fetched.
# ---------------------------------------------------------------------------
INK = "#0f2233"
MUTED = "#5c7085"
LINE = "#dbe3ea"
ACCENT = "#0e7c6b"
LOSS = "#b4342a"
MARKER = "#ffd400"
SANS = "DejaVu Sans"
MONO = "DejaVu Sans Mono"


def render_png(title, subtitle, sections, footer_lines):
    """sections: list of (heading, [(label, value, tone, indent), ...])."""
    line_h = 0.235
    # Height is computed from the exact drawing constants below, not estimated,
    # so long simulations never clip the disclaimer off the bottom.
    head_h = 0.55 + 0.34 + 0.26 + 0.34
    body_h = sum(0.20 + 0.10 + len(body) * line_h + 0.24 for _, body in sections)
    foot_h = 0.05 + 0.22 + 0.20 + len(footer_lines) * 0.175
    height = head_h + body_h + foot_h + 0.30
    width = 8.6

    fig = plt.figure(figsize=(width, height), dpi=200)
    fig.patch.set_facecolor("white")
    ax = fig.add_axes([0, 0, 1, 1])
    ax.set_xlim(0, width)
    ax.set_ylim(0, height)
    ax.axis("off")

    left, right = 0.55, width - 0.55
    y = height - 0.55

    ax.add_patch(plt.Rectangle((left, y - 0.30), 0.075, 0.42, color=MARKER, zorder=3))
    ax.text(left + 0.20, y, "KALKULATOR OBLIGASI", fontsize=7.5, color=MUTED,
            family=SANS, weight="bold", va="center")
    y -= 0.34
    ax.text(left + 0.20, y, title.upper(), fontsize=15, color=INK, family=SANS,
            weight="bold", va="center")
    ax.text(right, y, subtitle, fontsize=8, color=MUTED, family=MONO,
            ha="right", va="center")
    y -= 0.26
    ax.plot([left, right], [y, y], color=INK, lw=1.4)
    y -= 0.34

    for heading, body in sections:
        ax.text(left, y, heading.upper(), fontsize=7.5, color=MUTED,
                family=SANS, weight="bold", va="center")
        y -= 0.20
        ax.plot([left, right], [y + 0.04, y + 0.04], color=LINE, lw=0.7)
        y -= 0.10
        for label, value, tone, indent in body:
            colour = {"gain": ACCENT, "loss": LOSS, "strong": INK}.get(tone, INK)
            weight = "bold" if tone == "strong" else "normal"
            ax.text(left + (0.22 if indent else 0), y, ("– " if indent else "") + label,
                    fontsize=8.5, color=MUTED if indent else INK, family=SANS, va="center")
            ax.text(right, y, value, fontsize=9, color=colour, family=MONO,
                    weight=weight, ha="right", va="center")
            y -= line_h
        y -= 0.24

    y -= 0.05
    ax.plot([left, right], [y, y], color=LINE, lw=0.7)
    y -= 0.22
    ax.text(left, y, "DISCLAIMER", fontsize=6.5, color=MUTED, family=SANS,
            weight="bold", va="center")
    y -= 0.20
    for text in footer_lines:
        ax.text(left, y, "• " + text, fontsize=6, color=MUTED, family=SANS, va="center")
        y -= 0.175

    buf = io.BytesIO()
    FigureCanvasAgg(fig).print_png(buf)
    plt.close(fig)
    return buf.getvalue()


# ---------------------------------------------------------------------------
# Faithful "Simulasi Beli" export — reproduces the original workbook layout:
# yellow product banner, black header bars, the two-panel split, the numbered
# coupon schedule (collapsed to 1..7 + final for long tenors), the keterangan
# notes, and the full disclaimer box.
# ---------------------------------------------------------------------------
from matplotlib.patches import FancyBboxPatch  # noqa: E402

_EX_BLACK = "#111111"
_EX_HAIR = "#e4eaef"
_EX_LINE = "#c9d3db"
_EX_MARKER = "#ffe14d"
_EX_MARKER_EDGE = "#e8c400"
_EX_PANEL = "#f6f9fb"

DISCLAIMER_FULL = [
    "Kalkulator ini disediakan hanya sebagai alat bantu simulasi Obligasi dan tidak dimaksudkan untuk menyediakan rekomendasi atau saran apa pun.",
    "Simulasi, harga dan YTM Obligasi yang ditampilkan hanya bersifat indikatif, sehingga terdapat kemungkinan perbedaan dengan perhitungan, harga dan YTM Obligasi pada saat nasabah melakukan transaksi yang sebenarnya.",
    "YTM merupakan potensi tingkat pengembalian (disetahunkan) yang akan diperoleh jika Obligasi dipegang hingga jatuh tempo, dengan diasumsikan bahwa pembayaran kupon dilakukan sesuai jadwal dan diinvestasikan kembali pada rate yang sama.",
    "Tarif pajak yang digunakan dalam simulasi kalkulator ini menggunakan tarif pajak Obligasi sebesar 10%.",
    "Perhitungan ini belum dipotong biaya (apabila ada).",
    "Simulasi ini tidak dapat digunakan untuk produk FR0088 & FR0089 yang ditransaksikan pada tanggal 6 -7 Januari 2021",
]


DISCLAIMER_SWITCH = [
    "Kalkulator ini disediakan hanya sebagai alat bantu simulasi Obligasi dan tidak dimaksudkan untuk menyediakan rekomendasi atau saran apa pun.",
    "Simulasi, harga dan YTM Obligasi yang ditampilkan hanya bersifat indikatif, sehingga terdapat kemungkinan perbedaan dengan perhitungan, harga dan YTM Obligasi pada saat nasabah melakukan transaksi yang sebenarnya.",
    "Tarif pajak yang digunakan dalam simulasi kalkulator ini menggunakan tarif pajak Obligasi sebesar 10%.",
    "Perhitungan ini belum dipotong biaya (apabila ada).",
]


def _paren(v):
    """Accounting style — negatives in parentheses, as the workbook shows them."""
    if v is None:
        return "—"
    body = f"{abs(v):,.2f}"
    return f"({body})" if v < 0 else body


def render_beli_export(r, *, code, market, txn_serial, settle_serial, nominal, price_pct,
                       sim_title="SIMULASI BELI"):
    cur = r["meta"].currency
    meta = r["meta"]
    sched = r["schedule"]

    W = 12.4
    disc_h = 0.26 + len(DISCLAIMER_FULL) * 0.208 + 0.14
    title_h = 0.50
    H = 8.55 + disc_h + title_h
    fig = plt.figure(figsize=(W, H), dpi=170)
    fig.patch.set_facecolor("white")
    ax = fig.add_axes([0, 0, 1, 1])
    ax.set_xlim(0, W)
    ax.set_ylim(0, H)
    ax.axis("off")

    ML, MR, MT = 0.35, 0.35, 0.30
    gutter = 0.45
    col_x = ML
    col_w = (W - ML - MR - gutter) / 2
    rcol_x = ML + col_w + gutter
    top = H - MT

    def text(x, y, s, size=9, color=INK, weight="normal", ha="left", va="center", family=SANS):
        ax.text(x, y, s, fontsize=size, color=color, weight=weight, ha=ha, va=va, family=family)

    def bar(x, y, w, h, label):
        ax.add_patch(plt.Rectangle((x, y - h), w, h, facecolor=_EX_BLACK, edgecolor="none", zorder=2))
        text(x + w / 2, y - h / 2, label, size=11, color="white", weight="bold", ha="center")

    # ---- title band (full width) ----
    ax.text(ML, top - 0.02, sim_title, fontsize=17, color=INK, weight="bold",
            family=SANS, ha="left", va="top")
    ax.text(W - MR, top - 0.10, "Kalkulator Obligasi",
            fontsize=9.5, color=MUTED, family=SANS, ha="right", va="top")
    ax.plot([ML, W - MR], [top - title_h + 0.06, top - title_h + 0.06],
            color=INK, lw=1.6)
    top -= title_h + 0.14

    # ---- left column ----
    y = top
    bh = 0.52
    ax.add_patch(FancyBboxPatch((col_x, y - bh), col_w, bh,
                                boxstyle="round,pad=0,rounding_size=0.02",
                                facecolor=_EX_MARKER, edgecolor=_EX_MARKER_EDGE, lw=1.2, zorder=2))
    text(col_x + 0.14, y - bh / 2, code, size=20, weight="bold")
    y -= bh + 0.22

    bar(col_x, y, col_w, 0.34, "Data Produk Obligasi  (NASABAH BELI)")
    y -= 0.34

    left_rows = [
        ("Mata Uang", cur, False),
        ("Jenis Transaksi", market, True),
        ("Tanggal Transaksi", fmt_date_en(txn_serial), True),
        ("Tanggal Setelmen", fmt_date_en(settle_serial), True),
        ("Tanggal Kupon Terakhir", "N/A" if r["is_perdana"] else fmt_date_en(r["last_coupon"]), False),
        ("Tanggal Kupon Berikutnya", fmt_date_en(r["next_coupon"]), False),
        ("Tanggal Jatuh Tempo (JT)", fmt_date_en(meta.maturity), False),
        ("Kupon", pct(meta.coupon, 3), False),
        ("Nilai Nominal", f"{nominal:,.0f}", True),
        ("Harga Nasabah Beli", f"{price_pct:.4f}%", True),
    ]
    rh = 0.285
    for lab, val, hl in left_rows:
        ry = y - rh
        if hl:
            ax.add_patch(plt.Rectangle((col_x + col_w * 0.52, ry), col_w * 0.48, rh,
                                       facecolor=_EX_MARKER, edgecolor="none", zorder=1))
        ax.add_patch(plt.Rectangle((col_x, ry), col_w, rh, facecolor="none",
                                   edgecolor=_EX_LINE, lw=0.6, zorder=1.5))
        text(col_x + 0.10, ry + rh / 2, lab, size=8.5, weight="bold")
        text(col_x + col_w - 0.10, ry + rh / 2, val, size=8.5, weight="bold", ha="right", family=MONO)
        y = ry
    y -= 0.16

    for lab, val in [
        ("Imbal Hasil hingga JT / Yield To Maturity (gross)", pct(r["ytm"], 4)),
        ("Hari Kupon Berjalan / Days of Accrued Interest", num(r["accrued_days"])),
        ("Kupon Berjalan / Accrued Interest", _paren(r["accrued_interest"])),
    ]:
        ry = y - rh
        ax.add_patch(plt.Rectangle((col_x, ry), col_w, rh, facecolor=_EX_PANEL,
                                   edgecolor=_EX_LINE, lw=0.8, zorder=1))
        text(col_x + 0.10, ry + rh / 2, lab, size=8, weight="bold")
        text(col_x + col_w - 0.10, ry + rh / 2, val, size=8.5, weight="bold", ha="right", family=MONO)
        y = ry - 0.06
    y -= 0.14

    tbh = 0.56
    ax.add_patch(plt.Rectangle((col_x, y - tbh), col_w, tbh, facecolor=_EX_BLACK, zorder=2))
    text(col_x + 0.16, y - tbh / 2, "Jumlah Indikatif yang Dibayar", size=12.5, color="white", weight="bold")
    text(col_x + col_w - 0.16, y - tbh / 2, _paren(r["amount_paid"]), size=13,
         color="white", weight="bold", ha="right", family=MONO)
    y -= tbh + 0.18

    text(col_x, y, "Keterangan :", size=8.5)
    y -= 0.26
    ax.add_patch(plt.Rectangle((col_x, y - 0.02), 0.16, 0.16, facecolor=_EX_MARKER,
                               edgecolor=_EX_MARKER_EDGE, lw=0.6))
    text(col_x + 0.24, y + 0.06, "Kolom yang di-highlight kuning WAJIB diisi dengan data terkini.", size=7.5)
    y -= 0.30
    for i, note in enumerate([
        "Kupon yang diterima hingga JT = total kupon (net, tidak termasuk pajak capital gain/loss) - accrued kupon beli",
        "Pengembalian hingga JT = nominal + total kupon (net) yang diterima hingga JT - pajak capital gain/loss",
        "Nominal yang diterima saat JT = pengembalian pokok + kupon saat jatuh tempo (gross) - total pajak",
        "Total keuntungan/kerugian = total kupon (gross) yang diterima hingga JT + capital gain/loss - total pajak",
    ], 1):
        text(col_x + 0.02, y + 0.06, str(i), size=7.5, weight="bold")
        text(col_x + 0.24, y + 0.06, note, size=6.6)
        y -= 0.235

    # ---- right column ----
    y = top
    bar(rcol_x, y, col_w, 0.40, "Proyeksi Pendapatan yang Diterima Nasabah (nett)")
    y -= 0.40 + 0.06

    def rrow(cy, lab, val, *, lab_size=8.8, val_size=8.8, bold_lab=True, bold_val=False,
             col=INK, dot=False, indent=0.0, mid=None, hair=True):
        if hair:
            ax.plot([rcol_x, rcol_x + col_w], [cy, cy], color=_EX_HAIR, lw=0.6)
        yy = cy - 0.155
        prefix = "\u25cf " if dot else ""
        text(rcol_x + 0.06 + indent, yy, prefix + lab, size=lab_size,
             weight="bold" if bold_lab else "normal")
        if mid is not None:
            text(rcol_x + col_w * 0.60, yy, mid, size=val_size, ha="right", family=MONO)
        text(rcol_x + col_w - 0.08, yy, val, size=val_size, ha="right", family=MONO,
             weight="bold" if bold_val else "normal", color=col)
        return cy - 0.31

    y = rrow(y, "Nominal yang dibayar pada tanggal :", _paren(-r["amount_paid"]),
             mid=fmt_date_en(settle_serial), bold_val=True, hair=False)
    y = rrow(y, "Kupon (tidak termasuk pajak capital gain/loss) :", "")

    def sched_row(cy, n, date, amt):
        ax.text(rcol_x + col_w * 0.62, cy - 0.155, fmt_date_en(date), fontsize=8.4,
                family=MONO, ha="left", color=INK, va="center")
        return rrow(cy, "", _paren(amt), mid=str(n), bold_lab=False, val_size=8.4, lab_size=8.4)

    if len(sched) <= 8:
        for n, date, amt in sched:
            y = sched_row(y, n, date, amt)
    elif sched:
        for n, date, amt in sched[:7]:
            y = sched_row(y, n, date, amt)
        ax.plot([rcol_x, rcol_x + col_w], [y, y], color=_EX_HAIR, lw=0.6)
        for gx in (0.12, 0.60, 0.90):
            ax.text(rcol_x + col_w * gx, y - 0.16, "\u22ee", fontsize=9, ha="center",
                    va="center", color=MUTED)
        y -= 0.31
        last = sched[-1]
        y = sched_row(y, last[0], last[1], last[2])

    for lab, val, bold, dot, ind in [
        ("Kupon yang diterima hingga JT1 :", r["coupon_to_maturity"], True, False, 0),
        ("Pengembalian hingga JT2 :", r["proceeds_at_maturity"], True, False, 0),
        ("Nominal yang diterima saat JT3 :", r["received_at_maturity"], True, False, 0),
        ("Pengembalian pokok", r["principal_back"], False, True, 0),
        ("Kupon saat jatuh tempo (gross)", r["final_coupon_gross"], False, True, 0),
        ("Capital gain/loss", r["capital_gain"], False, True, 0),
        ("Total Pajak", r["total_tax"], False, True, 0),
        ("Pajak Kupon", r["last_coupon_tax"], False, False, 0.22),
        ("Pajak capital gain/loss", r["capital_gain_tax"], False, False, 0.22),
    ]:
        y = rrow(y, lab, _paren(val), bold_lab=bold, bold_val=bold, dot=dot, indent=ind)

    ax.plot([rcol_x, rcol_x + col_w], [y, y], color=INK, lw=1.0)
    y = rrow(y, "Total keuntungan/kerugian :", "", bold_lab=True, hair=False)
    y = rrow(y, "Nominal kumulatif", _paren(r["cumulative_nominal"]), dot=True, bold_val=True)
    y = rrow(y, "Persentase kumulatif", pct(r["cumulative_percent"]), dot=True, bold_val=True)

    # ---- disclaimer box, pinned to the bottom margin ----
    box_bottom = 0.28
    box_top = box_bottom + disc_h
    ax.add_patch(plt.Rectangle((ML, box_bottom), W - ML - MR, disc_h,
                               facecolor="white", edgecolor=INK, lw=1.0))
    dy = box_top - 0.20
    text(ML + 0.12, dy, "DISCLAIMER :", size=8, weight="bold")
    dy -= 0.215
    for line in DISCLAIMER_FULL:
        text(ML + 0.16, dy, "\u25cf", size=6)
        text(ML + 0.34, dy, line, size=6.4)
        dy -= 0.208

    buf = io.BytesIO()
    FigureCanvasAgg(fig).print_png(buf)
    plt.close(fig)
    return buf.getvalue()


# ---------------------------------------------------------------------------
# Shared drawing toolkit for the faithful Jual / Switching exports. Same visual
# vocabulary as render_beli_export: yellow code banner, black header bars, a
# left data panel with yellow-highlighted required cells, grey derived rows,
# black total bars, and a right projection panel with a numbered coupon
# schedule. Kept in one place so all three exports stay pixel-consistent.
# ---------------------------------------------------------------------------
class _Sheet:
    def __init__(self, ax, col_x, col_w):
        self.ax = ax
        self.col_x = col_x
        self.col_w = col_w

    def text(self, x, y, s, size=9, color=INK, weight="normal", ha="left",
             va="center", family=SANS):
        self.ax.text(x, y, s, fontsize=size, color=color, weight=weight,
                     ha=ha, va=va, family=family)

    def banner(self, x, y, w, code):
        """Yellow rounded product-code banner."""
        bh = 0.52
        self.ax.add_patch(FancyBboxPatch((x, y - bh), w, bh,
                          boxstyle="round,pad=0,rounding_size=0.02",
                          facecolor=_EX_MARKER, edgecolor=_EX_MARKER_EDGE, lw=1.2, zorder=2))
        self.text(x + 0.14, y - bh / 2, code, size=20, weight="bold")
        return y - bh

    def bar(self, x, y, w, label, h=0.34, size=11):
        self.ax.add_patch(plt.Rectangle((x, y - h), w, h, facecolor=_EX_BLACK,
                          edgecolor="none", zorder=2))
        self.text(x + w / 2, y - h / 2, label, size=size, color="white",
                  weight="bold", ha="center")
        return y - h

    def total_bar(self, x, y, w, label, value):
        h = 0.56
        self.ax.add_patch(plt.Rectangle((x, y - h), w, h, facecolor=_EX_BLACK, zorder=2))
        self.text(x + 0.16, y - h / 2, label, size=12.5, color="white", weight="bold")
        self.text(x + w - 0.16, y - h / 2, value, size=13, color="white",
                  weight="bold", ha="right", family=MONO)
        return y - h

    def data_rows(self, x, y, w, rows, rh=0.285):
        """rows: (label, value, highlight_bool). Yellow fill on highlighted cells."""
        for lab, val, hl in rows:
            ry = y - rh
            if hl:
                self.ax.add_patch(plt.Rectangle((x + w * 0.52, ry), w * 0.48, rh,
                                  facecolor=_EX_MARKER, edgecolor="none", zorder=1))
            self.ax.add_patch(plt.Rectangle((x, ry), w, rh, facecolor="none",
                              edgecolor=_EX_LINE, lw=0.6, zorder=1.5))
            self.text(x + 0.10, ry + rh / 2, lab, size=8.5, weight="bold")
            self.text(x + w - 0.10, ry + rh / 2, val, size=8.5, weight="bold",
                      ha="right", family=MONO)
            y = ry
        return y

    def derived_rows(self, x, y, w, rows, rh=0.285):
        """Grey panel rows for computed values."""
        for lab, val in rows:
            ry = y - rh
            self.ax.add_patch(plt.Rectangle((x, ry), w, rh, facecolor=_EX_PANEL,
                              edgecolor=_EX_LINE, lw=0.8, zorder=1))
            self.text(x + 0.10, ry + rh / 2, lab, size=8, weight="bold")
            self.text(x + w - 0.10, ry + rh / 2, val, size=8.5, weight="bold",
                      ha="right", family=MONO)
            y = ry - 0.06
        return y

    def rrow(self, x, w, cy, lab, val, *, lab_size=8.8, val_size=8.8,
             bold_lab=True, bold_val=False, col=INK, dot=False, indent=0.0,
             mid=None, hair=True):
        if hair:
            self.ax.plot([x, x + w], [cy, cy], color=_EX_HAIR, lw=0.6)
        yy = cy - 0.155
        prefix = "\u25cf " if dot else ""
        self.text(x + 0.06 + indent, yy, prefix + lab, size=lab_size,
                  weight="bold" if bold_lab else "normal")
        if mid is not None:
            self.text(x + w * 0.60, yy, mid, size=val_size, ha="right", family=MONO)
        self.text(x + w - 0.08, yy, val, size=val_size, ha="right", family=MONO,
                  weight="bold" if bold_val else "normal", color=col)
        return cy - 0.31

    def schedule(self, x, w, cy, sched, *, date_frac=0.62):
        """Numbered coupon schedule, collapsed 1..7 + ellipsis + final if long."""
        def one(cy, n, date, amt):
            self.ax.text(x + w * date_frac, cy - 0.155, fmt_date_en(date),
                         fontsize=8.4, family=MONO, ha="left", color=INK, va="center")
            return self.rrow(x, w, cy, "", _paren(amt), mid=str(n),
                             bold_lab=False, val_size=8.4, lab_size=8.4)
        if len(sched) <= 8:
            for n, date, amt in sched:
                cy = one(cy, n, date, amt)
        elif sched:
            for n, date, amt in sched[:7]:
                cy = one(cy, n, date, amt)
            self.ax.plot([x, x + w], [cy, cy], color=_EX_HAIR, lw=0.6)
            for gx in (0.12, 0.60, 0.90):
                self.ax.text(x + w * gx, cy - 0.16, "\u22ee", fontsize=9,
                             ha="center", va="center", color=MUTED)
            cy -= 0.31
            last = sched[-1]
            cy = one(cy, last[0], last[1], last[2])
        return cy


def _new_figure(sim_title, W, H):
    fig = plt.figure(figsize=(W, H), dpi=170)
    fig.patch.set_facecolor("white")
    ax = fig.add_axes([0, 0, 1, 1])
    ax.set_xlim(0, W)
    ax.set_ylim(0, H)
    ax.axis("off")
    MT = 0.30
    top = H - MT
    title_h = 0.50
    ax.text(0.35, top - 0.02, sim_title, fontsize=17, color=INK, weight="bold",
            family=SANS, ha="left", va="top")
    ax.text(W - 0.35, top - 0.10, "Kalkulator Obligasi", fontsize=9.5, color=MUTED,
            family=SANS, ha="right", va="top")
    ax.plot([0.35, W - 0.35], [top - title_h + 0.06, top - title_h + 0.06],
            color=INK, lw=1.6)
    return fig, ax, top - title_h - 0.14


def _disclaimer_box(ax, W, ML, MR, disc_h, lines=None, top=None):
    lines = lines if lines is not None else DISCLAIMER_FULL
    box_top = top if top is not None else (0.28 + disc_h)
    box_bottom = box_top - disc_h
    ax.add_patch(plt.Rectangle((ML, box_bottom), W - ML - MR, disc_h,
                 facecolor="white", edgecolor=INK, lw=1.0))
    dy = box_top - 0.20
    ax.text(ML + 0.12, dy, "DISCLAIMER :", fontsize=8, color=INK, weight="bold",
            family=SANS, va="center")
    dy -= 0.215
    for line in lines:
        ax.text(ML + 0.16, dy, "\u25cf", fontsize=6, color=INK, va="center")
        ax.text(ML + 0.34, dy, line, fontsize=6.4, color=INK, family=SANS, va="center")
        dy -= 0.208


# ---------------------------------------------------------------------------
# SIMULASI JUAL — two stacked data panels (beli + jual) on the left, projection
# with sell-vs-hold branches on the right. Matches the uploaded reference.
# ---------------------------------------------------------------------------
def render_jual_export(r, *, code, buy_txn, buy_settle, nominal, buy_price,
                       sell_txn, sell_settle, sell_price, sim_title="SIMULASI JUAL"):
    cur = r["meta"].currency
    meta = r["meta"]
    disc_h = 0.26 + len(DISCLAIMER_FULL) * 0.208 + 0.14
    W, H = 12.4, 11.35 + disc_h
    fig, ax, top = _new_figure(sim_title, W, H)

    ML, MR, gutter = 0.35, 0.35, 0.45
    col_w = (W - ML - MR - gutter) / 2
    rcol_x = ML + col_w + gutter
    s = _Sheet(ax, ML, col_w)

    # ===== LEFT =====
    y = top
    y = s.banner(ML, y, col_w, code) - 0.22
    y = s.bar(ML, y, col_w, "Data Produk Obligasi  (NASABAH BELI)")
    y = s.data_rows(ML, y, col_w, [
        ("Mata Uang", cur, False),
        ("Tanggal Transaksi", fmt_date_en(buy_txn), True),
        ("Tanggal Setelmen", fmt_date_en(buy_settle), True),
        ("Tanggal Kupon Terakhir", "N/A" if r["is_perdana"] else fmt_date_en(r["buy_last"]), False),
        ("Tanggal Kupon Pertama", fmt_date_en(r["buy_next"]), False),
        ("Tanggal Jatuh Tempo (JT)", fmt_date_en(meta.maturity), False),
        ("Kupon", pct(meta.coupon, 3), False),
        ("Nilai Nominal", f"{nominal:,.0f}", True),
        ("Harga Nasabah Beli", f"{buy_price:.4f}%", True),
    ])
    y -= 0.16
    y = s.derived_rows(ML, y, col_w, [
        ("Hari Kupon Berjalan / Days of Accrued Interest", num(r["buy_accrued_days"])),
        ("Kupon Berjalan / Accrued Interest", _paren(r["buy_accrued"])),
    ])
    y -= 0.14
    y = s.total_bar(ML, y, col_w, "Jumlah Indikatif yang Dibayar", _paren(r["amount_paid"]))
    y -= 0.26

    s.text(ML, y, f"Nasabah menjual Obligasi pada tanggal : {fmt_date_en(sell_settle)}", size=8.5)
    y -= 0.22
    y = s.bar(ML, y, col_w, "Data Produk Obligasi  (NASABAH JUAL)")
    y = s.data_rows(ML, y, col_w, [
        ("Tanggal Transaksi", fmt_date_en(sell_txn), True),
        ("Tanggal Setelmen", fmt_date_en(sell_settle), True),
        ("Tanggal Kupon Terakhir", fmt_date_en(r["sell_last"]), False),
        ("Tanggal Kupon Berikutnya", fmt_date_en(r["sell_next"]), False),
        ("Harga Nasabah Jual", f"{sell_price:.4f}%", True),
    ])
    y -= 0.16
    y = s.derived_rows(ML, y, col_w, [
        ("Hari Kupon Berjalan / Days of Accrued Interest", num(r["sell_accrued_days"])),
        ("Kupon Berjalan / Accrued Interest (Gross)", _paren(r["sell_accrued"])),
        ("Capital Gain / Loss (Gross)", _paren(r["capital_gain"])),
    ])
    y -= 0.14
    y = s.total_bar(ML, y, col_w, "Jumlah Indikatif yang Diterima (Nett)", _paren(r["net_proceeds"]))
    y -= 0.24

    s.text(ML, y, "Keterangan :", size=8.5)
    y -= 0.26
    ax.add_patch(plt.Rectangle((ML, y - 0.02), 0.16, 0.16, facecolor=_EX_MARKER,
                 edgecolor=_EX_MARKER_EDGE, lw=0.6))
    s.text(ML + 0.24, y + 0.06, "Kolom yang di-highlight kuning WAJIB diisi dengan data terkini.", size=7.5)
    y -= 0.28
    ax.add_patch(plt.Rectangle((ML, y - 0.02), 0.16, 0.16, facecolor="#8a9a5b",
                 edgecolor="#6f7d47", lw=0.6))
    s.text(ML + 0.24, y + 0.06, "Apabila nasabah membeli di Pasar Perdana, maka akan terisi N/A.", size=7.5)

    # ===== RIGHT =====
    y = top
    y = s.bar(rcol_x, y, col_w, "Proyeksi Pendapatan yang Diterima Nasabah (nett)", h=0.40) - 0.06
    y = s.rrow(rcol_x, col_w, y, "Nominal yang dibayar pada tanggal :",
               _paren(-r["amount_paid"]), mid=fmt_date_en(buy_settle), bold_val=True, hair=False)
    y = s.rrow(rcol_x, col_w, y, "Kupon (tidak termasuk pajak capital gain/loss) :", "")
    y = s.schedule(rcol_x, col_w, y, r["schedule"])
    y = s.rrow(rcol_x, col_w, y, "Total Kupon yang telah diterima", _paren(r["coupon_to_sale"]),
               bold_lab=True, bold_val=True)

    y = s.rrow(rcol_x, col_w, y, "Pengembalian yang diterima, jika :", "", bold_lab=True)
    y = s.rrow(rcol_x, col_w, y, "Dijual sebelum JT", _paren(r["proceeds_if_sold"]), dot=True)
    y = s.rrow(rcol_x, col_w, y, "Ditahan hingga JT", _paren(r["proceeds_if_held"]), dot=True)

    y = s.rrow(rcol_x, col_w, y, "Lama waktu investasi (bulan), jika :", "", bold_lab=True)
    y = s.rrow(rcol_x, col_w, y, "Dijual sebelum JT", num(r["months_if_sold"]), dot=True)
    y = s.rrow(rcol_x, col_w, y, "Ditahan hingga JT", num(r["months_if_held"]), dot=True)

    y = s.rrow(rcol_x, col_w, y, "Total keuntungan/kerugian, jika :", "", bold_lab=True)
    y = s.rrow(rcol_x, col_w, y, "Dijual sebelum JT", "", dot=True, bold_lab=True)
    y = s.rrow(rcol_x, col_w, y, "Nominal kumulatif", _paren(r["gain_if_sold"]), indent=0.22)
    y = s.rrow(rcol_x, col_w, y, "Persentase kumulatif", pct(r["pct_if_sold"]), indent=0.22)
    y = s.rrow(rcol_x, col_w, y, "Persentase disetahunkan (gross)", pct(r["ytm_sell"], 2), indent=0.22)
    y = s.rrow(rcol_x, col_w, y, "Ditahan hingga JT", "", dot=True, bold_lab=True)
    y = s.rrow(rcol_x, col_w, y, "Nominal kumulatif", _paren(r["gain_if_held"]), indent=0.22)
    y = s.rrow(rcol_x, col_w, y, "Persentase kumulatif", pct(r["pct_if_held"]), indent=0.22)
    y = s.rrow(rcol_x, col_w, y, "Persentase disetahunkan (gross)", pct(r["ytm_hold"], 2), indent=0.22)

    _disclaimer_box(ax, W, ML, MR, disc_h)
    buf = io.BytesIO()
    FigureCanvasAgg(fig).print_png(buf)
    plt.close(fig)
    return buf.getvalue()


# ---------------------------------------------------------------------------
# SIMULASI SWITCHING — three stacked panels on the left (P1 buy, P1 sell,
# P2 buy) plus two black totals; right side shows P1 coupons, P2 coupons, and
# the hold-vs-switch comparison. Matches the uploaded reference.
# ---------------------------------------------------------------------------
def render_switching_export(r, *, code1, bs1, ss1, nom1, bp1, sp1,
                            code2, s2, m2, nom2, p2, sim_title="SIMULASI SWITCHING"):
    p1 = r["p1"]
    cur1 = p1["meta"].currency
    cur2 = r["meta2"].currency
    meta1, meta2 = p1["meta"], r["meta2"]
    disc_h = 0.26 + len(DISCLAIMER_SWITCH) * 0.208 + 0.14
    # Left column is the tallest element; its content depth below `top` is a
    # fixed ~15.32in. Size the canvas so that depth plus the disclaimer plus the
    # bottom margin fit exactly, then pin the disclaimer under the keterangan.
    W = 12.6
    _left_depth = 15.32
    H = 0.94 + _left_depth + disc_h + 0.28
    fig, ax, top = _new_figure(sim_title, W, H)

    ML, MR, gutter = 0.35, 0.35, 0.45
    col_w = (W - ML - MR - gutter) / 2
    rcol_x = ML + col_w + gutter
    s = _Sheet(ax, ML, col_w)

    # ===== LEFT =====
    y = top
    y = s.banner(ML, y, col_w, code1) - 0.20
    y = s.bar(ML, y, col_w, "Data Produk Obligasi (NASABAH BELI PRODUK 1)", size=10)
    y = s.data_rows(ML, y, col_w, [
        ("Mata Uang", cur1, False),
        ("Tanggal Transaksi", fmt_date_en(bs1), True),
        ("Tanggal Setelmen", fmt_date_en(bs1), True),
        ("Tanggal Kupon Terakhir", "N/A" if p1["is_perdana"] else fmt_date_en(p1["buy_last"]), False),
        ("Tanggal Kupon Pertama", fmt_date_en(p1["buy_next"]), False),
        ("Tanggal Jatuh Tempo (JT)", fmt_date_en(meta1.maturity), False),
        ("Kupon", pct(meta1.coupon, 3), False),
        ("Nilai Nominal", f"{nom1:,.0f}", True),
        ("Harga Nasabah Beli", f"{bp1:.4f}%", True),
    ])
    y -= 0.14
    y = s.derived_rows(ML, y, col_w, [
        ("Hari Kupon Berjalan / Days of Accrued Interest", num(p1["buy_accrued_days"])),
        ("Kupon Berjalan / Accrued Interest", _paren(p1["buy_accrued"])),
    ])
    y -= 0.16
    y = s.bar(ML, y, col_w, "Data Produk Obligasi (NASABAH JUAL PRODUK 1)", size=10)
    y = s.data_rows(ML, y, col_w, [
        ("Tanggal Transaksi", fmt_date_en(ss1), True),
        ("Tanggal Setelmen", fmt_date_en(ss1), True),
        ("Harga Nasabah Jual", f"{sp1:.4f}%", True),
    ])
    y -= 0.14
    y = s.derived_rows(ML, y, col_w, [
        ("Hari Kupon Berjalan / Days of Accrued Interest", num(p1["sell_accrued_days"])),
        ("Kupon Berjalan / Accrued Interest", _paren(p1["sell_accrued"])),
        ("Capital Gain / Loss", _paren(p1["capital_gain"])),
    ])
    y -= 0.14
    y = s.total_bar(ML, y, col_w, "Jumlah Indikatif yang Diterima", _paren(p1["net_proceeds"]))
    y -= 0.26

    s.text(ML, y, "Nasabah melakukan switching dengan produk :", size=8.5)
    y -= 0.30
    y = s.banner(ML, y, col_w, code2) - 0.18
    y = s.bar(ML, y, col_w, "Data Produk Obligasi  (NASABAH BELI PRODUK 2)", size=10)
    y = s.data_rows(ML, y, col_w, [
        ("Mata Uang", cur2, False),
        ("Tanggal Setelmen", fmt_date_en(s2), True),
        ("Tanggal Kupon Terakhir", "N/A" if r["is_perdana2"] else fmt_date_en(r["last_coupon2"]), False),
        ("Tanggal Kupon Pertama", fmt_date_en(r["next_coupon2"]), False),
        ("Tanggal Jatuh Tempo (JT)", fmt_date_en(m2), True),
        ("Kupon", pct(meta2.coupon, 3), False),
        ("Nilai Nominal", f"{nom2:,.0f}", True),
        ("Harga Nasabah Beli", f"{p2:.4f}%", True),
    ])
    y -= 0.14
    y = s.derived_rows(ML, y, col_w, [
        ("Hari Kupon Berjalan / Days of Accrued Interest", num(r["accrued_days2"])),
        ("Kupon Berjalan / Accrued Interest", _paren(r["accrued2"])),
    ])
    y -= 0.16
    y = s.total_bar(ML, y, col_w, "Jumlah Indikatif yang Dibayar", _paren(r["amount_paid2"]))
    y -= 0.06
    topup_label = "Perlu Top up sebesar" if r["top_up"] < 0 else "Kelebihan dikreditkan"
    y = s.total_bar(ML, y, col_w, topup_label, _paren(abs(r["top_up"])))
    y -= 0.26

    # keterangan notes (match reference)
    s.text(ML, y, "Keterangan :", size=8.5)
    y -= 0.22
    s.text(ML + 0.02, y + 0.06, "1)", size=7.5, weight="bold")
    s.text(ML + 0.26, y + 0.06, "Produk 1 & 2 hanya dapat diisi dengan Produk Mata Uang yang sama.", size=7)
    y -= 0.205
    s.text(ML + 0.02, y + 0.06, "2)", size=7.5, weight="bold")
    s.text(ML + 0.26, y + 0.06, "Produk 2 dapat diisi dengan Produk Obligasi Pasar Perdana maupun Sekunder.", size=7)
    y -= 0.24
    ax.add_patch(plt.Rectangle((ML, y - 0.02), 0.15, 0.15, facecolor=_EX_MARKER,
                 edgecolor=_EX_MARKER_EDGE, lw=0.6))
    s.text(ML + 0.24, y + 0.055, "Kolom yang di-highlight kuning WAJIB diisi dengan data terkini.", size=7)
    y -= 0.22
    ax.add_patch(plt.Rectangle((ML, y - 0.02), 0.15, 0.15, facecolor="#8a9a5b",
                 edgecolor="#6f7d47", lw=0.6))
    s.text(ML + 0.24, y + 0.055, "Apabila nasabah membeli di Pasar Perdana, maka harus diisi N/A.", size=7)
    y -= 0.22
    ax.add_patch(plt.Rectangle((ML, y - 0.02), 0.15, 0.15, facecolor="#a9c7e8",
                 edgecolor="#7fa8d0", lw=0.6))
    s.text(ML + 0.24, y + 0.055, "Tanggal JT Produk 2 dapat diisi sesuai JT Produk 2 atau sama dengan JT Produk 1.", size=7)
    left_bottom = y - 0.22

    # ===== RIGHT =====
    y = top
    y = s.bar(rcol_x, y, col_w, "Proyeksi Pendapatan yang Diterima Nasabah (nett)", h=0.40) - 0.06
    y = s.rrow(rcol_x, col_w, y, "Nominal yang dibayar pada tanggal :",
               _paren(-p1["amount_paid"]), mid=fmt_date_en(bs1), bold_val=True, hair=False)
    y = s.rrow(rcol_x, col_w, y, "Kupon Produk 1 (tidak termasuk pajak capital gain/loss) :", "",
               lab_size=8.2)
    y = s.schedule(rcol_x, col_w, y, p1["schedule"])
    y = s.rrow(rcol_x, col_w, y, "Kupon yang diterima hingga jual Produk 1",
               _paren(p1["coupon_to_sale"]), bold_lab=True, bold_val=True, lab_size=8.2)

    y = s.rrow(rcol_x, col_w, y, "Kupon Produk 2 (tidak termasuk pajak capital gain/loss) :", "",
               lab_size=8.2)
    y = s.schedule(rcol_x, col_w, y, r["schedule2"])
    y = s.rrow(rcol_x, col_w, y, "Kupon yang diterima hingga JT Produk 2",
               _paren(r["coupon_to_horizon2"]), bold_lab=True, bold_val=True, lab_size=8.2)

    y = s.rrow(rcol_x, col_w, y, "Pengembalian yang diterima, jika :", "", bold_lab=True)
    y = s.rrow(rcol_x, col_w, y, "Produk 1 ditahan hingga JT", _paren(p1["proceeds_if_held"]), dot=True)
    y = s.rrow(rcol_x, col_w, y, "Produk 1 dialihkan dengan Produk 2",
               _paren(r["total_return_switch"]), dot=True)

    y = s.rrow(rcol_x, col_w, y, "Lama waktu investasi (bulan), jika :", "", bold_lab=True)
    y = s.rrow(rcol_x, col_w, y, "Produk 1 ditahan hingga JT", num(p1["months_if_held"]), dot=True)
    y = s.rrow(rcol_x, col_w, y, "Produk 1 dialihkan dengan Produk 2", num(r["months_switched"]), dot=True)

    y = s.rrow(rcol_x, col_w, y, "Total keuntungan :", "", bold_lab=True)
    y = s.rrow(rcol_x, col_w, y, "Produk 1 ditahan hingga JT", "", dot=True, bold_lab=True)
    y = s.rrow(rcol_x, col_w, y, "Nominal kumulatif", _paren(p1["gain_if_held"]), indent=0.22)
    y = s.rrow(rcol_x, col_w, y, "Persentase kumulatif", pct(p1["pct_if_held"]), indent=0.22)
    y = s.rrow(rcol_x, col_w, y, "Produk 1 dialihkan dengan Produk 2", "", dot=True, bold_lab=True)
    y = s.rrow(rcol_x, col_w, y, "Nominal kumulatif", _paren(r["gain_switch"]), indent=0.22)
    y = s.rrow(rcol_x, col_w, y, "Persentase kumulatif", pct(r["pct_switch"]), indent=0.22)

    _disclaimer_box(ax, W, ML, MR, disc_h, lines=DISCLAIMER_SWITCH, top=left_bottom)
    buf = io.BytesIO()
    FigureCanvasAgg(fig).print_png(buf)
    plt.close(fig)
    return buf.getvalue()


# ---------------------------------------------------------------------------
# SIMULASI KREDIT — single-column credit sheet. Improved over the early draft:
# correct currency symbol, a prominent net-carry verdict, clean sections, and
# a disclaimer that never overlaps the content.
# ---------------------------------------------------------------------------
DISCLAIMER_KREDIT = [
    "Kalkulator ini disediakan hanya sebagai alat bantu simulasi dan tidak dimaksudkan untuk menyediakan rekomendasi atau saran apa pun.",
    "Angka bersifat indikatif; suku bunga, biaya, dan ketentuan final ditetapkan saat pengajuan fasilitas kredit yang sebenarnya.",
    "Bunga investasi memakai pajak kupon 10% untuk obligasi IDR dan 0% untuk INDON/INDOIS. Bunga pinjaman dihitung atas plafon penuh (LTV × nominal).",
    "Cicilan anuitas mengasumsikan suku bunga tetap selama tenor. Perhitungan belum termasuk biaya lain di luar provisi dan admin.",
]

_EX_GAIN = "#0e7c6b"
_EX_LOSS = "#b4342a"
_EX_GAIN_BG = "#e3f2ee"
_EX_LOSS_BG = "#fbe9e7"


def _cur_sym(currency):
    return "Rp" if currency == "IDR" else "$"


def _kmoney(v, currency):
    """Currency-aware money for the export (Rp 1.000.000 / $ 1,000.00)."""
    if v is None or not isinstance(v, (int, float)) or not math.isfinite(v):
        return "—"
    sym = _cur_sym(currency)
    if currency == "IDR":
        body = f"{abs(v):,.0f}".replace(",", ".")
    else:
        body = f"{abs(v):,.2f}"
    sign = "-" if v < 0 else ""
    return f"{sign}{sym} {body}"


def render_kredit_export(r, *, code, prov_pct, sim_title="SIMULASI KREDIT"):
    cur = r["currency"]
    meta = r["meta"]
    disc_h = 0.26 + len(DISCLAIMER_KREDIT) * 0.208 + 0.14

    W = 8.9
    # depth is fixed (known number of rows); size the canvas to fit content + box
    _content = 7.75
    H = 0.94 + _content + disc_h + 0.30
    fig, ax, top = _new_figure(sim_title, W, H)
    # date just under the header rule, right-aligned, clear of the band text
    ax.text(W - 0.35, top - 0.02, f"Tanggal: {fmt_date_en(to_serial(dt.date.today()))}",
            fontsize=8.5, color=MUTED, family=MONO, ha="right", va="top")
    top -= 0.30

    ML, MR = 0.35, 0.35
    x, w = ML, W - ML - MR
    s = _Sheet(ax, x, w)
    y = top

    def krow(cy, lab, val, *, indent=False, strong=False, tone=None, rh=0.30):
        ry = cy - rh
        ax.add_patch(plt.Rectangle((x, ry), w, rh, facecolor=_EX_PANEL if strong else "white",
                     edgecolor=_EX_LINE, lw=0.6))
        col = {"gain": _EX_GAIN, "loss": _EX_LOSS}.get(tone, INK)
        lab_txt = ("   \u21b3 " + lab) if indent else lab
        s.text(x + 0.12, ry + rh / 2, lab_txt, size=9,
               weight="bold" if strong else "normal",
               color=MUTED if indent else INK)
        s.text(x + w - 0.12, ry + rh / 2, val, size=9.5,
               weight="bold" if strong else "normal", ha="right", family=MONO, color=col)
        return ry

    # ===== Section 1: parameters =====
    y = s.bar(x, y, w, "1.  PARAMETER KREDIT & INVESTASI", h=0.36, size=11) - 0.12
    y = krow(y, "Produk Obligasi", code)
    y = krow(y, f"Nominal Investasi ({cur})", _kmoney(r["nominal"], cur))
    y = krow(y, "Loan to Value (LTV)", f"{r['ltv'] * 100:.1f}%")
    y = krow(y, "Plafon Kredit (Nominal × LTV)", _kmoney(r["plafon"], cur), strong=True)
    y = krow(y, "Suku Bunga Kredit (% p.a)", f"{r['loan_rate'] * 100:.2f}%")
    y = krow(y, "Pajak Kupon", "0% (INDON/INDOIS)" if r["tax"] == 0 else "10% (obligasi IDR)")
    y = krow(y, "Tenor Pinjaman", f"{r['tenor_years']:.1f} tahun ({r['tenor_months']} bulan)")
    y -= 0.22

    # ===== Section 2: income vs expense =====
    y = s.bar(x, y, w, "2.  ESTIMASI PENDAPATAN & BEBAN", h=0.36, size=11) - 0.12
    y = krow(y, "Bunga Investasi / Tahun (Nett)", _kmoney(r["coupon_net_yr"], cur),
             strong=True, tone="gain")
    y = krow(y, "per Bulan", _kmoney(r["inv_month"], cur), indent=True, tone="gain")
    y = krow(y, "per Hari", _kmoney(r["inv_day"], cur), indent=True, tone="gain")
    y = krow(y, "Bunga Pinjaman / Tahun (Maksimal)", _kmoney(r["loan_interest_yr"], cur),
             strong=True, tone="loss")
    y = krow(y, "per Bulan", _kmoney(r["loan_month"], cur), indent=True, tone="loss")
    y = krow(y, "per Hari", _kmoney(r["loan_day"], cur), indent=True, tone="loss")
    y = krow(y, f"Cicilan / Bulan (Anuitas, {r['tenor_months']} bln)", _kmoney(r["annuity"], cur),
             strong=True)
    y = krow(y, "Total Pembayaran s.d. Lunas", _kmoney(r["total_repay"], cur))
    y = krow(y, "Total Biaya Provisi & Admin", _kmoney(r["total_fees"], cur))
    y -= 0.24

    # ===== Verdict banner =====
    carry = r["net_carry_yr"]
    ok = carry >= 0
    bh = 0.72
    ax.add_patch(plt.Rectangle((x, y - bh), w, bh, facecolor=_EX_GAIN_BG if ok else _EX_LOSS_BG,
                 edgecolor=_EX_GAIN if ok else _EX_LOSS, lw=1.2))
    s.text(x + 0.16, y - 0.24, "SELISIH CARRY / TAHUN", size=8.5, color=MUTED, weight="bold")
    s.text(x + 0.16, y - 0.50, _kmoney(carry, cur), size=16, family=MONO, weight="bold",
           color=_EX_GAIN if ok else _EX_LOSS)
    verdict = ("Kupon menutup bunga pinjaman" if ok else "Kupon belum menutup bunga pinjaman")
    s.text(x + w - 0.16, y - 0.30, verdict, size=8.5, ha="right", weight="bold",
           color=_EX_GAIN if ok else _EX_LOSS)
    s.text(x + w - 0.16, y - 0.50,
           f"{_kmoney(abs(r['net_carry_month']), cur)} / bulan", size=9, ha="right",
           family=MONO, color=_EX_GAIN if ok else _EX_LOSS)
    y -= bh + 0.10

    # tenor warning line inside the sheet
    if not r["tenor_ok"]:
        s.text(x, y - 0.12,
               f"\u26a0  Tenor melebihi batas: fasilitas harus berakhir \u2264 1 bulan sebelum "
               f"JT obligasi ({fmt_date_en(meta.maturity)}). Maks \u2248 {r['max_tenor_months']} bulan.",
               size=7.6, color=_EX_LOSS, weight="bold")

    _disclaimer_box(ax, W, ML, MR, disc_h, lines=DISCLAIMER_KREDIT)
    buf = io.BytesIO()
    FigureCanvasAgg(fig).print_png(buf)
    plt.close(fig)
    return buf.getvalue()


DISCLAIMER = [
    "Kalkulator ini hanya alat bantu simulasi dan bukan rekomendasi atau saran investasi.",
    "Simulasi, harga, dan YTM bersifat indikatif; dapat berbeda dengan perhitungan saat transaksi sebenarnya.",
    "YTM mengasumsikan kupon dibayar sesuai jadwal dan diinvestasikan kembali pada tingkat yang sama.",
    "Tarif pajak yang dipakai adalah 10%. Perhitungan belum memotong biaya lain (apabila ada).",
    "Tidak berlaku untuk FR0088 dan FR0089 yang ditransaksikan pada 6-7 Januari 2021.",
]

CSS = """
<style>
  .stApp { background: #eef2f5; }
  .block-container { padding-top: 2.2rem; max-width: 1140px; }
  h1, h2, h3, h4, p, label, span, div { color: #0f2233; }
  .ko-head { display:flex; align-items:center; gap:14px; margin-bottom:4px; }
  .ko-mark { width:8px; height:42px; border-radius:4px;
             background:linear-gradient(#ffd400 50%, #0e7c6b 50%); }
  .ko-title { font-size:22px; font-weight:700; margin:0; letter-spacing:-.01em; }
  .ko-sub { font-size:13px; color:#5c7085; margin:2px 0 0; }
  .ko-card { background:#fff; border:1px solid #dbe3ea; border-radius:10px;
             padding:14px 16px; margin-bottom:12px; }
  .ko-cap { font-size:11px; font-weight:700; letter-spacing:.08em;
            text-transform:uppercase; color:#5c7085; margin-bottom:8px; }
  .ko-row { display:flex; justify-content:space-between; gap:14px;
            font-size:13.5px; padding:3px 0; }
  .ko-row .v { font-family:ui-monospace,"DejaVu Sans Mono",Menlo,monospace;
               font-variant-numeric:tabular-nums; }
  .ko-ind { padding-left:14px; color:#5c7085; font-size:12.5px; }
  .ko-strong { border-top:1px solid #dbe3ea; padding-top:8px; margin-top:4px; font-weight:700; }
  .gain { color:#0e7c6b; } .loss { color:#b4342a; }
  .ko-big { background:#f6f9fb; border:1px solid #dbe3ea; border-radius:9px;
            padding:12px 14px; margin-bottom:10px; }
  .ko-big .l { font-size:11px; font-weight:700; letter-spacing:.06em;
               text-transform:uppercase; color:#5c7085; }
  .ko-big .v { font-family:ui-monospace,"DejaVu Sans Mono",Menlo,monospace;
               font-size:23px; font-weight:700; display:block; line-height:1.3; }
  .ko-big .s { font-size:12px; color:#5c7085; }
  .ko-req label p { font-weight:600 !important; }
  div[data-testid="stForm"] { border-color:#dbe3ea; }

  /* Streamlit renders each st.markdown as its own block; an opening <div> with
     no text becomes an empty bordered box. Collapse those stray empties. */
  div[data-testid="stMarkdownContainer"]:empty { display:none; }
  .ko-card:empty { display:none; }
  .stDivider { margin:0.4rem 0; }

  /* ---- Fillable inputs: uniform white backdrop + dark text ----
     Every control the user can fill (dropdowns, text, number with steppers,
     date pickers) gets a white fill and a light border, so slots stand out
     against the grey page and are easy to navigate on desktop and mobile. */

  /* dropdowns (closed control) */
  div[data-baseweb="select"] > div {
    background:#ffffff !important;
    border:1px solid #c3ced8 !important;
  }
  div[data-baseweb="select"] > div * { color:#0f2233 !important; }
  div[data-baseweb="select"] svg { fill:#5c7085 !important; color:#5c7085 !important; }
  div[data-baseweb="select"] input { color:#0f2233 !important; caret-color:#0f2233; }
  /* open option list: dark text on white */
  ul[data-baseweb="menu"] { background:#ffffff !important; }
  ul[data-baseweb="menu"] li { color:#0f2233 !important; }

  /* text inputs (nominal, admin) */
  .stTextInput div[data-baseweb="input"],
  .stTextInput div[data-baseweb="base-input"] {
    background:#ffffff !important;
    border:1px solid #c3ced8 !important;
    border-radius:8px !important;
  }
  .stTextInput input {
    background:#ffffff !important;
    color:#0f2233 !important;
    -webkit-text-fill-color:#0f2233 !important;
  }

  /* number inputs (price, LTV, rate, tenor, provisi) incl. stepper buttons */
  .stNumberInput div[data-baseweb="input"],
  .stNumberInput div[data-baseweb="base-input"] {
    background:#ffffff !important;
    border:1px solid #c3ced8 !important;
    border-radius:8px !important;
  }
  .stNumberInput input {
    background:#ffffff !important;
    color:#0f2233 !important;
    -webkit-text-fill-color:#0f2233 !important;
  }
  .stNumberInput button {
    background:#ffffff !important;
    color:#0f2233 !important;
    border-left:1px solid #e4eaef !important;
  }
  .stNumberInput button:hover { background:#eef2f5 !important; }
  .stNumberInput button svg { fill:#5c7085 !important; }

  /* date pickers */
  .stDateInput div[data-baseweb="input"],
  .stDateInput div[data-baseweb="base-input"] {
    background:#ffffff !important;
    border:1px solid #c3ced8 !important;
    border-radius:8px !important;
  }
  .stDateInput input {
    background:#ffffff !important;
    color:#0f2233 !important;
    -webkit-text-fill-color:#0f2233 !important;
  }

  /* focus ring for whichever input is active */
  div[data-baseweb="select"] > div:focus-within,
  .stTextInput div[data-baseweb="input"]:focus-within,
  .stNumberInput div[data-baseweb="input"]:focus-within,
  .stDateInput div[data-baseweb="input"]:focus-within {
    border-color:#0e7c6b !important;
    box-shadow:0 0 0 2px rgba(14,124,107,.15) !important;
  }
</style>
"""


def row(label, value, tone=None, indent=False, strong=False):
    cls = "ko-row" + (" ko-ind" if indent else "") + (" ko-strong" if strong else "")
    vcls = "v" + (f" {tone}" if tone in ("gain", "loss") else "")
    st.markdown(f'<div class="{cls}"><span>{label}</span>'
                f'<span class="{vcls}">{value}</span></div>', unsafe_allow_html=True)


def _row_html(label, value, tone=None, indent=False, strong=False):
    cls = "ko-row" + (" ko-ind" if indent else "") + (" ko-strong" if strong else "")
    vcls = "v" + (f" {tone}" if tone in ("gain", "loss") else "")
    return f'<div class="{cls}"><span>{label}</span><span class="{vcls}">{value}</span></div>'


def card(heading, rows):
    """
    Render a whole card — heading + every row — in ONE markdown call, so the
    styled border actually wraps its content instead of leaving an empty box
    (a limitation of splitting the opening/closing <div> across st.markdown).
    rows: list of (label, value, tone, indent, strong) tuples; tone/indent/
    strong are optional per tuple.
    """
    parts = []
    for r in rows:
        label = r[0]
        value = r[1]
        tone = r[2] if len(r) > 2 else None
        indent = r[3] if len(r) > 3 else False
        strong = r[4] if len(r) > 4 else False
        parts.append(_row_html(label, value, tone, indent, strong))
    head = f'<div class="ko-cap" style="margin-bottom:6px">{heading}</div>' if heading else ""
    st.markdown(f'<div class="ko-card">{head}{"".join(parts)}</div>', unsafe_allow_html=True)


def big(label, value, sub="", tone=None):
    vcls = "v" + (f" {tone}" if tone in ("gain", "loss") else "")
    st.markdown(f'<div class="ko-big"><span class="l">{label}</span>'
                f'<span class="{vcls}">{value}</span>'
                f'<span class="s">{sub}</span></div>', unsafe_allow_html=True)


def tone_of(value):
    return "gain" if value >= 0 else "loss"


def show_meta(meta: Meta):
    st.caption(
        f"{meta.currency} · kupon {pct(meta.coupon, 4)} · "
        f"{'bulanan' if meta.frequency == 'Monthly' else 'semesteran'} · "
        f"jatuh tempo {fmt_date(meta.maturity)}"
    )


def show_schedule(rows, currency, caption):
    st.markdown(f'<div class="ko-cap">{caption}'
                f'{"" if not rows else f" — {len(rows)} pembayaran"}</div>',
                unsafe_allow_html=True)
    if not rows:
        st.caption("Tidak ada pembayaran kupon dalam periode ini.")
        return
    st.dataframe(
        {"#": [r[0] for r in rows],
         "Tanggal": [fmt_date_short(r[1]) for r in rows],
         "Kupon nett": [money(r[2], currency) for r in rows]},
        hide_index=True, height=min(38 + len(rows) * 35, 460), **FULL_WIDTH,
    )


def png_to_jpg(png_bytes, quality=92):
    """Flatten PNG onto white and re-encode as JPEG (no alpha; iOS-friendly)."""
    from PIL import Image
    im = Image.open(io.BytesIO(png_bytes))
    if im.mode in ("RGBA", "LA", "P"):
        bg = Image.new("RGB", im.size, "white")
        im = im.convert("RGBA")
        bg.paste(im, mask=im.split()[-1])
        im = bg
    else:
        im = im.convert("RGB")
    out = io.BytesIO()
    im.save(out, format="JPEG", quality=quality, subsampling=0, optimize=True)
    return out.getvalue()


def export_filename(sim_label, seri):
    """SIMULASI BELI_FR0110_14-Aug-2026 — the naming convention the desk uses."""
    stamp = f"{dt.date.today():%d-%b-%Y}"
    safe = "".join(c for c in f"{sim_label}_{seri}_{stamp}" if c not in '\\/:*?"<>|')
    return f"{safe}.jpg"


def export_section(png_factory, sim_label, seri, state_key):
    """
    Build-on-demand export. The heavy image is only rendered once the RM taps
    the button, so the normal working view stays a single clean results panel
    rather than the same simulation printed twice. After the first tap the
    preview persists (session state) so re-runs from other inputs don't hide it.

    png_factory: zero-arg callable returning PNG bytes.
    """
    import base64

    made = st.session_state.get(state_key, False)
    label = "Perbarui gambar" if made else "Buat gambar untuk dibagikan"
    if st.button(label, key=f"{state_key}_btn", type="primary", **FULL_WIDTH):
        made = True
        st.session_state[state_key] = True

    if not made:
        st.caption("Tekan tombol di atas untuk membuat gambar simulasi yang bisa "
                   "diunduh atau dibagikan ke nasabah.")
        return

    try:
        jpg = png_to_jpg(png_factory())
    except Exception as exc:  # noqa: BLE001 - surfaced to the user
        st.warning(f"Gambar tidak dapat dibuat: {exc}")
        return

    b64 = base64.b64encode(jpg).decode("ascii")
    st.markdown(
        f'<img src="data:image/jpeg;base64,{b64}" '
        'style="width:100%;border:1px solid #dbe3ea;border-radius:8px;margin:6px 0;" '
        'alt="Pratinjau simulasi" />',
        unsafe_allow_html=True,
    )
    st.caption(
        "📱 iPhone/iPad: tekan lama (long-press) gambar di atas, lalu pilih "
        "**Save to Photos** untuk menyimpan sebagai JPG. "
        "💻 Komputer: gunakan tombol unduh di bawah."
    )
    st.download_button(
        "Unduh gambar (JPG)",
        data=jpg,
        file_name=export_filename(sim_label, seri),
        mime="image/jpeg",
        **FULL_WIDTH,
    )


# ===========================================================================
# Tabs
# ===========================================================================
def tab_beli():
    left, right = st.columns([5, 6], gap="large")

    with left:
        st.markdown('<div class="ko-cap">Data produk</div>', unsafe_allow_html=True)
        code = st.selectbox("Kode obligasi", PRODUCT_CODES,
                            index=PRODUCT_CODES.index("FR0110"), key="b_code")
        meta = product_meta(code)
        show_meta(meta)
        market = st.selectbox("Jenis transaksi", ["Pasar Sekunder", "Pasar Perdana"], key="b_market")
        txn = st.date_input("Tanggal transaksi", dt.date.today() - dt.timedelta(days=2),
                            min_value=dt.date(2000, 1, 1), max_value=dt.date(2100, 1, 1), key="b_txn",
                            help="Tanggal order nasabah. Hanya ditampilkan pada gambar.")
        settle = st.date_input("Tanggal setelmen", dt.date.today(),
                               min_value=dt.date(2000, 1, 1), max_value=dt.date(2100, 1, 1), key="b_settle",
                               help="Tanggal setelmen inilah yang dipakai untuk menghitung "
                                    "kupon berjalan, YTM, dan seluruh proyeksi.")
        nominal = parse_number(money_input("Nilai nominal", "200.000.000", "b_nominal", cur_hint="Rp"))
        price_in = st.number_input("Harga nasabah beli (%)", value=100.61, step=0.05,
                                   format="%.4f", key="b_price")

    try:
        r = simulate_beli(code, market, to_serial(settle), nominal,
                          None if price_in is None else price_in / 100)
    except SimError as exc:
        with right:
            st.info(str(exc))
        return

    cur = r["meta"].currency
    with left:
        card("Rincian pembelian", [
            ("Tanggal kupon terakhir", "N/A" if r["is_perdana"] else fmt_date(r["last_coupon"])),
            ("Tanggal kupon berikutnya", fmt_date(r["next_coupon"])),
            ("Hari kupon berjalan", f"{num(r['accrued_days'])} hari"),
            ("Kupon berjalan", money(r["accrued_interest"], cur)),
            ("Imbal hasil hingga JT (gross)", pct(r["ytm"], 4)),
            ("Jumlah indikatif dibayar", money(r["amount_paid"], cur), None, False, True),
        ])

    with right:
        st.markdown('<div class="ko-cap">Proyeksi pendapatan (nett)</div>', unsafe_allow_html=True)
        a, b = st.columns(2)
        with a:
            big("Total keuntungan", money(r["cumulative_nominal"], cur),
                f"{pct(r['cumulative_percent'])} kumulatif atas modal",
                tone_of(r["cumulative_nominal"]))
        with b:
            big("Pengembalian hingga JT", money(r["proceeds_at_maturity"], cur),
                f"{num(r['periods'])} periode · {num(r['months'])} bulan")
        card(None, [
            ("Kupon diterima hingga JT", money(r["coupon_to_maturity"], cur)),
            ("Nominal diterima saat JT", money(r["received_at_maturity"], cur)),
            ("Pengembalian pokok", money(r["principal_back"], cur), None, True),
            ("Kupon saat JT (gross)", money(r["final_coupon_gross"], cur), None, True),
            ("Capital gain / loss", money(r["capital_gain"], cur), tone_of(r["capital_gain"]), True),
            ("Total pajak", money(r["total_tax"], cur), None, True),
            ("Pajak kupon", money(r["last_coupon_tax"], cur), None, True),
            ("Pajak capital gain", money(r["capital_gain_tax"], cur), None, True),
        ])

    # ---- schedule (full width, not collapsed) ----
    show_schedule(r["schedule"], cur, "Jadwal kupon")

    # ---- export section, at the bottom ----
    st.divider()
    st.markdown('<div class="ko-cap">Gambar untuk dibagikan</div>', unsafe_allow_html=True)
    export_section(
        lambda: render_beli_export(
            r, code=code, market=market,
            txn_serial=to_serial(txn), settle_serial=to_serial(settle),
            nominal=nominal, price_pct=price_in, sim_title="SIMULASI BELI"),
        "SIMULASI BELI", code, "b_export")


def tab_jual():
    left, right = st.columns([5, 6], gap="large")

    with left:
        st.markdown('<div class="ko-cap">Saat nasabah beli</div>', unsafe_allow_html=True)
        code = st.selectbox("Kode obligasi", PRODUCT_CODES,
                            index=PRODUCT_CODES.index("FR0110"), key="j_code")
        meta = product_meta(code)
        show_meta(meta)
        buy_settle = st.date_input("Tanggal setelmen beli", dt.date.today() - dt.timedelta(days=365),
                                   min_value=dt.date(2000, 1, 1), max_value=dt.date(2100, 1, 1), key="j_bs")
        nominal = parse_number(money_input("Nilai nominal", "200.000.000", "j_nominal"))
        buy_price = st.number_input("Harga nasabah beli (%)", value=100.61, step=0.05,
                                    format="%.4f", key="j_bp")
        st.markdown('<div class="ko-cap">Saat nasabah jual</div>', unsafe_allow_html=True)
        sell_settle = st.date_input("Tanggal setelmen jual", dt.date.today(),
                                    min_value=dt.date(2000, 1, 1), max_value=dt.date(2100, 1, 1), key="j_ss")
        sell_price = st.number_input("Harga nasabah jual (%)", value=99.00, step=0.05,
                                     format="%.4f", key="j_sp")

    try:
        r = simulate_jual(code, to_serial(buy_settle), nominal, buy_price / 100,
                          to_serial(sell_settle), sell_price / 100)
    except SimError as exc:
        with right:
            st.info(str(exc))
        return

    cur = r["meta"].currency
    with left:
        card("Rincian penjualan", [
            ("Hari kupon berjalan", f"{num(r['sell_accrued_days'])} hari"),
            ("Kupon berjalan (gross)", money(r["sell_accrued"], cur)),
            ("Capital gain / loss (gross)", money(r["capital_gain"], cur), tone_of(r["capital_gain"])),
            ("Jumlah indikatif diterima (nett)", money(r["net_proceeds"], cur), None, False, True),
        ])

    with right:
        st.markdown('<div class="ko-cap">Jika dijual sebelum jatuh tempo</div>', unsafe_allow_html=True)
        a, b = st.columns(2)
        with a:
            big("Total keuntungan", money(r["gain_if_sold"], cur),
                f"{pct(r['pct_if_sold'])} · {num(r['months_if_sold'])} bulan",
                tone_of(r["gain_if_sold"]))
        with b:
            big("Pengembalian diterima", money(r["proceeds_if_sold"], cur),
                f"imbal hasil {pct(r['ytm_sell'], 4)}")

        st.markdown('<div class="ko-cap">Jika ditahan hingga jatuh tempo</div>', unsafe_allow_html=True)
        c, d = st.columns(2)
        with c:
            big("Total keuntungan", money(r["gain_if_held"], cur),
                f"{pct(r['pct_if_held'])} · {num(r['months_if_held'])} bulan",
                tone_of(r["gain_if_held"]))
        with d:
            big("Pengembalian diterima", money(r["proceeds_if_held"], cur),
                f"imbal hasil {pct(r['ytm_hold'], 4)}")

        card(None, [
            ("Jumlah indikatif dibayar saat beli", money(r["amount_paid"], cur)),
            ("Total kupon telah diterima", money(r["coupon_to_sale"], cur)),
            ("Total pajak saat jual", money(r["total_tax"], cur)),
        ])

    show_schedule(r["schedule"], cur, "Jadwal kupon hingga penjualan")

    st.divider()
    st.markdown('<div class="ko-cap">Gambar untuk dibagikan</div>', unsafe_allow_html=True)
    export_section(
        lambda: render_jual_export(
            r, code=code,
            buy_txn=to_serial(buy_settle) - 2, buy_settle=to_serial(buy_settle),
            nominal=nominal, buy_price=buy_price,
            sell_txn=to_serial(sell_settle) - 2, sell_settle=to_serial(sell_settle),
            sell_price=sell_price, sim_title="SIMULASI JUAL"),
        "SIMULASI JUAL", code, "j_export")


def tab_switching():
    left, right = st.columns([5, 6], gap="large")

    with left:
        st.markdown('<div class="ko-cap">Produk 1 — beli</div>', unsafe_allow_html=True)
        code1 = st.selectbox("Kode obligasi Produk 1", PRODUCT_CODES,
                             index=PRODUCT_CODES.index("FR0110"), key="s_c1")
        show_meta(product_meta(code1))
        bs1 = st.date_input("Setelmen beli Produk 1", dt.date.today() - dt.timedelta(days=365),
                            min_value=dt.date(2000, 1, 1), max_value=dt.date(2100, 1, 1), key="s_bs1")
        nom1 = parse_number(money_input("Nilai nominal Produk 1", "200.000.000", "s_n1"))
        bp1 = st.number_input("Harga beli Produk 1 (%)", value=100.61, step=0.05, format="%.4f", key="s_bp1")
        st.markdown('<div class="ko-cap">Produk 1 — jual</div>', unsafe_allow_html=True)
        ss1 = st.date_input("Setelmen jual Produk 1", dt.date.today(),
                            min_value=dt.date(2000, 1, 1), max_value=dt.date(2100, 1, 1), key="s_ss1")
        sp1 = st.number_input("Harga jual Produk 1 (%)", value=99.00, step=0.05, format="%.4f", key="s_sp1")

        st.markdown('<div class="ko-cap">Produk 2 — beli</div>', unsafe_allow_html=True)
        code2 = st.selectbox("Kode obligasi Produk 2", PRODUCT_CODES,
                             index=PRODUCT_CODES.index("FR0100"), key="s_c2")
        meta2 = product_meta(code2)
        show_meta(meta2)
        s2 = st.date_input("Setelmen Produk 2", dt.date.today(),
                           min_value=dt.date(2000, 1, 1), max_value=dt.date(2100, 1, 1), key="s_s2")
        m2 = st.date_input("Jatuh tempo dipakai untuk Produk 2", to_date(meta2.maturity),
                           min_value=dt.date(2000, 1, 1), max_value=dt.date(2100, 1, 1), key="s_m2",
                           help="Boleh disamakan dengan jatuh tempo Produk 1 agar kedua pilihan "
                                "dibandingkan dalam rentang waktu yang sama.")
        nom2 = parse_number(money_input("Nilai nominal Produk 2", "200.000.000", "s_n2"))
        p2 = st.number_input("Harga beli Produk 2 (%)", value=99.50, step=0.05, format="%.4f", key="s_p2")

    try:
        r = simulate_switching(code1, to_serial(bs1), nom1, bp1 / 100, to_serial(ss1), sp1 / 100,
                               code2, to_serial(s2), to_serial(m2), nom2, p2 / 100)
    except SimError as exc:
        with right:
            st.info(str(exc))
        return

    cur1 = r["p1"]["meta"].currency
    cur2 = r["meta2"].currency

    with right:
        if r["currency_mismatch"]:
            st.error(
                f"Produk 1 ({cur1}) dan Produk 2 ({cur2}) berbeda mata uang. Simulasi switching "
                "hanya berlaku untuk dua produk dengan mata uang yang sama — angka di bawah tidak "
                "dapat dipakai. Ganti salah satu produk."
            )
        st.markdown('<div class="ko-cap">Perbandingan dua pilihan</div>', unsafe_allow_html=True)
        a, b = st.columns(2)
        with a:
            st.caption("Produk 1 ditahan hingga JT")
            big("Total keuntungan", money(r["p1"]["gain_if_held"], cur1),
                f"{pct(r['p1']['pct_if_held'])} · {num(r['p1']['months_if_held'])} bulan",
                tone_of(r["p1"]["gain_if_held"]))
            row("Pengembalian diterima", money(r["p1"]["proceeds_if_held"], cur1))
        with b:
            st.caption("Dialihkan ke Produk 2")
            big("Total keuntungan", money(r["gain_switch"], cur1),
                f"{pct(r['pct_switch'])} · {num(r['months_switched'])} bulan",
                tone_of(r["gain_switch"]))
            row("Pengembalian diterima", money(r["total_return_switch"], cur1))

        card("Rincian peralihan", [
            ("Modal awal (harga bersih Produk 1)", money(r["capital"], cur1)),
            ("Hasil jual Produk 1 (nett)", money(r["proceeds_sell1"], cur1)),
            ("Untung / rugi jual Produk 1", money(r["gain_sell1"], cur1), tone_of(r["gain_sell1"]), True),
            ("Kupon s.d. jual Produk 1", money(r["coupon_until_sale1"], cur1), None, True),
            ("Jumlah dibayar untuk Produk 2", money(r["paid2"], cur2)),
            ("Kupon berjalan Produk 2", money(r["accrued2"], cur2), None, True),
            ("Imbal hasil Produk 2 hingga JT", pct(r["ytm2"], 4), None, True),
            ("Kupon Produk 2 s.d. JT", money(r["coupon_to_horizon2"], cur2), None, True),
            ("Perlu top up sebesar" if r["top_up"] < 0 else "Kelebihan dikreditkan sebesar",
             money(abs(r["top_up"]), cur1), None, False, True),
        ])

    show_schedule(r["schedule2"], cur2, "Jadwal kupon Produk 2")

    st.divider()
    st.markdown('<div class="ko-cap">Gambar untuk dibagikan</div>', unsafe_allow_html=True)
    if r["currency_mismatch"]:
        st.caption("Gambar tidak dibuat karena kedua produk berbeda mata uang.")
    else:
        export_section(
            lambda: render_switching_export(
                r, code1=code1, bs1=to_serial(bs1), ss1=to_serial(ss1),
                nom1=nom1, bp1=bp1, sp1=sp1,
                code2=code2, s2=to_serial(s2), m2=to_serial(m2),
                nom2=nom2, p2=p2, sim_title="SIMULASI SWITCHING"),
            "SIMULASI SWITCHING", f"{code1}-{code2}", "s_export")


# ===========================================================================
# Entry point
# ===========================================================================
def tab_kredit():
    left, right = st.columns([5, 6], gap="large")

    with left:
        st.markdown('<div class="ko-cap">Parameter kredit</div>', unsafe_allow_html=True)
        code = st.selectbox("Nama produk (agunan)", PRODUCT_CODES,
                            index=PRODUCT_CODES.index("FR0110"), key="k_code")
        meta = product_meta(code)
        show_meta(meta)

        nominal = parse_number(money_input(
            "Nominal investasi", "1.000.000.000", "k_nominal",
            cur_hint="Rp" if meta.currency == "IDR" else "$"))

        # LTV auto-fills from the bond's currency/maturity tier; RM may override.
        auto_ltv = default_ltv(meta)
        years = max(0.0, (meta.maturity - to_serial(dt.date.today())) / 365.25)
        st.caption(
            f"LTV otomatis **{auto_ltv * 100:.0f}%** "
            f"({meta.currency}, sisa tenor ±{years:.1f} tahun). "
            "Ubah bila perlu."
        )
        # reset the LTV field to the tier default whenever the bond changes
        if st.session_state.get("_k_last_code") != code:
            st.session_state["k_ltv"] = auto_ltv * 100
            st.session_state["_k_last_code"] = code
        ltv_pct = st.number_input("Loan to Value / LTV (%)", min_value=1.0, max_value=100.0,
                                  step=1.0, format="%.1f", key="k_ltv")

        loan_rate = st.number_input("Suku bunga kredit (% p.a)", min_value=0.0, max_value=100.0,
                                    value=7.50, step=0.05, format="%.4f", key="k_rate",
                                    help="Diisi manual oleh RM sesuai penawaran fasilitas.")
        tenor = st.number_input("Periode angsuran (tahun)", min_value=0.5, max_value=30.0,
                                value=3.0, step=0.5, format="%.1f", key="k_tenor")

        with st.expander("Biaya provisi & admin (opsional)"):
            prov_pct = st.number_input("Provisi (% dari plafon)", min_value=0.0, max_value=100.0,
                                       value=1.00, step=0.05, format="%.4f", key="k_prov")
            admin = parse_number(money_input(
                "Biaya admin (flat)", "0", "k_admin",
                cur_hint="Rp" if meta.currency == "IDR" else "$")) or 0

    try:
        r = simulate_kredit(code, nominal, ltv_pct / 100, loan_rate / 100, tenor,
                            prov_pct / 100, admin)
    except SimError as exc:
        with right:
            st.info(str(exc))
        return

    cur = r["currency"]
    with left:
        card("Ringkasan fasilitas", [
            ("Plafon kredit (Nominal × LTV)", money(r["plafon"], cur), None, False, True),
            ("Pajak kupon", "0% (INDON/INDOIS)" if r["tax"] == 0 else "10% (obligasi IDR)"),
        ])

    with right:
        st.markdown('<div class="ko-cap">Perbandingan tahunan</div>', unsafe_allow_html=True)
        a, b = st.columns(2)
        with a:
            big("Bunga investasi / th (nett)", money(r["coupon_net_yr"], cur),
                f"{money(r['inv_month'], cur)}/bln", "gain")
        with b:
            big("Bunga pinjaman / th (maks)", money(r["loan_interest_yr"], cur),
                f"{money(r['loan_month'], cur)}/bln", "loss")

        big("Selisih carry / tahun", money(r["net_carry_yr"], cur),
            (f"Kupon menutup bunga pinjaman — surplus {money(r['net_carry_month'], cur)}/bln"
             if r["net_carry_yr"] >= 0
             else f"Kupon belum menutup bunga — defisit {money(abs(r['net_carry_month']), cur)}/bln"),
            tone_of(r["net_carry_yr"]))

        card("Beban pinjaman", [
            ("Bunga pinjaman / tahun", money(r["loan_interest_yr"], cur)),
            ("Bunga pinjaman / bulan", money(r["loan_month"], cur), None, True),
            ("Bunga pinjaman / hari", money(r["loan_day"], cur), None, True),
            (f"Cicilan / bulan (anuitas, {r['tenor_months']} bln)", money(r["annuity"], cur), None, False, True),
            ("Total pembayaran", money(r["total_repay"], cur), None, True),
            ("Total bunga selama tenor", money(r["total_loan_interest"], cur), None, True),
        ])
        card("Biaya di muka", [
            (f"Provisi ({pct(prov_pct / 100, 2)} × plafon)", money(r["provision"], cur)),
            ("Biaya admin", money(r["admin_fee"], cur)),
            ("Total biaya provisi & admin", money(r["total_fees"], cur), None, False, True),
        ])

    if not r["tenor_ok"]:
        max_yr = r["max_tenor_months"] / 12
        st.error(
            f"⚠️ Periode angsuran {tenor:.1f} tahun melampaui batas. Fasilitas kredit harus "
            f"berakhir paling lambat 1 bulan sebelum jatuh tempo obligasi "
            f"({fmt_date(meta.maturity)}). Tenor maksimal ≈ {r['max_tenor_months']} bulan "
            f"({max_yr:.1f} tahun)."
        )

    st.divider()
    st.markdown('<div class="ko-cap">Gambar untuk dibagikan</div>', unsafe_allow_html=True)
    export_section(
        lambda: render_kredit_export(
            r, code=code, prov_pct=prov_pct / 100, sim_title="SIMULASI KREDIT"),
        "SIMULASI KREDIT", code, "k_export")


def main():
    st.set_page_config(page_title="Kalkulator Obligasi", page_icon="🧮",
                       layout="wide", initial_sidebar_state="collapsed")
    st.markdown(CSS, unsafe_allow_html=True)
    st.markdown(
        '<div class="ko-head"><div class="ko-mark"></div><div>'
        '<p class="ko-title">Kalkulator Obligasi</p>'
        '<p class="ko-sub">Simulasi beli, jual, switching, dan kredit · versi 2.5.0</p>'
        "</div></div>",
        unsafe_allow_html=True,
    )

    beli, jual, switching, kredit = st.tabs(
        ["Simulasi Beli", "Simulasi Jual", "Simulasi Switching", "Simulasi Kredit"])
    with beli:
        tab_beli()
    with jual:
        tab_jual()
    with switching:
        tab_switching()
    with kredit:
        tab_kredit()

    with st.expander("Disclaimer"):
        for text in DISCLAIMER:
            st.markdown(f"- {text}")


if __name__ == "__main__":
    main()
