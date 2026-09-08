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
import textwrap
from dataclasses import dataclass

import streamlit as st
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.backends.backend_agg import FigureCanvasAgg
from matplotlib.patches import FancyBboxPatch

# ---------------------------------------------------------------------------
# Product master table — C333:R497 of "Simulasi Beli".
# Columns: code, issued, maturity, coupon, coupon day, first month, second
# month, first accrual date, first coupon date. Dates are Excel serials.
# ---------------------------------------------------------------------------
PRODUCTS_RAW = json.loads(r"""[["FR0056",40444,46280,0.08375,15,3,9,null,null],["FR0058",40745,48380,0.0825,15,12,6,null,null],["FR0059",40801,46522,0.07,15,11,5,null,null],["FR0064",41134,46888,0.06125,15,11,5,null,null],["FR0065",41151,48714,0.06625,15,11,5,null,null],["FR0068",41487,49018,0.08375,15,9,3,null,null],["FR0071",41529,47192,0.09,15,9,3,null,null],["FR0072",42194,49810,0.0825,15,11,5,null,null],["FR0073",42222,47983,0.0875,15,11,5,null,null],["FR0074",42684,48441,0.075,15,2,8,null,null],["FR0075",42957,50540,0.075,15,11,5,null,null],["FR0076",43000,54193,0.07375,15,11,5,null,null],["FR0078",43235,47253,0.0825,15,11,5,null,null],["FR0079",43388,50875,0.08375,15,4,10,null,null],["FR0080",43631,49475,0.075,15,12,6,null,null],["FR0082",43539,47741,0.07,15,9,3,null,null],["FR0083",43753,51241,0.075,15,4,10,null,null],["FR0084",43876,46068,0.0725,15,8,2,null,null],["FR0085",43936,47953,0.0775,15,10,4,null,null],["FR0086",43936,46127,0.055,15,10,4,null,null],["FR0087",43876,47894,0.065,15,8,2,null,null],["FR0088",44180,49841,0.0625,15,6,12,null,null],["FR0089",44058,55380,0.06875,15,2,8,null,null],["FR0090",44301,46492,0.05125,15,10,4,44301,44484],["FR0091",44301,48319,0.06375,15,10,4,44301,44484],["FR0092",44362,52032,0.07125,15,12,6,44362,44545],["FR0093",44392,50236,0.06375,15,1,7,null,null],["FR0094",44576,46767,0.056,15,7,1,null,null],["FR0095",44788,46980,0.06375,15,2,8,null,null],["FR0096",44788,48625,0.07,15,2,8,null,null],["FR0097",44727,52397,0.07125,15,12,6,null,null],["FR0098",44727,50571,0.07125,15,12,6,44727,44910],["FR0099",44941,47133,0.064,15,7,1,null,null],["FR0100",45153,48990,0.06625,15,2,8,45153,45337],["FR0101",45214,47223,0.06875,15,4,10,45214,45397],["FR0102",45306,56445,0.06875,15,1,7,null,null],["FR0103",45512,49505,0.0675,15,1,7,null,null],["FR0104",45525,47679,0.065,15,1,7,null,null],["FR0106",45298,51363,0.07125,15,2,8,null,null],["FR0107",45298,53189,0.07125,15,2,8,null,null],["FR0108",45867,49780,0.065,15,4,10,null,null],["FR0109",45881,47922,0.05875,15,9,3,null,null],["FR0110",46238,48349,0.0725,15,11,5,null,null],["USDFR0003",44576,48228,0.03,15,7,1,null,null],["INDOIS26",42458,46110,0.0455,29,9,3,42458,42642],["INDOIS26NEW",44356,46182,0.015,9,12,6,44356,44539],["INDOIS27",42823,46475,0.0415,29,9,3,42823,43007],["INDOIS27NEW",44718,46544,0.044,6,12,6,44718,44901],["INDOIS28",43160,46813,0.044000000000000004,1,9,3,43160,43344],["INDOIS28NEW",45245,47072,0.054,15,5,11,45245,45427],["INDOIS29",43516,47169,0.044500000000000005,20,8,2,43516,43697],["INDOIS29NEW",45475,47301,0.051,2,1,7,45475,45659],["INDOIS30",44005,47657,0.028,23,12,6,44005,44188],["INDOIS30NEW",45621,47628,0.05,25,5,11,45621,45802],["INDOIS30NEWNEW",45861,47687,0.0455,23,1,7,45861,46045],["INDOIS30NEWNEWNEW",45992,47818,0.045,1,6,12,45992,46174],["INDOIS31",44356,48008,0.0255,9,12,6,44356,44539],["INDOIS32",44719,48371,0.047,6,12,6,44718,44901],["INDOIS33",45245,48898,0.056,15,5,11,45245,45427],["INDOIS34",45475,49127,0.052,2,1,7,45475,45659],["INDOIS34NEW",45621,49273,0.0525,25,5,11,45621,45802],["INDOIS35",45861,49513,0.052,23,1,7,45861,46045],["INDOIS35NEW",45992,49644,0.05,1,6,12,45992,46174],["INDOIS50",44005,54962,0.038,23,12,6,44005,44188],["INDOIS51",44356,55313,0.0355,9,12,6,44356,44539],["INDOIS54",45475,56432,0.055,2,1,7,45475,45659],["INDOIS54NEW",45621,56578,0.0565,25,5,11,45621,45802],["INDON26",42346,46030,0.0475,8,7,1,42346,42559],["INDON27",42712,46395,0.0435,8,7,1,42712,42924],["INDON27NEW",42934,46586,0.0385,18,1,7,42934,43118],["INDON27NEWNEW",44824,46650,0.0415,20,3,9,44824,45005],["INDON28",43080,46763,0.035,11,7,1,43080,43292],["INDON28NEW",43214,46867,0.040999999999999995,24,10,4,43214,43397],["INDON28NEWNEW",44937,46763,0.0455,11,7,1,44937,45118],["INDON29",43445,47160,0.0475,11,8,2,43445,43688],["INDON29NEW",43634,47379,0.034,18,3,9,43634,43908],["INDON29NEWNEW",45301,47187,0.044,10,9,3,45301,45545],["INDON30",43844,47528,0.0285,14,8,2,43844,44057],["INDON30NEW",43936,47771,0.0385,15,10,4,43936,44119],["INDON30NEWNEW",45672,47498,0.0525,15,7,1,45672,45853],["INDON31",44208,47919,0.0185,12,9,3,44208,44451],["INDON31NEW",44405,48057,0.0215,28,1,7,44405,44589],["INDON31NEWNEW",45946,47954,0.043,16,4,10,45946,46128],["INDON31NEWNEWNEW",46043,47900,0.0435,21,8,2,46043,46255],["INDON31NEWNEWNEWNEW",46175,47997,0.0503,29,11,5,46171,46355],["INDON32",44651,48304,0.0355,30,9,1,44651,44834],["INDON32NEW",44824,48477,0.0465,20,3,9,44824,45005],["INDON33",44937,48590,0.0485,11,7,1,44937,45118],["INDON34",45301,48985,0.047,10,8,2,45301,45514],["INDON34NEW",45545,49197,0.0475,10,3,9,45545,45726],["INDON35",38637,49594,0.085,12,4,10,38637,38819],["INDON35NEW",45672,49324,0.056,15,7,1,45672,45853],["INDON36",45946,49781,0.049,16,4,10,45946,46128],["INDON36NEW",46043,49726,0.0495,21,8,2,46043,46255],["INDON36NEWNEW",46175,49824,0.0569,29,11,5,46171,46355],["INDON37",39127,50088,0.06625,17,8,2,39130,39311],["INDON38",39464,50422,0.0775,17,7,1,39464,39646],["INDON42",40925,51883,0.0525,17,7,1,40925,41107],["INDON43",41379,52336,0.04625,15,10,4,41379,41562],["INDON44",41654,52611,0.0675,15,7,1,41654,41835],["INDON45",42019,52977,0.05125,15,7,1,42019,42200],["INDON46",42346,53335,0.059500000000000004,8,7,1,42346,42559],["INDON47",42712,53700,0.0525,8,7,1,42712,42924],["INDON47NEW",42934,53891,0.0475,18,1,7,42934,43118],["INDON48",43080,54068,0.0435,11,7,1,43080,43292],["INDON49",43445,54465,0.0535,11,8,2,43445,43688],["INDON49NEW",43768,54726,0.037000000000000005,30,4,10,43768,43951],["INDON50",43844,54833,0.035,14,8,2,43844,44057],["INDON50NEW",43936,55076,0.042,15,10,4,43936,44119],["INDON51",44208,55224,0.0305,12,9,3,44208,44451],["INDON52",44651,55609,0.043,30,9,3,44651,44834],["INDON52NEW",44824,55782,0.0545,20,3,9,44824,45005],["INDON53",44937,55895,0.0565,11,7,1,44937,45118],["INDON54",45301,56290,0.051,10,8,2,45301,45514],["INDON54NEW",45545,56502,0.0515,10,3,9,45545,45726],["INDON56",46043,57031,0.05475,21,8,2,46043,46255],["INDON61",44462,59072,0.032,23,3,9,44462,44643],["INDON70",43936,62198,0.0445,15,10,4,43936,44119],["INDON71",44208,62529,0.0335,12,9,3,44208,44451],["ORI030T3",46239,47314,0.069,15,9,10,null,null],["ORI030T6",46239,48410,0.07,15,9,10,null,null],["ORI029T3",46078,47164,0.0545,15,4,5,null,null],["ORI029T6",46078,48259,0.058,15,4,5,null,null],["ORI028T3",45959,47041,0.0535,15,12,1,null,null],["ORI028T6",45959,48136,0.0565,15,12,1,null,null],["ORI027T3",45714,46798,0.0665,15,4,5,null,null],["ORI027T6",45714,47894,0.0675,15,4,5,null,null],["ORI026T3",45595,46675,0.063,15,12,1,null,null],["ORI026T6",45595,47771,0.064,15,12,1,null,null],["ORI025T3",45350,46433,0.0625,15,4,5,null,null],["ORI025T6",45350,47529,0.064,15,4,5,null,null],["ORI024T3",45238,46310,0.061,15,12,1,null,null],["ORI024T6",45238,47406,0.0635,15,12,1,null,null],["ORI023T3",45133,46218,0.059,15,9,10,null,null],["ORI023T6",45133,47314,0.061,15,9,10,null,null],["ORI022",44860,45945,0.0595,15,12,1,null,null],["PBS003",40923,46402,0.06,15,7,1,null,null],["PBS004",40955,50086,0.061,15,8,2,null,null],["PBS012",42397,48167,0.08875,15,5,11,null,null],["PBS029",44089,49018,0.06375,15,3,9,null,null],["PBS030",44392,46949,0.05875,15,1,7,null,null],["PBS032",44392,46218,0.04875,15,1,7,null,null],["PBS033",44545,53858,0.0675,15,6,12,null,null],["PBS034",44545,50936,0.065,15,6,12,null,null],["PBS035",44635,51940,0.0675,15,9,3,null,null],["PBS036",44788,45884,0.05375,15,2,8,null,null],["PBS037",44819,49749,0.06875,15,3,11,null,null],["PBS038",45267,54772,0.06875,15,12,6,null,null],["PBS040",45792,47802,0.05,15,11,5,null,null],["SR024T3",46134,47187,0.0555,10,5,6,null,null],["SR024T5",46134,47917,0.059,10,5,6,null,null],["SR023T3",45922,47036,0.058,10,11,12,null,null],["SR023T5",45922,47766,0.0595,10,11,12,null,null],["SR022T3",45833,46914,0.0645,10,8,9,null,null],["SR022T5",45833,47644,0.0655,10,8,9,null,null],["SR021T3",45560,46640,0.0635,10,11,12,null,null],["SR021T5",45560,47371,0.0645,10,11,12,null,null],["SR020T3",45385,46456,0.063,10,5,6,null,null],["SR020T5",45385,47187,0.064,10,5,6,null,null],["SR019T3",45196,46275,0.0595,10,11,12,null,null],["SR019T5",45196,47006,0.061,10,11,12,null,null],["SR018T3",45021,46091,0.0625,10,5,6,null,null],["SR018T5",45021,46822,0.064,10,5,6,null,null],["ST016T2",46183,46914,0.0605,10,7,8,null,null],["ST016T4",46183,47613,0.0625,10,7,8,null,null]]""")

TAX = 0.10
EPOCH = dt.date(1899, 12, 30)
MONTHS_ID = ["Januari", "Februari", "Maret", "April", "Mei", "Juni",
             "Juli", "Agustus", "September", "Oktober", "November", "Desember"]

INK = "#0f2233"
MUTED = "#5c7085"
LINE = "#dbe3ea"
ACCENT = "#0e7c6b"
LOSS = "#b4342a"
MARKER = "#ffd400"
SANS = "DejaVu Sans"
MONO = "DejaVu Sans Mono"

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
    if serial is None:
        return "N/A"
    d = to_date(serial)
    return f"{d.day:02d}-{_MON_EN[d.month - 1]}-{str(d.year)[2:]}"

def _days_in_month(y: int, m: int) -> int:
    leap = (y % 4 == 0 and y % 100 != 0) or y % 400 == 0
    return [31, 29 if leap else 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31][m - 1]

def edate(serial, months: int) -> int:
    d = to_date(serial)
    total = d.year * 12 + (d.month - 1) + months
    y, m = total // 12, total % 12 + 1
    return to_serial(dt.date(y, m, min(d.day, _days_in_month(y, m))))

def xdate(y: int, m: int, day: int) -> int:
    y2, m2 = y + (m - 1) // 12, (m - 1) % 12 + 1
    return to_serial(dt.date(y2, m2, 1)) + (day - 1)

def datedif_m(a, b) -> int:
    if b < a:
        return -datedif_m(b, a)
    da, db = to_date(a), to_date(b)
    n = (db.year - da.year) * 12 + (db.month - da.month)
    if db.day < da.day:
        n -= 1
    return n

def days360_eu(a, b) -> int:
    da, db = to_date(a), to_date(b)
    d1 = 30 if da.day == 31 else da.day
    d2 = 30 if db.day == 31 else db.day
    return (db.year - da.year) * 360 + (db.month - da.month) * 30 + (d2 - d1)

def days360_us(a, b) -> int:
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
    f = 10 ** digits
    v = x * f
    return (-math.floor(-v + 0.5) if v < 0 else math.floor(v + 0.5)) / f

def bbg_round(x):
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
    for row in PRODUCTS_RAW:
        if len(row) < 9:
            row = row + [None] * (9 - len(row))
        code, issued, maturity, coupon, cday, m1, m2, fad, fcd = row[:9]
        prefix = code[:2]
        semi = prefix in ("FR", "IN", "PB", "US")
        currency = "USD" if prefix in ("IN", "US") else "IDR"
        idx[code] = Meta(
            code=code, issued=issued, maturity=maturity, coupon=coupon,
            cday=cday, m1=m1, m2=m2, first_accrual=fad, first_coupon=fcd,
            prefix=prefix,
            frequency="Semi Annually" if semi else "Monthly",
            freq_months=6 if semi else 1,
            currency=currency,
            unit_value=1_000_000 if currency == "IDR" else 1000,
        )
    return idx

PRODUCT_INDEX = _build_index()
PRODUCT_CODES = sorted(PRODUCT_INDEX)

def product_meta(code):
    return PRODUCT_INDEX.get((code or "").strip().upper())

class SimError(Exception):
    pass

def _bracket_coupons(meta: Meta, settlement: int):
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
    units = nominal / meta.unit_value
    per_period = 12 if meta.freq_months == 1 else 2

    if is_perdana:
        days = 0
    elif meta.currency == "IDR" or meta.prefix == "US":
        days = settlement - last_coupon
    else:
        days = days360_eu(last_coupon, settlement)

    if is_perdana:
        per_unit = 0.0
    elif meta.currency == "IDR" or meta.prefix == "US":
        per_unit = (meta.unit_value * meta.coupon / per_period
                    * (settlement - last_coupon) / (next_coupon - last_coupon))
    else:
        raw = days360_eu(last_coupon, settlement) / 360 * nominal * meta.coupon
        per_unit = xround(raw, 2) if round_accrual else raw

    rounded = bbg_round(per_unit) if meta.currency == "IDR" else per_unit

    if meta.currency == "IDR":
        accrued = units * rounded
    elif meta.prefix == "US":
        accrued = units * (xround(per_unit, 2) if round_accrual else per_unit)
    else:
        accrued = per_unit

    gross = nominal * px + accrued
    total = trunc(gross) if meta.currency == "IDR" else gross
    return dict(units=units, days=days, per_unit=per_unit, rounded=rounded,
                accrued=accrued, total=total)

def _coupon_stream(meta: Meta, *, last_coupon, next_coupon, second_coupon,
                   settlement, nominal, units, accrued, horizon,
                   is_perdana, months_horizon=None, net_check=None):
    fm, cur, pref, cpn_rate = meta.freq_months, meta.currency, meta.prefix, meta.coupon
    per_period = 12 if fm == 1 else 2

    span = datedif_m(edate(next_coupon, -fm), horizon)
    periods = span if (cur == "IDR" and fm == 1) else int(xround(span / 6))
    months_invested = datedif_m(settlement, horizon if months_horizon is None else months_horizon)

    base = meta.unit_value * cpn_rate / per_period
    if cur == "IDR":
        per_unit = bbg_round(base)
    elif pref == "US":
        per_unit = xround(base, 2)
    else:
        per_unit = nominal * cpn_rate / per_period

    gross_coupon = units * per_unit if cur == "IDR" or pref == "US" else per_unit
    coupon_tax = gross_coupon * TAX if cur == "IDR" or pref == "US" else 0.0
    coupon_tax_r = xround(coupon_tax)

    short_per_unit = 0.0
    if not is_perdana:
        if cur == "IDR":
            short_per_unit = bbg_round(meta.unit_value * cpn_rate / (12 / fm)
                                       * (next_coupon - settlement) / (next_coupon - last_coupon))
        elif pref == "US":
            short_per_unit = per_unit * (next_coupon - settlement) / (next_coupon - last_coupon)
    short_gross = units * (xround(short_per_unit, 2) if pref == "US" else short_per_unit)
    short_tax_r = xround(short_gross * TAX)

    one_back = edate(next_coupon, -fm)
    two_back = edate(one_back, -fm)

    first_gap = (next_coupon - settlement) if is_perdana else (next_coupon - last_coupon)
    ctype = "LONG COUPON" if first_gap > second_coupon - next_coupon else "SHORT COUPON"

    ipo_short = ipo_long = 0.0
    if is_perdana:
        if ctype == "SHORT COUPON":
            ipo_short = bbg_round(cpn_rate * meta.unit_value / 12
                                  * (next_coupon - settlement) / (next_coupon - one_back))
        else:
            ipo_long = bbg_round(cpn_rate * meta.unit_value / 12
                                 * (one_back - settlement) / (one_back - two_back))
    ipo_short_gross = units * ipo_short
    ipo_short_tax_r = xround(ipo_short_gross * TAX)

    long_per_unit = 0.0
    if not is_perdana and ctype == "LONG COUPON":
        if cur == "IDR":
            long_per_unit = bbg_round(meta.unit_value * cpn_rate / (12 / fm)
                                      * (one_back - last_coupon) / (one_back - two_back))
        elif pref == "US":
            long_per_unit = (meta.unit_value * cpn_rate / (12 / fm)
                             * (one_back - last_coupon) / (one_back - two_back))
        else:
            long_per_unit = xround(days360_eu(last_coupon, one_back) / 360 * nominal * cpn_rate, 2)

    long_tax_per_unit = 0.0
    if not is_perdana and long_per_unit != 0 and ctype == "LONG COUPON":
        if cur == "IDR":
            long_tax_per_unit = bbg_round(meta.unit_value * cpn_rate / (12 / fm)
                                          * (one_back - settlement) / (one_back - two_back))
        elif pref == "US":
            long_tax_per_unit = (meta.unit_value * cpn_rate / (12 / fm)
                                 * (one_back - settlement) / (one_back - two_back))

    long_gross = 0.0
    if ctype == "LONG COUPON":
        if is_perdana:
            long_gross = units * (ipo_long + per_unit)
        elif cur == "IDR" or pref == "US":
            long_gross = units * (long_per_unit + per_unit)
        else:
            long_gross = long_per_unit + per_unit

    if is_perdana:
        long_tax = long_gross * TAX
    elif long_tax_per_unit == 0:
        long_tax = short_gross * TAX
    elif cur == "IDR" or pref == "US":
        long_tax = units * (long_tax_per_unit + per_unit) * TAX
    else:
        long_tax = 0.0
    long_tax_r = xround(long_tax)

    if cur == "IDR":
        if is_perdana:
            first_net = long_gross - long_tax_r if ctype == "LONG COUPON" else ipo_short_gross - ipo_short_tax_r
        else:
            first_net = long_gross - long_tax_r if ctype == "LONG COUPON" else gross_coupon - short_tax_r
    elif ctype == "LONG COUPON":
        first_net = long_gross - long_tax_r if pref == "US" else long_gross
    else:
        first_net = gross_coupon - short_tax_r if pref == "US" else gross_coupon

    first_gross = gross_coupon
    if cur == "IDR" and is_perdana:
        first_gross = long_gross if ctype == "LONG COUPON" else ipo_short_gross

    check_periods, check_months = net_check if net_check else (periods, months_invested)
    if check_periods == 1 and check_months < fm:
        per_period_net = first_net
    elif cur == "IDR" or pref == "US":
        per_period_net = gross_coupon - coupon_tax_r
    else:
        per_period_net = gross_coupon

    total_net = 0.0 if periods == 0 else (periods - 1) * per_period_net + first_net - accrued
    total_gross = (periods - 1) * gross_coupon + first_gross

    return dict(periods=periods, months_invested=months_invested, per_unit=per_unit,
                gross_coupon=gross_coupon, coupon_tax=coupon_tax, coupon_tax_r=coupon_tax_r,
                short_tax_r=short_tax_r, type=ctype, first_net=first_net,
                first_gross=first_gross, per_period_net=per_period_net,
                total_net=total_net, total_gross=total_gross)

def _schedule(start, freq_months, periods, first_amount, later_amount, stop):
    rows, d = start
    rows = []
    d = start
    for i in range(min(periods, 400)):
        rows.append((i + 1, d, first_amount if i == 0 else later_amount))
        if d >= stop:
            break
        d = edate(d, freq_months)
    return rows

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
    second_c = edate(next_c, meta.freq_months)

    acc = _accrued_block(meta, last_c, next_c, settlement, nominal, px, is_perdana)

    basis = 1 if meta.currency == "IDR" else 0
    yfreq = 4 if meta.freq_months == 1 else 2
    ytm = meta.coupon if is_perdana else xyield(
        settlement, meta.maturity, meta.coupon, px * 100, 100, yfreq, basis)

    s = _coupon_stream(meta, last_coupon=last_c, next_coupon=next_c, second_coupon=second_c,
                       settlement=settlement, nominal=nominal, units=acc["units"],
                       accrued=acc["accrued"], horizon=meta.maturity, is_perdana=is_perdana)

    capital_gain = xround((1 - px) * nominal)
    cg_tax = xround(TAX * capital_gain) if meta.currency == "IDR" else 0.0
    last_cpn_tax = (s["short_tax_r"] if (s["periods"] == 1 and s["months_invested"] < meta.freq_months)
                    else s["coupon_tax_r"]) if meta.currency == "IDR" else s["coupon_tax"]
    total_tax = trunc(0 if cg_tax + last_cpn_tax < 0 else cg_tax + last_cpn_tax)
    final_cpn_net = s["gross_coupon"] - total_tax

    coupon_adj = 0.0 if s["periods"] == 0 else (
        (s["periods"] - 2) * s["per_period_net"] + s["first_net"] - acc["accrued"] + final_cpn_net)

    return dict(
        meta=meta, is_perdana=is_perdana, last_coupon=last_c, next_coupon=next_c,
        accrued_days=acc["days"], accrued_interest=acc["accrued"], amount_paid=acc["total"],
        ytm=ytm, units=acc["units"], periods=s["periods"], months=s["months_invested"],
        coupon_to_maturity=s["total_net"],
        proceeds_at_maturity=nominal + coupon_adj,
        received_at_maturity=final_cpn_net + nominal,
        principal_back=nominal, final_coupon_gross=s["gross_coupon"],
        capital_gain=capital_gain, capital_gain_tax=cg_tax,
        last_coupon_tax=last_cpn_tax, total_tax=total_tax,
        cumulative_nominal=nominal - nominal * px + coupon_adj,
        cumulative_percent=(nominal - nominal * px + coupon_adj) / (nominal * px),
        schedule=_schedule(next_c, meta.freq_months, s["periods"],
                           s["first_net"], s["per_period_net"], meta.maturity),
    )

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

    is_perdana = buy_settlement == meta.issued
    buy_last, buy_next = _resolve_coupons(meta, buy_settlement, is_perdana)
    buy_second = edate(buy_next, meta.freq_months)
    buy_acc = _accrued_block(meta, buy_last, buy_next, buy_settlement, nominal, buy_price, is_perdana)

    if _first_coupon_window(meta, sell_settlement):
        sell_last, sell_next = meta.first_accrual, meta.first_coupon
    else:
        sell_last, sell_next = _bracket_coupons(meta, sell_settlement)
    sell_acc = _accrued_block(meta, sell_last, sell_next, sell_settlement, nominal, sell_price, False)

    holding = (sell_settlement - sell_last) if buy_settlement < sell_last else (sell_settlement - buy_settlement)
    capital_gain = xround((sell_price - buy_price) * nominal)
    cg_tax = xround(TAX * capital_gain)
    accrual_for_tax = xround(
        (sell_acc["rounded"] if meta.currency == "IDR" else sell_acc["per_unit"])
        * (holding / sell_acc["days"] if sell_acc["days"] and holding < sell_acc["days"] else 1))
    accrual_tax = xround(TAX * (sell_acc["units"] * accrual_for_tax
                                if meta.currency == "IDR" else accrual_for_tax))
    total_tax = trunc(0 if cg_tax + accrual_tax < 0 else cg_tax + accrual_tax) if meta.currency == "IDR" else 0.0
    net_proceeds = (trunc(sell_acc["total"] if total_tax < 0 else sell_acc["total"] - total_tax)
                    if meta.currency == "IDR" else sell_acc["total"])

    basis = 1 if meta.currency == "IDR" else 0
    yfreq = 4 if meta.freq_months == 1 else 2
    ytm_hold = xyield(buy_settlement, meta.maturity, meta.coupon, buy_price * 100, 100, yfreq, basis)
    ytm_sell = xyield(buy_settlement, sell_settlement, meta.coupon,
                      buy_price * 100, sell_price * 100, yfreq, basis)

    common = dict(last_coupon=buy_last, next_coupon=buy_next, second_coupon=buy_second,
                  settlement=buy_settlement, nominal=nominal, units=buy_acc["units"],
                  accrued=buy_acc["accrued"], is_perdana=is_perdana)
    to_mat = _coupon_stream(meta, horizon=meta.maturity, **common)
    to_sale = _coupon_stream(meta, horizon=sell_last, months_horizon=sell_settlement,
                             net_check=(to_mat["periods"], to_mat["months_invested"]), **common)

    coupon_to_sale = (sell_acc["accrued"] - buy_acc["accrued"] - total_tax
                      if to_sale["periods"] == 0 else
                      (to_sale["periods"] - 1) * to_sale["per_period_net"] + to_sale["first_net"]
                      - buy_acc["accrued"] + sell_acc["accrued"])

    gain_sold = (net_proceeds - nominal * buy_price + coupon_to_sale
                 - sell_acc["accrued"] + total_tax)

    cg_mat = xround((1 - buy_price) * nominal)
    cg_mat_tax = xround(TAX * cg_mat) if meta.currency == "IDR" else 0.0
    last_cpn_tax = ((to_mat["short_tax_r"] if (to_mat["periods"] == 1
                     and to_mat["months_invested"] < meta.freq_months) else to_mat["coupon_tax_r"])
                    if meta.currency == "IDR" else to_mat["coupon_tax"])
    total_tax_mat = trunc(0 if cg_mat_tax + last_cpn_tax < 0 else cg_mat_tax + last_cpn_tax)
    final_cpn_net = to_mat["gross_coupon"] - total_tax_mat
    coupon_to_mat = 0.0 if to_mat["periods"] == 0 else (
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
        proceeds_if_sold=net_proceeds + coupon_to_sale - sell_acc["accrued"],
        months_if_sold=to_sale["months_invested"],
        gain_if_sold=gain_sold, pct_if_sold=gain_sold / (nominal * buy_price),
        proceeds_if_held=nominal + coupon_to_mat,
        months_if_held=to_mat["months_invested"],
        gain_if_held=nominal - nominal * buy_price + coupon_to_mat,
        pct_if_held=(nominal - nominal * buy_price + coupon_to_mat) / (nominal * buy_price),
        schedule=_schedule(buy_next, meta.freq_months, to_sale["periods"],
                           to_sale["first_net"], to_sale["per_period_net"], sell_last),
    )

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

    is_perdana2 = settlement2 == meta2.issued
    last2, next2 = _resolve_coupons(meta2, settlement2, is_perdana2)
    acc2 = _accrued_block(meta2, last2, next2, settlement2, nominal2, price2,
                          is_perdana2, round_accrual=False)

    basis2 = 1 if meta2.currency == "IDR" else 0
    yfreq2 = 4 if meta2.freq_months == 1 else 2
    ytm2 = xyield(settlement2, maturity2, meta2.coupon, price2 * 100, 100, yfreq2, basis2)

    span2 = datedif_m(edate(next2, -meta2.freq_months), maturity2)
    periods2 = span2 if meta2.freq_months == 1 else span2 // 6
    months2 = datedif_m(settlement2, maturity2)

    per_period2 = 12 if meta2.freq_months == 1 else 2
    base2 = meta2.unit_value * meta2.coupon / per_period2
    if meta2.currency == "IDR":
        cpu2 = bbg_round(base2)
    elif meta2.prefix == "US":
        cpu2 = xround(base2, 2)
    else:
        cpu2 = nominal2 * meta2.coupon / per_period2
    gross2 = acc2["units"] * cpu2 if meta2.currency == "IDR" else cpu2
    tax2_r = xround(gross2 * TAX) if meta2.currency == "IDR" else 0.0
    net2 = gross2 - tax2_r

    coupon_to_horizon2 = (acc2["accrued"] if periods2 == 0
                          else (periods2 - 1) * net2 + gross2 - acc2["accrued"])

    capital = nominal1 * buy_price1
    gain_sell1 = p1["net_proceeds"] - capital
    gain_buy2 = p1["net_proceeds"] - acc2["total"]
    gain_switch = gain_sell1 + p1["coupon_to_sale"] + gain_buy2 + coupon_to_horizon2

    return dict(
        p1=p1, meta2=meta2, is_perdana2=is_perdana2,
        currency_mismatch=p1["meta"].currency != meta2.currency,
        last_coupon2=last2, next_coupon2=next2,
        accrued_days2=acc2["days"], accrued2=acc2["accrued"], amount_paid2=acc2["total"],
        ytm2=ytm2, coupon_to_horizon2=coupon_to_horizon2,
        periods2=periods2, months2=months2,
        months_switched=datedif_m(buy_settlement1, maturity2),
        capital=capital, proceeds_sell1=p1["net_proceeds"], gain_sell1=gain_sell1,
        coupon_until_sale1=p1["coupon_to_sale"], paid2=acc2["total"], gain_buy2=gain_buy2,
        total_return_switch=gain_buy2 + nominal1 + coupon_to_horizon2,
        gain_switch=gain_switch, pct_switch=gain_switch / capital,
        top_up=p1["net_proceeds"] - acc2["total"],
        schedule2=_schedule(next2, meta2.freq_months, periods2, gross2, net2, maturity2),
    )

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
    raw = st.text_input(label, default, key=key)
    val = parse_number(raw)
    if val is not None:
        st.caption(f"= {cur_hint} {val:,.0f}".replace(",", ".").strip())
    elif raw.strip():
        st.caption("⚠️ Angka tidak dikenali")
    return raw

def _paren(v):
    if v is None:
        return "—"
    body = f"{abs(v):,.2f}"
    return f"({body})" if v < 0 else body

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
]

def render_beli_export(r, *, code, market, txn_serial, settle_serial, nominal, price_pct,
                       sim_title="SIMULASI BELI"):
    cur = r["meta"].currency
    meta = r["meta"]

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
    col_w = (W - ML - MR - gutter) / 2
    top = H - MT

    def text(x, y, s, size=9, color=INK, weight="normal", ha="left", va="center", family=SANS):
        ax.text(x, y, s, fontsize=size, color=color, weight=weight, ha=ha, va=va, family=family)

    def bar(x, y, w, h, label):
        ax.add_patch(plt.Rectangle((x, y - h), w, h, facecolor=_EX_BLACK, edgecolor="none", zorder=2))
        text(x + w / 2, y - h / 2, label, size=11, color="white", weight="bold", ha="center")

    ax.text(ML, top - 0.02, sim_title, fontsize=17, color=INK, weight="bold", family=SANS, ha="left", va="top")
    ax.text(W - MR, top - 0.10, "Kalkulator Obligasi", fontsize=9.5, color=MUTED, family=SANS, ha="right", va="top")
    ax.plot([ML, W - MR], [top - title_h + 0.06, top - title_h + 0.06], color=INK, lw=1.6)
    top -= title_h + 0.14

    y = top
    bh = 0.52
    ax.add_patch(FancyBboxPatch((ML, y - bh), col_w, bh, boxstyle="round,pad=0,rounding_size=0.02",
                                facecolor=_EX_MARKER, edgecolor=_EX_MARKER_EDGE, lw=1.2, zorder=2))
    text(ML + 0.14, y - bh / 2, code, size=20, weight="bold")
    y -= bh + 0.22

    bar(ML, y, col_w, 0.34, "Data Produk Obligasi  (NASABAH BELI)")
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
            ax.add_patch(plt.Rectangle((ML + col_w * 0.52, ry), col_w * 0.48, rh, facecolor=_EX_MARKER, edgecolor="none", zorder=1))
        ax.add_patch(plt.Rectangle((ML, ry), col_w, rh, facecolor="none", edgecolor=_EX_LINE, lw=0.6, zorder=1.5))
        text(ML + 0.10, ry + rh / 2, lab, size=8.5, weight="bold")
        text(ML + col_w - 0.10, ry + rh / 2, val, size=8.5, weight="bold", ha="right", family=MONO)
        y = ry
    y -= 0.16

    for lab, val in [
        ("Imbal Hasil hingga JT / Yield To Maturity (gross)", pct(r["ytm"], 4)),
        ("Hari Kupon Berjalan / Days of Accrued Interest", num(r["accrued_days"])),
        ("Kupon Berjalan / Accrued Interest", _paren(r["accrued_interest"])),
    ]:
        ry = y - rh
        ax.add_patch(plt.Rectangle((ML, ry), col_w, rh, facecolor=_EX_PANEL, edgecolor=_EX_LINE, lw=0.8, zorder=1))
        text(ML + 0.10, ry + rh / 2, lab, size=8, weight="bold")
        text(ML + col_w - 0.10, ry + rh / 2, val, size=8.5, weight="bold", ha="right", family=MONO)
        y = ry - 0.06
    y -= 0.14

    tbh = 0.56
    ax.add_patch(plt.Rectangle((ML, y - tbh), col_w, tbh, facecolor=_EX_BLACK, zorder=2))
    text(ML + 0.16, y - tbh / 2, "Jumlah Indikatif yang Dibayar", size=12.5, color="white", weight="bold")
    text(ML + col_w - 0.16, y - tbh / 2, _paren(r["amount_paid"]), size=13, color="white", weight="bold", ha="right", family=MONO)
    y -= tbh + 0.18

    buf = io.BytesIO()
    FigureCanvasAgg(fig).print_png(buf)
    plt.close(fig)
    return buf.getvalue()

# ---------------------------------------------------------------------------
# Simple & Easy-to-Understand Export for Simulasi Kredit (Opsi 1 / Opsi 2)
# ---------------------------------------------------------------------------
def render_kredit_export(opt_title, data_dict):
    W = 9.5  # Width for readability
    ML, MR = 0.55, 0.55
    usable_w = W - ML - MR

    # Calculate disclaimer lines and compact height (< 15% of total page length)
    wrapped_lines = []
    wrapped_lines.append("DISCLAIMER:")
    for line in DISCLAIMER_FULL:
        wrapped_lines.extend(textwrap.wrap(line, width=130))

    disc_h = 0.25 + len(wrapped_lines) * 0.15 + 0.10
    
    # Height components dynamically calculated
    top_margin = 0.45
    title_block = 1.1
    sec1_h = 0.46 + 6 * 0.38 + 0.15
    sec2_h = 0.46 + 8 * 0.38 + 0.15
    gap_before_disc = 0.30
    bottom_margin = 0.35

    H = top_margin + title_block + sec1_h + sec2_h + gap_before_disc + disc_h + bottom_margin

    fig = plt.figure(figsize=(W, H), dpi=170)
    fig.patch.set_facecolor("white")
    ax = fig.add_axes([0, 0, 1, 1])
    ax.set_xlim(0, W)
    ax.set_ylim(0, H)
    ax.axis("off")

    top = H - top_margin

    def text(x, y, s, size=10.5, color=INK, weight="normal", ha="left", va="center", family=SANS):
        ax.text(x, y, s, fontsize=size, color=color, weight=weight, ha=ha, va=va, family=family)

    text(ML, top, f"SIMULASI KREDIT OBLIGASI — {opt_title.upper()}", size=17, weight="bold", va="top")
    text(W - MR, top, f"Tanggal: {dt.date.today():%d-%b-%Y}", size=10, color=MUTED, ha="right", va="top", family=MONO)
    top -= 0.40
    ax.plot([ML, W - MR], [top, top], color=INK, lw=1.5)
    top -= 0.40

    def section_header(title, y_pos):
        ax.add_patch(plt.Rectangle((ML, y_pos - 0.32), usable_w, 0.32, facecolor=_EX_BLACK, edgecolor="none"))
        text(ML + 0.15, y_pos - 0.16, title, size=12, color="white", weight="bold")
        return y_pos - 0.46

    def add_rows(rows, y_pos):
        rh = 0.38
        for row in rows:
            lab = row[0]
            val = row[1]
            bold = row[2]
            color_val = row[3] if len(row) > 3 else INK
            indent = row[4] if len(row) > 4 else 0.0
            
            ry = y_pos - rh
            ax.add_patch(plt.Rectangle((ML, ry), usable_w, rh, facecolor="#f9fbfe", edgecolor=_EX_LINE, lw=0.6))
            text(ML + 0.15 + indent, ry + rh/2, lab, size=10.5, weight="bold" if bold else "normal")
            text(W - MR - 0.15, ry + rh/2, val, size=11, weight="bold" if bold else "normal", color=color_val, ha="right", family=MONO)
            y_pos = ry
        return y_pos - 0.15

    top = section_header("1. PARAMETER KREDIT & INVESTASI", top)
    top = add_rows([
        ("Produk Obligasi", data_dict["code"], False),
        ("Nominal Investasi", data_dict["nominal_fmt"], False),
        ("LTV (%)", f"{data_dict['ltv']*100:.1f}%", False),
        ("Plafon Kredit", data_dict["plafon_fmt"], True),
        ("Suku Bunga Kredit (% p.a)", f"{data_dict['bunga_kredit']*100:.2f}%", False),
        ("Tenor Pinjaman", f"{data_dict['tenor']} Tahun", False),
    ], top)

    top = section_header("2. ESTIMASI PENDAPATAN & BEBAN", top)
    top = add_rows([
        ("Bunga Investasi / Tahun (Nett)", data_dict["inv_nett_fmt"], True, ACCENT),
        ("  ↳ per Bulan", data_dict["inv_nett_bulan_fmt"], False, ACCENT, 0.2),
        ("  ↳ per Hari", data_dict["inv_nett_hari_fmt"], False, ACCENT, 0.2),
        ("Bunga Pinjaman / Tahun (Maksimal)", data_dict["pinj_tahun_fmt"], True, LOSS),
        ("  ↳ per Bulan", data_dict["pinj_bulan_fmt"], False, LOSS, 0.2),
        ("  ↳ per Hari", data_dict["pinj_hari_fmt"], False, LOSS, 0.2),
        ("Cicilan / Bulan (Metode Anuitas - IL)", data_dict["pmt_fmt"], True, LOSS),
        ("Total Biaya Provisi & Admin", data_dict["tot_biaya_fmt"], False, INK),
    ], top)

    # Position disclaimer box cleanly below Section 2 with gap
    top -= gap_before_disc
    box_top = top
    box_bottom = box_top - disc_h
    ax.add_patch(plt.Rectangle((ML, box_bottom), usable_w, disc_h, facecolor="white", edgecolor=INK, lw=0.8))
    
    dy = box_top - 0.18
    text(ML + 0.15, dy, "DISCLAIMER:", size=8.0, weight="bold")
    dy -= 0.18
    for line in DISCLAIMER_FULL:
        wrapped = textwrap.wrap(line, width=130)
        for i, w_line in enumerate(wrapped):
            prefix = "• " if i == 0 else "   "
            text(ML + 0.20, dy, prefix + w_line, size=6.5, color=MUTED)
            dy -= 0.15

    buf = io.BytesIO()
    FigureCanvasAgg(fig).print_png(buf)
    plt.close(fig)
    return buf.getvalue()


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


def render_jual_export(r, *, code, buy_txn, buy_settle, nominal, buy_price,
                       sell_txn, sell_settle, sell_price, sim_title="SIMULASI JUAL"):
    cur = r["meta"].currency
    meta = r["meta"]
    disc_h = 0.26 + len(DISCLAIMER_FULL) * 0.208 + 0.14
    W, H = 12.4, 11.35 + disc_h
    fig, ax, top = _new_figure(sim_title, W, H)

    ML, MR, gutter = 0.35, 0.45, 0.45
    col_w = (W - ML - MR - gutter) / 2
    rcol_x = ML + col_w + gutter
    s = _Sheet(ax, ML, col_w)

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


def render_switching_export(r, *, code1, bs1, ss1, nom1, bp1, sp1,
                            code2, s2, m2, nom2, p2, sim_title="SIMULASI SWITCHING"):
    p1 = r["p1"]
    cur1 = p1["meta"].currency
    cur2 = r["meta2"].currency
    meta1, meta2 = p1["meta"], r["meta2"]
    disc_h = 0.26 + len(DISCLAIMER_FULL) * 0.208 + 0.14
    W = 12.6
    _left_depth = 15.32
    H = 0.94 + _left_depth + disc_h + 0.28
    fig, ax, top = _new_figure(sim_title, W, H)

    ML, MR, gutter = 0.35, 0.35, 0.45
    col_w = (W - ML - MR - gutter) / 2
    rcol_x = ML + col_w + gutter
    s = _Sheet(ax, ML, col_w)

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
    left_bottom = y - 0.22

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

    _disclaimer_box(ax, W, ML, MR, disc_h, lines=DISCLAIMER_FULL, top=left_bottom)
    buf = io.BytesIO()
    FigureCanvasAgg(fig).print_png(buf)
    plt.close(fig)
    return buf.getvalue()


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
  div[data-testid="stForm"] { border-color:#dbe3ea; }
  div[data-testid="stMarkdownContainer"]:empty { display:none; }
  .ko-card:empty { display:none; }
  .stDivider { margin:0.4rem 0; }

  .stTextInput input, .stNumberInput input, .stDateInput input {
      background-color: #ffffff !important;
      color: #0f2233 !important;
  }
  div[data-baseweb="select"] > div {
      background-color: #ffffff !important;
      border-color: #dbe3ea !important;
      box-shadow: 0 1px 2px rgba(0,0,0,0.05);
      border-radius: 4px;
  }
  div[data-baseweb="select"] > div * {
      color: #0f2233 !important;
  }
  div[data-baseweb="select"] svg {
      fill: #0f2233 !important;
  }
  div[data-baseweb="input"], div[data-baseweb="base-input"] {
      background-color: #ffffff !important;
  }
</style>
"""


def _row_html(label, value, tone=None, indent=False, strong=False):
    cls = "ko-row" + (" ko-ind" if indent else "") + (" ko-strong" if strong else "")
    vcls = "v" + (f" {tone}" if tone in ("gain", "loss") else "")
    return f'<div class="{cls}"><span>{label}</span><span class="{vcls}">{value}</span></div>'


def card(heading, rows):
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
                f'<span class="{sub}</span></div>', unsafe_allow_html=True)


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
    stamp = f"{dt.date.today():%d-%b-%Y}"
    safe = "".join(c for c in f"{sim_label}_{seri}_{stamp}" if c not in '\\/:*?"<>|')
    return f"{safe}.jpg"


def export_section(png_factory, sim_label, seri, state_key):
    import base64
    made = st.session_state.get(state_key, False)
    label = "Perbarui gambar" if made else "Buat gambar untuk dibagikan"
    if st.button(label, key=f"{state_key}_btn", type="primary", **FULL_WIDTH):
        made = True
        st.session_state[state_key] = True

    if not made:
        st.caption("Tekan tombol di atas untuk membuat gambar simulasi yang bisa diunduh atau dibagikan ke nasabah.")
        return

    try:
        jpg = png_to_jpg(png_factory())
    except Exception as exc:
        st.warning(f"Gambar tidak dapat dibuat: {exc}")
        return

    b64 = base64.b64encode(jpg).decode("ascii")
    st.markdown(
        f'<img src="data:image/jpeg;base64,{b64}" '
        'style="width:100%;border:1px solid #dbe3ea;border-radius:8px;margin:6px 0;" '
        'alt="Pratinjau simulasi" />',
        unsafe_allow_html=True,
    )
    st.download_button(
        "Unduh gambar (JPG)",
        data=jpg,
        file_name=export_filename(sim_label, seri),
        mime="image/jpeg",
        **FULL_WIDTH,
    )


def tab_beli():
    left, right = st.columns([5, 6], gap="large")
    with left:
        st.markdown('<div class="ko-cap">Data produk</div>', unsafe_allow_html=True)
        code = st.selectbox("Kode obligasi", PRODUCT_CODES, index=PRODUCT_CODES.index("FR0110"), key="b_code")
        meta = product_meta(code)
        show_meta(meta)
        market = st.selectbox("Jenis transaksi", ["Pasar Sekunder", "Pasar Perdana"], key="b_market")
        txn = st.date_input("Tanggal transaksi", dt.date.today() - dt.timedelta(days=2), key="b_txn")
        settle = st.date_input("Tanggal setelmen", dt.date.today(), key="b_settle")
        nominal = parse_number(money_input("Nilai nominal", "", "b_nominal", cur_hint="Rp"))
        price_in = st.number_input("Harga nasabah beli (%)", value=0.0, step=0.05, format="%.4f", key="b_price")

    try:
        r = simulate_beli(code, market, to_serial(settle), nominal, None if price_in is None else price_in / 100)
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
            big("Total keuntungan", money(r["cumulative_nominal"], cur), f"{pct(r['cumulative_percent'])} kumulatif", tone_of(r["cumulative_nominal"]))
        with b:
            big("Pengembalian hingga JT", money(r["proceeds_at_maturity"], cur), f"{num(r['periods'])} periode")
        card(None, [
            ("Kupon diterima hingga JT", money(r["coupon_to_maturity"], cur)),
            ("Nominal diterima saat JT", money(r["received_at_maturity"], cur)),
            ("Pengembalian pokok", money(r["principal_back"], cur), None, True),
            ("Capital gain / loss", money(r["capital_gain"], cur), tone_of(r["capital_gain"]), True),
            ("Total pajak", money(r["total_tax"], cur), None, True),
        ])

    show_schedule(r["schedule"], cur, "Jadwal kupon")
    st.divider()
    st.markdown('<div class="ko-cap">Gambar untuk dibagikan</div>', unsafe_allow_html=True)
    export_section(
        lambda: render_beli_export(r, code=code, market=market, txn_serial=to_serial(txn),
                                   settle_serial=to_serial(settle), nominal=nominal, price_pct=price_in),
        "SIMULASI BELI", code, "b_export")


def tab_jual():
    left, right = st.columns([5, 6], gap="large")
    with left:
        st.markdown('<div class="ko-cap">Saat nasabah beli</div>', unsafe_allow_html=True)
        code = st.selectbox("Kode obligasi", PRODUCT_CODES, index=PRODUCT_CODES.index("FR0110"), key="j_code")
        meta = product_meta(code)
        show_meta(meta)
        buy_settle = st.date_input("Tanggal setelmen beli", dt.date.today() - dt.timedelta(days=365), key="j_bs")
        nominal = parse_number(money_input("Nilai nominal", "", "j_nominal"))
        buy_price = st.number_input("Harga nasabah beli (%)", value=0.0, step=0.05, format="%.4f", key="j_bp")
        st.markdown('<div class="ko-cap">Saat nasabah jual</div>', unsafe_allow_html=True)
        sell_settle = st.date_input("Tanggal setelmen jual", dt.date.today(), key="j_ss")
        sell_price = st.number_input("Harga nasabah jual (%)", value=0.0, step=0.05, format="%.4f", key="j_sp")

    try:
        r = simulate_jual(code, to_serial(buy_settle), nominal, buy_price / 100, to_serial(sell_settle), sell_price / 100)
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
            big("Total keuntungan", money(r["gain_if_sold"], cur), f"{pct(r['pct_if_sold'])}", tone_of(r["gain_if_sold"]))
        with b:
            big("Pengembalian diterima", money(r["proceeds_if_sold"], cur), f"YTM {pct(r['ytm_sell'], 4)}")

    show_schedule(r["schedule"], cur, "Jadwal kupon hingga penjualan")
    st.divider()
    st.markdown('<div class="ko-cap">Gambar untuk dibagikan</div>', unsafe_allow_html=True)
    export_section(
        lambda: render_jual_export(r, code=code, buy_txn=to_serial(buy_settle) - 2, buy_settle=to_serial(buy_settle),
                                   nominal=nominal, buy_price=buy_price, sell_txn=to_serial(sell_settle) - 2,
                                   sell_settle=to_serial(sell_settle), sell_price=sell_price),
        "SIMULASI JUAL", code, "j_export")


def tab_switching():
    left, right = st.columns([5, 6], gap="large")
    with left:
        st.markdown('<div class="ko-cap">Produk 1 — beli</div>', unsafe_allow_html=True)
        code1 = st.selectbox("Kode obligasi Produk 1", PRODUCT_CODES, index=PRODUCT_CODES.index("FR0110"), key="s_c1")
        show_meta(product_meta(code1))
        bs1 = st.date_input("Setelmen beli Produk 1", dt.date.today() - dt.timedelta(days=365), key="s_bs1")
        nom1 = parse_number(money_input("Nilai nominal Produk 1", "", "s_n1"))
        bp1 = st.number_input("Harga beli Produk 1 (%)", value=0.0, step=0.05, format="%.4f", key="s_bp1")
        st.markdown('<div class="ko-cap">Produk 1 — jual</div>', unsafe_allow_html=True)
        ss1 = st.date_input("Setelmen jual Produk 1", dt.date.today(), key="s_ss1")
        sp1 = st.number_input("Harga jual Produk 1 (%)", value=0.0, step=0.05, format="%.4f", key="s_sp1")

        st.markdown('<div class="ko-cap">Produk 2 — beli</div>', unsafe_allow_html=True)
        code2 = st.selectbox("Kode obligasi Produk 2", PRODUCT_CODES, index=PRODUCT_CODES.index("FR0100"), key="s_c2")
        meta2 = product_meta(code2)
        show_meta(meta2)
        s2 = st.date_input("Setelmen Produk 2", dt.date.today(), key="s_s2")
        m2 = st.date_input("Jatuh tempo Produk 2", to_date(meta2.maturity), key="s_m2")
        nom2 = parse_number(money_input("Nilai nominal Produk 2", "", "s_n2"))
        p2 = st.number_input("Harga beli Produk 2 (%)", value=0.0, step=0.05, format="%.4f", key="s_p2")

    try:
        r = simulate_switching(code1, to_serial(bs1), nom1, bp1 / 100, to_serial(ss1), sp1 / 100,
                               code2, to_serial(s2), to_serial(m2), nom2, p2 / 100)
    except SimError as exc:
        with right:
            st.info(str(exc))
        return

    cur1 = r["p1"]["meta"].currency
    with right:
        if r["currency_mismatch"]:
            st.error("Perbedaan mata uang antara Produk 1 & 2.")
        st.markdown('<div class="ko-cap">Perbandingan dua pilihan</div>', unsafe_allow_html=True)
        a, b = st.columns(2)
        with a:
            big("Ditahan hingga JT", money(r["p1"]["gain_if_held"], cur1), tone=tone_of(r["p1"]["gain_if_held"]))
        with b:
            big("Dialihkan ke Produk 2", money(r["gain_switch"], cur1), tone=tone_of(r["gain_switch"]))

    show_schedule(r["schedule2"], r["meta2"].currency, "Jadwal kupon Produk 2")
    st.divider()
    st.markdown('<div class="ko-cap">Gambar untuk dibagikan</div>', unsafe_allow_html=True)
    if not r["currency_mismatch"]:
        export_section(
            lambda: render_switching_export(r, code1=code1, bs1=to_serial(bs1), ss1=to_serial(ss1),
                                            nom1=nom1, bp1=bp1, sp1=sp1, code2=code2, s2=to_serial(s2),
                                            m2=to_serial(m2), nom2=nom2, p2=p2),
            "SIMULASI SWITCHING", f"{code1}-{code2}", "s_export")


def tab_kredit():
    st.markdown('<div class="ko-cap" style="margin-bottom:12px;">Simulasi Kredit Agunan Obligasi</div>', unsafe_allow_html=True)
    left, right = st.columns(2, gap="large")

    def _apply_ltv_logic(idx, default_code=None):
        code = st.session_state.get(f"k_prod_{idx}", default_code)
        if not code:
            return
        meta = product_meta(code)
        months_to_mat = datedif_m(to_serial(dt.date.today()), meta.maturity)
        years = months_to_mat / 12.0
        if meta.currency == "IDR":
            val = 90.0 if years < 5.0 else (85.0 if years <= 10.0 else 80.0)
        else:
            val = 80.0 if years <= 10.0 else 70.0
        st.session_state[f"k_ltv_{idx}"] = val

    def render_kredit_col(idx, def_nom, def_prod, def_tenor):
        st.markdown(f'<div class="ko-cap" style="margin-top:0px;">Opsi {idx}</div>', unsafe_allow_html=True)
        if f"k_ltv_{idx}" not in st.session_state:
            _apply_ltv_logic(idx, default_code=def_prod)

        code = st.selectbox("Produk", PRODUCT_CODES, index=PRODUCT_CODES.index(def_prod), 
                            key=f"k_prod_{idx}", on_change=_apply_ltv_logic, args=(idx,))
        meta = product_meta(code)
        show_meta(meta)
        cur = meta.currency

        nom_raw = money_input("Nominal Investasi", def_nom, f"k_nom_{idx}", cur_hint="")
        nom = parse_number(nom_raw) or 0.0
        
        colA, colB = st.columns(2)
        with colA:
            ltv = st.number_input("LTV (%)", step=1.0, format="%.1f", key=f"k_ltv_{idx}") / 100.0
            bunga_kredit = st.number_input("Suku Bunga Kredit (% p.a)", value=7.18, step=0.1, format="%.2f", key=f"k_bk_{idx}") / 100.0
        with colB:
            prov_pct = st.number_input("Provisi (%)", value=0.0, step=0.1, format="%.2f", key=f"k_prov_{idx}") / 100.0
            admin_raw = money_input("Admin", "", f"k_adm_{idx}", cur_hint=cur)
            admin = parse_number(admin_raw) or 0.0
        
        tenor = st.number_input("Tenor Pinjaman (Tahun)", value=def_tenor, step=1, key=f"k_tenor_{idx}")
        
        plafon = nom * ltv
        prov_nom = plafon * prov_pct
        tot_biaya = prov_nom + admin
        
        inv_gross = nom * meta.coupon
        inv_nett = inv_gross * (1 - TAX)
        inv_nett_bulan = inv_nett / 12.0
        inv_nett_hari = inv_nett / 365.0
        
        pinj_tahun = plafon * bunga_kredit
        pinj_bulan = pinj_tahun / 12.0
        pinj_hari = pinj_tahun / 365.0
        
        if tenor > 0:
            n = int(tenor * 12)
            if bunga_kredit > 0:
                r_rate = bunga_kredit / 12
                pmt = (plafon * r_rate * math.pow(1+r_rate, n)) / (math.pow(1+r_rate, n) - 1)
            else:
                pmt = plafon / n
        else:
            pmt = 0
            
        st.markdown('<div class="ko-cap" style="margin-top:16px;">Ringkasan Parameter</div>', unsafe_allow_html=True)
        card(None, [
            ("Kupon Produk", pct(meta.coupon, 3)),
            ("Plafon Kredit", money(plafon, cur), None, False, True)
        ])
        
        st.markdown('<div class="ko-cap">Simulasi Pendapatan & Beban</div>', unsafe_allow_html=True)
        card(None, [
            ("Bunga Investasi / Tahun (Nett)", money(inv_nett, cur), "gain", False, True),
            ("↳ per Bulan", money(inv_nett_bulan, cur), "gain", True, False),
            ("↳ per Hari", money(inv_nett_hari, cur), "gain", True, False),
            ("Bunga Pinjaman / Tahun (Maksimal)", money(pinj_tahun, cur), "loss", False, True),
            ("↳ per Bulan", money(pinj_bulan, cur), "loss", True, False),
            ("↳ per Hari", money(pinj_hari, cur), "loss", True, False),
            ("Cicilan / Bulan (Metode Anuitas - IL)", money(pmt, cur), "loss", False, True),
            ("Total Biaya Provisi & Admin", money(tot_biaya, cur), None, False, False)
        ])
        
        export_data = {
            "code": code,
            "currency": cur,
            "nominal_fmt": money(nom, cur),
            "ltv": ltv,
            "plafon_fmt": money(plafon, cur),
            "bunga_kredit": bunga_kredit,
            "tenor": tenor,
            "inv_nett_fmt": money(inv_nett, cur),
            "inv_nett_bulan_fmt": money(inv_nett_bulan, cur),
            "inv_nett_hari_fmt": money(inv_nett_hari, cur),
            "pinj_tahun_fmt": money(pinj_tahun, cur),
            "pinj_bulan_fmt": money(pinj_bulan, cur),
            "pinj_hari_fmt": money(pinj_hari, cur),
            "pmt_fmt": money(pmt, cur),
            "tot_biaya_fmt": money(tot_biaya, cur),
        }

        st.divider()
        st.markdown(f'<div class="ko-cap">Unduh Gambar Opsi {idx}</div>', unsafe_allow_html=True)
        export_section(
            lambda: render_kredit_export(f"Opsi {idx}", export_data),
            f"SIMULASI_KREDIT_OPSI_{idx}", code, f"k_export_{idx}")

    with left:
        render_kredit_col(1, "", "INDOIS30NEWNEW", 0)
    with right:
        render_kredit_col(2, "", "FR0082", 0)


def main():
    st.set_page_config(page_title="Kalkulator Obligasi", page_icon="🧮",
                       layout="wide", initial_sidebar_state="collapsed")
    st.markdown(CSS, unsafe_allow_html=True)
    st.markdown(
        '<div class="ko-head"><div class="ko-mark"></div><div>'
        '<p class="ko-title">Kalkulator Obligasi</p>'
        '<p class="ko-sub">Simulasi beli, jual, dan switching · versi 2.4.8</p>'
        "</div></div>",
        unsafe_allow_html=True,
    )

    beli, jual, switching, kredit = st.tabs(["Simulasi Beli", "Simulasi Jual", "Simulasi Switching", "Simulasi Kredit"])
    
    with beli:
        tab_beli()
    with jual:
        tab_jual()
    with switching:
        tab_switching()
    with kredit:
        tab_kredit()


if __name__ == "__main__":
    main()
