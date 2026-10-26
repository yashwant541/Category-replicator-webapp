"""Dummy data for the Category Replicator: loan exposures split into 4 sheets, mapping-driven
formulas, and a different set of derived columns on every output sheet.

    python dev/make_dummy_data.py            -> dev/sample_data/*.xlsx

Files
  1_input_exposures.xlsx      raw input, 4 sheets (split by Segment / Region), NO derived columns
  2_mapping_tables.xlsx       5 mapping sheets (Ratings, Products, Countries, FX, Tenor)
  3_expected_output.xlsx      4 sheets with live Excel formulas (+ cached values) and the mapping sheets
  4_new_input_exposures.xlsx  fresh raw data in the same layout, to run the learned system on
  5_new_expected_output.xlsx  what the new input should produce (to compare against)

Formula cells carry cached values (as a file saved by Excel would), so the learner can verify itself.
"""
import datetime as dt
import io
import pathlib
import random
import re
import sys
import zipfile
from xml.sax.saxutils import escape

import openpyxl

OUT = pathlib.Path(__file__).resolve().parent / "sample_data"

# ------------------------------------------------------------------ mapping tables
RATINGS = [("AAA", 0.0003, 0.20), ("AA", 0.0010, 0.20), ("A", 0.0030, 0.50), ("BBB", 0.0100, 1.00),
           ("BB", 0.0400, 1.00), ("B", 0.1000, 1.50)]
PRODUCTS = [("Mortgage", 0.15, 0.30), ("Credit Card", 0.65, 0.00), ("Personal Loan", 0.45, 0.05),
            ("Term Loan", 0.35, 0.10), ("Revolver", 0.40, 0.15), ("Trade Finance", 0.30, 0.20)]
COUNTRIES = [("UK", "EMEA", "Low"), ("DE", "EMEA", "Low"), ("FR", "EMEA", "Low"), ("AE", "EMEA", "Medium"),
             ("EG", "EMEA", "High"), ("SG", "APAC", "Low"), ("IN", "APAC", "Medium"), ("JP", "APAC", "Low"),
             ("ID", "APAC", "High"), ("US", "AMER", "Low"), ("BR", "AMER", "High"), ("CA", "AMER", "Low")]
FX = [("USD", 1.0), ("EUR", 1.08), ("GBP", 1.27), ("AED", 0.2723), ("SGD", 0.74), ("JPY", 0.0067), ("INR", 0.012)]
CCY_BY_COUNTRY = {"UK": "GBP", "DE": "EUR", "FR": "EUR", "AE": "AED", "EG": "USD", "SG": "SGD", "IN": "INR",
                  "JP": "JPY", "ID": "USD", "US": "USD", "BR": "USD", "CA": "USD"}

MAP_SHEETS = {
    "Ratings": (["Rating", "PD", "RiskWeight"], RATINGS),
    "Products": (["Product", "LGD", "CollateralHaircut"], PRODUCTS),
    "Countries": (["Country", "Region", "CountryRisk"], COUNTRIES),
    "FX": (["Currency", "RateToUSD"], FX),
}
rating = {r[0]: r for r in RATINGS}
product = {r[0]: r for r in PRODUCTS}
country = {r[0]: r for r in COUNTRIES}
fx = dict(FX)

BASE = ["LoanID", "Customer", "Segment", "Region", "Product", "Rating", "Country", "Currency",
        "Exposure", "Collateral", "Maturity"]

# sheet name -> (segment, regions allowed, products allowed)
SHEETS = {
    "Retail_EMEA": ("Retail", ["EMEA"], ["Mortgage", "Credit Card", "Personal Loan"]),
    "Retail_APAC": ("Retail", ["APAC"], ["Mortgage", "Credit Card", "Personal Loan"]),
    "Corporate_Global": ("Corporate", ["EMEA", "APAC", "AMER"], ["Term Loan", "Revolver", "Trade Finance"]),
    "SME_Global": ("SME", ["EMEA", "APAC", "AMER"], ["Term Loan", "Revolver", "Personal Loan"]),
}
NAMES = ["Atlas", "Borealis", "Cedar", "Delta", "Everest", "Falcon", "Granite", "Harbor", "Indigo", "Juniper",
         "Keystone", "Lumen", "Meridian", "Nimbus", "Orion", "Pinnacle", "Quartz", "Redwood", "Summit", "Titan"]


def xl_round(x, d=0):
    import math
    f = 10 ** d
    return math.floor(abs(x) * f + 0.5 + 1e-9) / f * (1 if x >= 0 else -1)


def make_rows(sheet, n, rnd, start):
    seg, regions, prods = SHEETS[sheet]
    rows = []
    for i in range(n):
        reg = rnd.choice(regions)
        cn = rnd.choice([c[0] for c in COUNTRIES if c[1] == reg])
        prod = rnd.choice(prods)
        big = seg == "Corporate"
        expo = round(rnd.choice([8e5, 2.5e6, 7e6, 1.5e7]) * rnd.uniform(0.6, 1.4) if big
                     else rnd.choice([5e3, 2e4, 9e4, 3e5, 1.2e6]) * rnd.uniform(0.7, 1.3), 0)
        rows.append({
            "LoanID": "%s%05d" % (seg[0], start + i), "Customer": "%s %s" % (rnd.choice(NAMES), rnd.choice(["Ltd", "Corp", "Group", "Holdings"])),
            "Segment": seg, "Region": reg, "Product": prod, "Rating": rnd.choice([r[0] for r in RATINGS]),
            "Country": cn, "Currency": CCY_BY_COUNTRY[cn], "Exposure": expo,
            "Collateral": round(expo * rnd.choice([0, 0, 0.3, 0.6, 0.9])),
            "Maturity": dt.datetime(rnd.choice([2027, 2028, 2030, 2031, 2034]), rnd.randint(1, 12), rnd.randint(1, 28)),
        })
    return rows


# ------------------------------------------------------------ derived columns per sheet
# (name, excel template with [Column] placeholders, python fn(row dict) -> value)
def L(table, key, col):  # python-side lookups
    return {"Ratings": rating, "Products": product, "Countries": country, "FX": {k: (k, v) for k, v in FX}}[table][key][col]


FXRATE = ("FX_Rate", "VLOOKUP([Currency],FX!$A$2:$B$8,2,FALSE)", lambda r: fx[r["Currency"]])
EXPUSD = ("Exposure_USD", "ROUND([Exposure]*[FX_Rate],2)", lambda r: xl_round(r["Exposure"] * r["FX_Rate"], 2))
RW = ("Risk_Weight", "VLOOKUP([Rating],Ratings!$A$2:$C$7,3,FALSE)", lambda r: rating[r["Rating"]][2])

DERIVED = {
    "Retail_EMEA": [
        FXRATE, EXPUSD,
        ("PD", "VLOOKUP([Rating],Ratings!$A$2:$C$7,2,FALSE)", lambda r: rating[r["Rating"]][1]),
        ("LGD", "VLOOKUP([Product],Products!$A$2:$C$7,2,FALSE)", lambda r: product[r["Product"]][1]),
        ("Expected_Loss", "ROUND([Exposure_USD]*[PD]*[LGD],2)", lambda r: xl_round(r["Exposure_USD"] * r["PD"] * r["LGD"], 2)),
    ],
    "Retail_APAC": [
        FXRATE, EXPUSD,
        ("Country_Risk", "VLOOKUP([Country],Countries!$A$2:$C$13,3,FALSE)", lambda r: country[r["Country"]][2]),
        RW,
        ("RWA", "ROUND([Exposure_USD]*[Risk_Weight],2)", lambda r: xl_round(r["Exposure_USD"] * r["Risk_Weight"], 2)),
        ("Review_Flag", 'IF([Country_Risk]="High","Review","OK")', lambda r: "Review" if r["Country_Risk"] == "High" else "OK"),
    ],
    "Corporate_Global": [
        FXRATE, EXPUSD,
        ("Haircut", "VLOOKUP([Product],Products!$A$2:$C$7,3,FALSE)", lambda r: product[r["Product"]][2]),
        ("Net_Exposure", "MAX(0,[Exposure_USD]-[Collateral]*[FX_Rate]*(1-[Haircut]))",
         lambda r: max(0, r["Exposure_USD"] - r["Collateral"] * r["FX_Rate"] * (1 - r["Haircut"]))),
        RW,
        ("RWA", "ROUND([Net_Exposure]*[Risk_Weight],2)", lambda r: xl_round(r["Net_Exposure"] * r["Risk_Weight"], 2)),
        ("Years_To_Maturity", "YEAR([Maturity])-2026", lambda r: r["Maturity"].year - 2026),
        ("Tenor_Bucket", 'IF([Years_To_Maturity]>=5,"Long",IF([Years_To_Maturity]>=2,"Medium","Short"))',
         lambda r: "Long" if r["Years_To_Maturity"] >= 5 else ("Medium" if r["Years_To_Maturity"] >= 2 else "Short")),
    ],
    "SME_Global": [
        FXRATE, EXPUSD, RW,
        ("SME_Factor", "IF([Exposure_USD]<=1000000,0.7619,1)", lambda r: 0.7619 if r["Exposure_USD"] <= 1000000 else 1),
        ("RWA", "ROUND([Exposure_USD]*[Risk_Weight]*[SME_Factor],2)",
         lambda r: xl_round(r["Exposure_USD"] * r["Risk_Weight"] * r["SME_Factor"], 2)),
        ("Capital_Required", "ROUND([RWA]*0.08,2)", lambda r: xl_round(r["RWA"] * 0.08, 2)),
        ("Label", '[Customer]&"-"&[Country]', lambda r: r["Customer"] + "-" + r["Country"]),
    ],
}


# ----------------------------------------------------------------------- writers
def _bytes(wb):
    b = io.BytesIO()
    wb.save(b)
    return b.getvalue()


def input_workbook(data):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for sh, rows in data.items():
        ws = wb.create_sheet(sh)
        ws.append(BASE)
        for r in rows:
            ws.append([r[c] for c in BASE])
    return _bytes(wb)


def mapping_workbook():
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for name, (hdr, rows) in MAP_SHEETS.items():
        ws = wb.create_sheet(name)
        ws.append(hdr)
        for r in rows:
            ws.append(list(r))
    return _bytes(wb)


def output_workbook(data):
    """Data sheets with live formulas + cached values, followed by the mapping sheets."""
    from openpyxl.utils import get_column_letter
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    cached = {}
    for sh, rows in data.items():
        ws = wb.create_sheet(sh)
        spec = DERIVED[sh]
        cols = BASE + [d[0] for d in spec]
        letter = {c: get_column_letter(i + 1) for i, c in enumerate(cols)}
        ws.append(cols)
        cached[sh] = {}
        for i, r in enumerate(rows, start=2):
            row = dict(r)
            for name, tpl, fn in spec:
                row[name] = fn(row)
            line = [r[c] for c in BASE]
            for name, tpl, fn in spec:
                f = re.sub(r"\[([A-Za-z_]+)\]", lambda m: "%s%d" % (letter[m.group(1)], i), tpl)
                line.append("=" + f)
                cached[sh]["%s%d" % (letter[name], i)] = row[name]
            ws.append(line)
    for name, (hdr, rows) in MAP_SHEETS.items():
        ws = wb.create_sheet(name)
        ws.append(hdr)
        for r in rows:
            ws.append(list(r))
    return _patch_cached(_bytes(wb), list(data), cached)


def _patch_cached(blob, sheet_names, cached):
    """openpyxl writes formulas without results; add <v> so the file looks like one saved by Excel."""
    zin = zipfile.ZipFile(io.BytesIO(blob))
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            m = re.match(r"xl/worksheets/sheet(\d+)\.xml$", item.filename)
            if m and int(m.group(1)) <= len(sheet_names):
                vals = cached[sheet_names[int(m.group(1)) - 1]]
                text = data.decode("utf-8")

                def fix(mm):
                    coord, attrs, f = mm.group(1), mm.group(2), mm.group(3)
                    v = vals[coord]
                    attrs = re.sub(r'\s+t="[^"]*"', "", attrs)
                    if isinstance(v, str):
                        return '<c r="%s"%s t="str"><f>%s</f><v>%s</v></c>' % (coord, attrs, f, escape(v))
                    return '<c r="%s"%s><f>%s</f><v>%s</v></c>' % (coord, attrs, f, repr(float(v)) if not float(v).is_integer() else int(v))
                text = re.sub(r'<c r="([A-Z]+\d+)"([^>]*)><f>(.*?)</f><v\s*/?>(?:</v>)?</c>', fix, text)
                data = text.encode("utf-8")
            zout.writestr(item, data)
    return out.getvalue()


def build(seed, sizes, start):
    rnd = random.Random(seed)
    return {sh: make_rows(sh, n, rnd, start + k * 1000) for k, (sh, n) in enumerate(sizes.items())}


def main():
    OUT.mkdir(exist_ok=True)
    train = build(11, {"Retail_EMEA": 28, "Retail_APAC": 24, "Corporate_Global": 26, "SME_Global": 22}, 1)
    new = build(97, {"Retail_EMEA": 12, "Retail_APAC": 10, "Corporate_Global": 14, "SME_Global": 9}, 50001)
    files = {"1_input_exposures.xlsx": input_workbook(train), "2_mapping_tables.xlsx": mapping_workbook(),
             "3_expected_output.xlsx": output_workbook(train), "4_new_input_exposures.xlsx": input_workbook(new),
             "5_new_expected_output.xlsx": output_workbook(new)}
    for n, b in files.items():
        (OUT / n).write_bytes(b)
        print("wrote", OUT / n, len(b) // 1024, "KB")


if __name__ == "__main__":
    main()
