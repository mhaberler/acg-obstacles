import pandas as pd
import json
from geojson import FeatureCollection, Feature, Point, LineString, Polygon, MultiPoint
import re

fn = 'LO_OBS_DS_AREA1_20230127.xlsx'
sheet = 'Alle - All'

hi = 0  # meters; use 1 for ft
li = 1  # english; use 0 for german

# cols = [
#     'Region', 'District', 'Location', 'Type', 'Geometry', 'Coordinates',
#     'Coordinates (decimal degrees)', 'Vertical reference system',
#     'ELEV\n(M / FT)', 'MAX HGT AGL\n(M / FT)', 'Day marking', 'Lighted',
#     'Data quality requirements met', 'Identifier'
# ]


def is_yes(field):
    return "yes" in field


def is_number(s):
    try:
        float(s)
        return True
    except ValueError:
        return False


def splitloc(field):
    # split 'LO_ODS_001304 - Fiderepasshütte'
    # into ['LO_ODS_001304','Fiderepasshütte']
    s = re.split('[\-]', field)
    return [x.strip() for x in s]


def splitslash(field):
    # split on slash.
    # remove leading + trailing ws
    s = field.split('/')
    return [x.strip() for x in s]


def numlist(field):
    # remove fluff from numeric lists
    s = re.split('([\/\*\-]|\s+)', field)
    return [float(x) for x in s if is_number(x)]


def coord2gj(row, ispolygon=False):
    c = numlist(row['Coordinates (decimal degrees)'])
    ele = numlist(row['ELEV\n(M / FT)'])
    l = []
    for i in range(0, len(c) // 2, 1):
        if len(ele) < len(c):
            l.append([c[i * 2 + 1], c[i * 2], ele[hi]])
        else:
            l.append([c[i * 2 + 1], c[i * 2], ele[i * 2 + hi]])
    if ispolygon:
        l.append(l[0])
    return l


def props(row, ispoint=False):
    p = {}
    s = splitloc(row['Location'])
    p["id"] = s[0]
    p["location"] = s[1]
    t = splitslash(row['Type'])
    p["type"] = t[li]
    if ispoint:
        p["daymark"] = is_yes(row['Day marking'])
        p["lighted"] = is_yes(row['Lighted'])
        aglist = numlist(row['MAX HGT AGL\n(M / FT)'])
        if aglist:
            p["agl"] = float(aglist[hi])
    return p


def decode_line(row):
    c = coord2gj(row)
    return Feature(geometry=LineString(c), properties=props(row))


def decode_point(row):
    c = coord2gj(row)
    return Feature(geometry=Point(c[0]), properties=props(row, True))


def decode_pointgroup(row):
    c = coord2gj(row)
    return Feature(geometry=MultiPoint(c), properties=props(row))


def decode_linegroup(row):
    c = coord2gj(row)
    return Feature(geometry=LineString(c), properties=props(row))


def decode_surface(row):
    c = coord2gj(row, ispolygon=True)
    return Feature(geometry=Polygon([c]), properties=props(row))


decode = {
    'Curve / Linie': decode_line,
    'Point / Punkt': decode_point,
    'Point (grouped) / Punkt (gruppiert)': decode_pointgroup,
    'Curve (grouped) / Linie (gruppiert)': decode_linegroup,
    'Surface / Fläche': decode_surface,
}

df = pd.read_excel(fn, sheet_name=sheet, header=2)

fc = FeatureCollection([])

for i in range(len(df)):
    row = df.iloc[i]
    typ = row['Geometry']
    o = decode[typ](row)
    if o:
        fc.features.append(o)

errs = fc.errors()
if len(errs):
    raise Exception(fc.errors())

print(json.dumps(fc, indent=4))
