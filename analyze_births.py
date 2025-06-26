import zipfile
import xml.etree.ElementTree as ET
from collections import defaultdict
import csv

AGE_FILE = 'nac_viv_annio__g_edad_limpio.xlsx'
FACILITY_FILE = 'loc_ocurr_part_annio__l_ocu_limpio.csv'


def read_age_data():
    zf = zipfile.ZipFile(AGE_FILE)
    strings = [si.find('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t').text
               for si in ET.fromstring(zf.read('xl/sharedStrings.xml'))]
    ns = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
    rows = []
    for row in ET.fromstring(zf.read('xl/worksheets/sheet1.xml')).iter(ns + 'row'):
        r = [
            (strings[int(c.find(ns + 'v').text)] if c.get('t') == 's' else c.find(ns + 'v').text)
            if c.find(ns + 'v') is not None else ''
            for c in row.findall(ns + 'c')
        ]
        rows.append(r)
    headers = rows[0]
    data = []
    for r in rows[1:]:
        if not r or r[0] in ('', '-'):
            continue
        try:
            year = int(r[0])
        except ValueError:
            continue
        births = r[1]
        births = int(births) if births != '-' else None
        data.append({'anio': year, 'cantidad_nacidos_vivos': births, 'grupo_edad': r[2]})
    return data


def read_facility_data():
    rows = []
    with open(FACILITY_FILE) as f:
        for _ in range(3):
            next(f)
        for line in f:
            line = line.strip()
            if not line or line.startswith('Nota'):
                break
            rows.append(line.split(';'))
    header = ['anio', 'total', 'publico', 'privado', 'vivienda', 'otro', 'extra1', 'extra2']
    data = []
    for r in rows:
        d = dict(zip(header, r))
        d['anio'] = int(d['anio'])
        for key in ['publico', 'privado', 'vivienda', 'otro', 'total']:
            val = d[key]
            if val in ('-', ''):
                d[key] = None
            else:
                d[key] = float(val.replace('.', '').replace(',', '.'))
        data.append(d)
    return data


def linear_slope(years, values):
    N = len(years)
    mean_x = sum(years) / N
    mean_y = sum(values) / N
    num = sum((x - mean_x) * (y - mean_y) for x, y in zip(years, values))
    den = sum((x - mean_x) ** 2 for x in years)
    return num / den if den else 0.0


if __name__ == '__main__':
    age_data = read_age_data()
    facility_data = read_facility_data()

    # Trend for births from mothers aged 15-19
    age_year = defaultdict(dict)
    for d in age_data:
        if d['cantidad_nacidos_vivos'] is not None:
            age_year[d['grupo_edad']][d['anio']] = d['cantidad_nacidos_vivos']
    young_years = sorted(age_year['15 - 19'])
    young_values = [age_year['15 - 19'][y] for y in young_years]
    slope_young = linear_slope(young_years, young_values)
    trend_young = 'aumenta' if slope_young > 0 else 'disminuye'

    # Trend for births in public hospitals
    public_years = [d['anio'] for d in facility_data if d['publico'] is not None]
    public_values = [d['publico'] for d in facility_data if d['publico'] is not None]
    slope_public = linear_slope(public_years, public_values)
    trend_public = 'aumenta' if slope_public > 0 else 'disminuye'

    print(f"Tendencia nacimientos con madres de 15-19: {trend_young} (pendiente {slope_young:.2f})")
    print(f"Tendencia nacimientos en hospitales públicos: {trend_public} (pendiente {slope_public:.2f})")
