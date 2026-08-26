#!/usr/bin/env python3
"""Onafhankelijke referentie-implementatie van het rekenmodel
"Expatregeling herrekenen" (tabblad '30% regeling'), voor controle/audit
van expatregeling-tool.html.

Gebruik:
  python3 berekening_referentie.py <expatregeling.csv> <historisch_overzicht.csv>
      [--jaar 2025] [--toetsloon 46660] [--maxpct 30]
      [--lc-svw 9970] [--lc-netto 5990] [--geen-grenswaarde]
"""
import argparse, csv, math, sys
from datetime import date, timedelta


def networkdays(a, b):
    """Werkdagen ma-vr, begin en eind inclusief (zonder feestdagen, zoals in het Excel-model)."""
    if b < a:
        return 0
    n, d = 0, a
    while d <= b:
        if d.weekday() < 5:
            n += 1
        d += timedelta(days=1)
    return n


def eom(d):
    return date(d.year + (d.month == 12), d.month % 12 + 1, 1) - timedelta(days=1)


def bom(d):
    return date(d.year, d.month, 1)


def datedif_m(a, b):
    """DATEDIF(a, b, "m"): volledige maanden."""
    m = (b.year - a.year) * 12 + (b.month - a.month)
    if b.day < a.day:
        m -= 1
    return m


def excel_round(x, n=0):
    """Excel ROUND: half weg van nul."""
    f = 10 ** n
    return math.copysign(math.floor(abs(x) * f + 0.5), x) / f


def parse_date(s):
    s = (s or '').strip()
    if not s:
        return None
    d, m, y = s.replace('/', '-').split('-')[:3]
    return date(int(y), int(m), int(d))


def parse_num(s):
    s = (s or '').strip().replace('€', '').replace(' ', '')
    if not s:
        return None
    if ',' in s and '.' in s:
        if s.rindex(',') > s.rindex('.'):
            s = s.replace('.', '').replace(',', '.')
        else:
            s = s.replace(',', '')
    elif ',' in s:
        s = s.replace('.', '').replace(',', '.')
    try:
        return float(s)
    except ValueError:
        return None


def bereken(b5, b6, b3, b14, b15, max_pct):
    same = (b5.year, b5.month) == (b6.year, b6.month)
    if same:
        b7 = networkdays(b5, b6)
    elif networkdays(b5, eom(b5)) == networkdays(bom(b5), eom(b5)):
        b7 = 0  # eerste maand telt als volledige maand
    else:
        b7 = networkdays(b5, eom(b5))
    if same:
        b8 = 0
    elif b5.day == 1 or b6 == eom(b6):
        b8 = datedif_m(b5, b6 + timedelta(days=1))
    else:
        b8 = datedif_m(b5, b6 + timedelta(days=1)) - 1
    b9 = 0 if (same or b6 == eom(b6)) else networkdays(bom(b6), b6)
    b12 = (b7 / networkdays(bom(b5), eom(b5)) * (b3 / 12)) \
        + (b8 / 12 * b3) \
        + (b9 / networkdays(bom(b6), eom(b6))) * (b3 / 12)
    b16 = b14 + b15
    b18 = (b15 / b16) * 100 if b16 else None
    b20 = excel_round(b16 - b12, 0)
    b21 = b20 - b15
    b22 = None
    if b16:
        ratio = b20 / b16
        b22 = max_pct if ratio > max_pct else excel_round(ratio, 5)
    return dict(b7=b7, b8=b8, b9=b9, b12=b12, b16=b16, b18=b18, b20=b20, b21=b21, b22=b22)


def read_csv(path):
    with open(path, encoding='utf-8-sig', newline='') as f:
        sample = f.readline()
        delim = ';' if sample.count(';') >= sample.count(',') else ','
        f.seek(0)
        return [r for r in csv.reader(f, delimiter=delim) if any(c.strip() for c in r)]


def col(header, name, fallback_idx, label):
    lower = [h.strip().lower() for h in header]
    if name.lower() in lower:
        return lower.index(name.lower())
    print(f'LET OP: kolomnaam "{name}" niet gevonden voor {label}; vaste positie gebruikt.', file=sys.stderr)
    return fallback_idx


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('expat')
    ap.add_argument('historie')
    ap.add_argument('--jaar', type=int, default=date.today().year)
    ap.add_argument('--toetsloon', type=float, default=46660)
    ap.add_argument('--maxpct', type=float, default=30)
    ap.add_argument('--lc-svw', default='9970')
    ap.add_argument('--lc-netto', default='5990')
    ap.add_argument('--geen-grenswaarde', action='store_true',
                    help='Grenswaarde-kolom negeren; altijd --toetsloon gebruiken')
    a = ap.parse_args()

    expat, hist = read_csv(a.expat), read_csv(a.historie)
    eh, hh = expat[0], hist[0]
    e = dict(pers=col(eh, 'Persnr', 9, 'persnr expat'), naam=col(eh, 'LijstNaamCompleet', 10, 'naam'),
             per=col(eh, 'Periode', 3, 'periode'), run=col(eh, 'Runnr', 2, 'runnr'),
             n=col(eh, 'Datum uit Dienst', 13, 'kolom N'), u=col(eh, 'Regeling Vanaf', 20, 'kolom U'),
             v=col(eh, 'Regeling Tm', 21, 'kolom V'), gw=col(eh, 'Grenswaarde', 29, 'kolom AD'))
    h = dict(pers=col(hh, 'Persnr', 4, 'persnr historie'), lc=col(hh, 'MasterLooncode', 11, 'looncode'),
             cum=col(hh, 'cumulatief', 27, 'kolom AB'))

    cum = {}
    for r in hist[1:]:
        key = (r[h['pers']].strip(), r[h['lc']].strip())
        if key[1] in (a.lc_svw, a.lc_netto):
            v = parse_num(r[h['cum']])
            if v is not None:
                cum[key] = cum.get(key, 0) + v

    by = {}
    for r in expat[1:]:
        p = r[e['pers']].strip()
        if p:
            by.setdefault(p, []).append(r)

    max_pct = a.maxpct / 100
    w = ('Persnr;Naam;Status;B3;B5;B6;B7;B8;B9;B12;B14;B15;B16;B18;B20;B21;B22_pct')
    print(w)
    for p in sorted(by, key=lambda x: (int(x) if x.isdigit() else 0, x)):
        rows = by[p]
        last = max(rows, key=lambda r: (int(r[e['per']] or 0), int(r[e['run']] or 0)))
        u, v, n = (parse_date(last[i]) for i in (e['u'], e['v'], e['n']))
        b5 = max(d for d in (u, date(a.jaar, 1, 1)) if d)
        b6 = min(d for d in (v, n, date(a.jaar, 12, 31)) if d)
        gw = None if a.geen_grenswaarde else parse_num(last[e['gw']])
        b3 = gw if gw else a.toetsloon
        b14, b15 = cum.get((p, a.lc_svw)), cum.get((p, a.lc_netto))
        naam = last[e['naam']].strip()
        if b14 is None or b15 is None:
            print(f'{p};{naam};Handmatig controleren (looncodes ontbreken);;;;;;;;;;;;;;')
            continue
        if b6 < b5:
            print(f'{p};{naam};Niet actief in jaar;;;;;;;;;;;;;;')
            continue
        c = bereken(b5, b6, b3, b14, b15, max_pct)
        status = ('Volledig' if c['b22'] is not None and c['b22'] >= max_pct
                  else 'Geen ruimte' if c['b22'] is not None and c['b22'] <= 0
                  else 'Aanpassen')
        f = lambda x, d=2: ('' if x is None else f'{x:.{d}f}'.replace('.', ','))
        print(';'.join([p, naam, status, f(b3), b5.strftime('%d-%m-%Y'), b6.strftime('%d-%m-%Y'),
                        str(c['b7']), str(c['b8']), str(c['b9']), f(c['b12']), f(b14), f(b15),
                        f(c['b16']), f(c['b18'], 4), f(c['b20'], 0), f(c['b21']),
                        f(c['b22'] * 100 if c['b22'] is not None else None, 3)]))


if __name__ == '__main__':
    main()
