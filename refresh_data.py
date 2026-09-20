#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Refresh the market-share data embedded in index.html from the KBA workbook.

index.html is a built React bundle -- the source project is not in this repo --
so the data is not rebuilt from source. It does not need to be: the bundle keeps
its entire dataset in one generated object literal,

    ge={schemaVersion:1,sourceFile:`...`,generatedAt:`...`,units:`k EUR`,
        banks:[...],dimensions:[{name,sourceRow,metricType,periods,series},...]}

which is self-contained and can be regenerated from the workbook and spliced
back in. Everything around it -- the React code, the styling, the layout -- is
left byte-for-byte untouched, so refreshing the data cannot change behaviour.

Every period is re-read on each run, not just the new month, so a figure the
workbook restates is corrected rather than left frozen at whatever the last
build captured. That matters here: December is restated after audit, and the
March 2026 file had RBKO's housing and consumer loans transposed.

The rebuild is self-validating. Dimensions are located by heading name and the
row found is checked against the sourceRow the previous build recorded; period
columns are read from the header rather than assumed; and the periods the
previous build already covered must be reproduced from the workbook before
anything is written. Wide disagreement means a column moved and aborts the run.

Usage:
    python refresh_data.py [workbook.xlsx] [--write]

Without --write it reports what would change and leaves index.html alone.
"""
import os
import re
import sys
import datetime

import openpyxl

BASE = os.path.dirname(os.path.abspath(__file__))
INDEX = os.path.join(BASE, 'index.html')
SHEET = 'Market Share EUR'

DEFAULT_WORKBOOK = (r'C:\Users\TechStore\Downloads'
                    r'\Market Share Dec 2021 - August 2026 - Trend.xlsx')

BT = chr(96)      # backtick — the bundler quotes strings with these
BS = chr(92)

ANCHOR = 'ge={schemaVersion:1'


# ----------------------------------------------------------------- bundle
def match_brace(s, i):
    """i points at '{'; return the index just past its matching '}'."""
    depth, instr, j = 0, None, i
    while j < len(s):
        ch = s[j]
        if instr:
            if ch == BS:
                j += 2
                continue
            if ch == instr:
                instr = None
        elif ch in (BT, '"', "'"):
            instr = ch
        elif ch == '{':
            depth += 1
        elif ch == '}':
            depth -= 1
            if depth == 0:
                return j + 1
        j += 1
    raise ValueError('unbalanced braces in bundle')


def read_blob(html):
    start = html.index(ANCHOR)
    i = html.index('{', start)
    end = match_brace(html, i)
    return i, end, html[i:end]


def parse_blob(blob):
    """-> (meta, [{name, sourceRow, metricType, periods, series}])."""
    meta = {}
    for key in ('sourceFile', 'generatedAt', 'units'):
        m = re.search(key + ':' + BT + '([^' + BT + ']*)' + BT, blob)
        meta[key] = m.group(1) if m else None
    meta['banks'] = re.findall(BT + '([A-Z]{3,4})' + BT,
                               blob[blob.index('banks:['):blob.index('dimensions:[')])

    dims = []
    for m in re.finditer(r'\{name:' + BT, blob):
        obj = blob[m.start():match_brace(blob, m.start())]
        head = re.match(r'\{name:' + BT + '([^' + BT + ']+)' + BT +
                        r',sourceRow:(\d+),metricType:' + BT + '([^' + BT + ']+)' + BT, obj)
        periods = re.search('periods:' + BT + '([^' + BT + ']+)' + BT, obj).group(1).split('.')
        def nums(text):
            return [None if v.strip() == 'null' else float(v) for v in text.split(',')]

        si = obj.index('series:{')
        se = match_brace(obj, si + len('series:'))
        series = {sm.group(1): nums(sm.group(2))
                  for sm in re.finditer(r'([A-Z]{3,4}):\[([^\]]*)\]', obj[si:se])}
        sector = re.search(r'sector:\[([^\]]*)\]', obj[se:])
        dims.append({'name': head.group(1), 'sourceRow': int(head.group(2)),
                     'metricType': head.group(3), 'periods': periods, 'series': series,
                     'sector': nums(sector.group(1)) if sector else None})
    return meta, dims


# ------------------------------------------------------------- workbook
def clean(v):
    return re.sub(r'[\s\u200b\u00a0]+', ' ', str(v)).strip()


def num(v):
    if v is None or isinstance(v, bool):
        return None
    return float(v) if isinstance(v, (int, float)) else None


def parse_period(label):
    """'8-2026' -> '2026-08'. Datetimes are accepted too, in case the workbook
    reverts to real dates."""
    if hasattr(label, 'year') and hasattr(label, 'month'):
        return '%04d-%02d' % (label.year, label.month)
    if not isinstance(label, str):
        return None
    m = re.match(r'^\s*(\d{1,2})-(\d{4})\s*$', label)
    if not m:
        return None
    month, year = int(m.group(1)), int(m.group(2))
    return '%04d-%02d' % (year, month) if 1 <= month <= 12 else None


SECTOR_KEY = '__sector__'


def read_workbook(path, banks):
    """-> (periods, {heading: (excel_row, {bank: [values]})}).

    The 'Banking sector' row that closes each block is carried under
    SECTOR_KEY: the bundle stores it per dimension as `sector`, and the app
    divides by it to get market share, so it is not optional.
    """
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    rows = list(wb[SHEET].iter_rows(values_only=True))
    wb.close()

    period_cols, periods = None, None
    blocks, heading, row_no = {}, None, None

    for ri, row in enumerate(rows):
        c1 = clean(row[1]) if len(row) > 1 and row[1] is not None else ''
        c2 = row[2] if len(row) > 2 else None

        if c1 and parse_period(c2):
            if period_cols is None:
                period_cols, periods = [], []
                for ci in range(2, len(row)):
                    p = parse_period(row[ci])
                    if p is None:
                        break          # QoQ / MoM / YoY columns start here
                    period_cols.append(ci)
                    periods.append(p)
            heading, row_no = c1, ri + 1            # 1-based, as sourceRow is
            blocks[heading] = (row_no, {})
            continue

        if not heading:
            continue
        if c1 in banks:
            key = c1
        elif c1.lower().startswith('banking sector'):
            key = SECTOR_KEY
        else:
            continue
        blocks[heading][1][key] = [
            (num(row[ci]) if ci < len(row) else None) for ci in period_cols]

    return periods, blocks


# ------------------------------------------------------------ serialize
def jsnum(v):
    """Shortest round-trip literal, written the way the bundler writes it."""
    if v is None:
        return 'null'
    s = repr(float(v))
    if 'e' in s or 'E' in s:
        raise ValueError('exponent form would change the literal: %r' % v)
    if s.endswith('.0'):
        s = s[:-2]
    if s.startswith('0.'):
        s = s[1:]
    elif s.startswith('-0.'):
        s = '-' + s[2:]
    return s


def build_blob(meta, dims, source_file):
    q = lambda s: BT + s + BT
    generated = datetime.datetime.now(datetime.timezone.utc).strftime('%Y-%m-%dT%H:%M:%S+00:00')
    out = ['{schemaVersion:1',
           'sourceFile:' + q(source_file),
           'generatedAt:' + q(generated),
           'units:' + q(meta['units']),
           'banks:[' + ','.join(q(b) for b in meta['banks']) + ']']
    parts = []
    for d in dims:
        series = ','.join(
            '%s:[%s]' % (b, ','.join(jsnum(v) for v in d['series'][b]))
            for b in meta['banks'])
        sector = ''
        if d.get('sector') is not None:
            sector = ',sector:[' + ','.join(jsnum(v) for v in d['sector']) + ']'
        parts.append('{name:%s,sourceRow:%d,metricType:%s,periods:%s.split(%s),series:{%s}%s}'
                     % (q(d['name']), d['sourceRow'], q(d['metricType']),
                        q('.'.join(d['periods'])), q('.'), series, sector))
    out.append('dimensions:[' + ','.join(parts) + ']')
    return ','.join(out) + '}'


# ------------------------------------------------------------- validate
def main():
    args = [a for a in sys.argv[1:] if a != '--write']
    write = '--write' in sys.argv
    workbook = args[0] if args else DEFAULT_WORKBOOK

    html = open(INDEX, encoding='utf-8').read()
    i, end, old_blob = read_blob(html)
    meta, old_dims = parse_blob(old_blob)
    old_by_name = {d['name']: d for d in old_dims}

    print('Current bundle')
    print('  source   %s' % meta['sourceFile'])
    print('  built    %s' % meta['generatedAt'])
    print('  %d dimensions, %d periods through %s'
          % (len(old_dims), len(old_dims[0]['periods']), old_dims[0]['periods'][-1]))

    print('\nWorkbook  %s' % os.path.basename(workbook))
    periods, blocks = read_workbook(workbook, meta['banks'])
    print('  %d periods, %s -> %s' % (len(periods), periods[0], periods[-1]))

    # ---- rebuild, in the bundle's dimension order
    new_dims, problems, moved = [], [], []
    for d in old_dims:
        if d['name'] not in blocks:
            problems.append('%s: heading not found in the workbook' % d['name'])
            continue
        row_no, series = blocks[d['name']]
        if row_no != d['sourceRow']:
            moved.append('%s: row %d -> %d' % (d['name'], d['sourceRow'], row_no))
        missing = [b for b in meta['banks'] if b not in series]
        if d.get('sector') is not None and SECTOR_KEY not in series:
            missing.append('Banking sector')
        if missing:
            problems.append('%s: no row for %s' % (d['name'], ', '.join(missing)))
            continue
        new_dims.append({'name': d['name'], 'sourceRow': row_no,
                         'metricType': d['metricType'], 'periods': periods,
                         'series': {b: series[b] for b in meta['banks']},
                         'sector': series.get(SECTOR_KEY) if d.get('sector') is not None else None})

    print('\nValidation')
    if problems:
        for p in problems:
            print('  FAIL  %s' % p)
        print('\nNot written.')
        return 1
    n_sector = sum(1 for d in new_dims if d.get('sector') is not None)
    print('  OK    all %d dimensions found, all %d bank series present, '
          '%d sector totals' % (len(new_dims), len(meta['banks']), n_sector))
    for m in moved:
        print('  NOTE  %s' % m)

    # every series as long as the period axis
    bad = ['%s/%s' % (d['name'], b) for d in new_dims for b in meta['banks']
           if len(d['series'][b]) != len(periods)]
    bad += ['%s/sector' % d['name'] for d in new_dims
            if d.get('sector') is not None and len(d['sector']) != len(periods)]
    if bad:
        print('  FAIL  wrong length: %s' % ', '.join(bad[:6]))
        print('\nNot written.')
        return 1
    print('  OK    every series aligned to %d periods' % len(periods))

    # The app divides by `sector` to get share, so the ten banks must add up to
    # it -- on value dimensions; a ratio dimension's sector row is its own ratio.
    off = []
    for d in new_dims:
        if d.get('sector') is None or d['metricType'] != 'value':
            continue
        for k, p in enumerate(periods):
            sec = d['sector'][k]
            if not sec:
                continue
            tot = sum(d['series'][b][k] or 0 for b in meta['banks'])
            if abs(tot - sec) > max(abs(sec) * 0.005, 1.0):
                off.append('%s@%s (%.0f vs %.0f)' % (d['name'], p, tot, sec))
    if off:
        print('  FAIL  banks do not sum to the sector total: %s' % '; '.join(off[:6]))
        print('\nNot written.')
        return 1
    print('  OK    the ten banks sum to the sector total in every value dimension')

    # reconcile the periods the previous build already covered
    hits = checked = 0
    restated = []
    for d in new_dims:
        old = old_by_name[d['name']]
        shared = [p for p in old['periods'] if p in periods]
        tracks = [(b, old['series'][b], d['series'][b]) for b in meta['banks']]
        if old.get('sector') is not None and d.get('sector') is not None:
            tracks.append(('sector', old['sector'], d['sector']))
        for b, old_arr, new_arr in tracks:
            for p in shared:
                o = old_arr[old['periods'].index(p)]
                nv = new_arr[periods.index(p)]
                if o is None or nv is None:
                    if o != nv:
                        restated.append((p, d['name'], b, o, nv))
                    continue
                checked += 1
                tol = max(abs(o) * 1e-6, 1e-6 if d['metricType'] == 'ratio' else 0.5)
                if abs(o - nv) <= tol:
                    hits += 1
                else:
                    restated.append((p, d['name'], b, o, nv))
    rate = hits / float(checked) if checked else 0.0
    line = ('reproduces %d/%d (%.2f%%) of the cells the previous build published'
            % (hits, checked, rate * 100))
    if rate < 0.97:
        print('  FAIL  %s -- too widespread to be restatement; a column moved' % line)
        print('\nNot written.')
        return 1
    print('  OK    %s' % line)

    new_periods = [p for p in periods if p not in old_dims[0]['periods']]
    dropped = [p for p in old_dims[0]['periods'] if p not in periods]
    print('  OK    adds %d period(s): %s' % (len(new_periods), ', '.join(new_periods) or 'none'))
    if dropped:
        print('  NOTE  the workbook no longer carries: %s' % ', '.join(dropped))

    if restated:
        by_period = {}
        for p, _, _, _, _ in restated:
            by_period[p] = by_period.get(p, 0) + 1
        print('\nRestated by the workbook (%d cells):' % len(restated))
        for p in sorted(by_period):
            print('  %s  %d cell(s)' % (p, by_period[p]))
        big = [x for x in restated
               if x[3] and x[4] is not None and abs((x[4] - x[3]) / x[3]) > 0.20]
        for p, name, b, o, nv in sorted(big)[:12]:
            print('    %s  %-28s %-5s %13.1f -> %13.1f  %+.1f%%'
                  % (p, name, b, o, nv, (nv - o) / o * 100))

    # ---- splice
    blob = build_blob(meta, new_dims, os.path.basename(workbook))
    new_html = html[:i] + blob + html[end:]

    # the rebuilt bundle must differ only in the blob
    assert new_html[:i] == html[:i] and new_html[len(new_html) - (len(html) - end):] == html[end:]
    j, k, check = read_blob(new_html)
    assert check == blob, 'spliced blob does not read back'
    m2, d2 = parse_blob(check)
    assert len(d2) == len(new_dims) and d2[0]['periods'] == periods, 're-parse mismatch'
    print('\n  OK    rebuilt blob re-parses: %d dimensions, %d periods through %s'
          % (len(d2), len(d2[0]['periods']), d2[0]['periods'][-1]))
    print('  OK    bundle changes only inside the data literal '
          '(%d bytes before, %d after, blob %d -> %d)'
          % (i, len(html) - end, len(old_blob), len(blob)))

    if not write:
        print('\nDry run — pass --write to update index.html.')
        return 0

    open(INDEX, 'w', encoding='utf-8', newline='').write(new_html)
    print('\nWrote %s  (%.0f KB)' % (INDEX, len(new_html) / 1024.0))

    a = new_dims[0]
    total = a['sector'][-1] if a.get('sector') else sum(a['series'][b][-1] or 0
                                                        for b in meta['banks'])
    print('\nAssets at %s (k EUR / share):' % periods[-1])
    for b in sorted(meta['banks'], key=lambda b: -(a['series'][b][-1] or 0)):
        v = a['series'][b][-1] or 0
        print('  %-5s %12s   %5.2f%%' % (b, '{:,.0f}'.format(v), v / total * 100))
    print('  %-5s %12s' % ('Total', '{:,.0f}'.format(total)))
    return 0


if __name__ == '__main__':
    sys.exit(main())
