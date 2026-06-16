"""Parse the Poonam-format Excel roster into structured data."""
import pandas as pd
from collections import defaultdict
from datetime import datetime, timedelta
import re

HOUSES = ['Daffodil','Christopher','Hilary','Gabriel','Parzival','Michael','Magnolia','Wake']
HOUSE_SHORT = {
    'Daffodil':'DC','Christopher':'CH','Hilary':'HH','Gabriel':'GH',
    'Parzival':'PH','Michael':'MiH','Magnolia':'MaG','Wake':'Wake'
}

LEAVE_CODES = {'AL','SL','BL','ML','ACC','LWOP','LTW','ALT','RDO',
               'Admin Day (HL)','Training','Induction','Open Shift','Closed Shift',
               'Move to AM','Move to PM','Moved from AM','Moved from PM',
               'Added Hours','Added Hours, Replacement','Changed Hours','LWOP',
               'Swapped Shift','No Replacement','DNA','RDO','Incorrect Roster',
               'Not Worked','Given Day-off','Move to DC','Move to HH','Move to GH',
               'Move to MiH','Moved From DC','Moved From HH','Moved From GH',
               'Move to MaG','Move to CH','Move to PH','Move to wake',
               'Admin Day (HL)','Admin Day'}

TIME_PATTERN = re.compile(r'^\d{1,2}:\d{2}')
HOURS_PATTERN= re.compile(r'^\d{1,2}:\d{2}:\d{2}$')
AM_PM_PATTERN= re.compile(r'(AM|PM)', re.IGNORECASE)

def parse_hours(val):
    if pd.isna(val) or val is None: return 0.0
    if isinstance(val, timedelta): return round(val.total_seconds()/3600, 2)
    s = str(val).strip()
    m = re.match(r'(\d+):(\d+)', s)
    if m: return int(m.group(1)) + int(m.group(2))/60
    try: return float(s)
    except: return 0.0

def is_shift_time(v):
    """True if value looks like a shift time (e.g. '6:00 AM - 2:00 PM')"""
    return bool(AM_PM_PATTERN.search(str(v))) or bool(HOURS_PATTERN.match(str(v).strip()))

def is_hours_val(v):
    """True if value looks like duration (8:00:00)"""
    return bool(HOURS_PATTERN.match(str(v).strip()))

def is_leave_code(v):
    v = str(v).strip()
    return (v in LEAVE_CODES or
            any(v.startswith(lc) for lc in ['AL','SL','BL','ML','ACC','LWOP','Open Shift','Move','Moved','Admin','Added','Changed','Swapped']))

def is_valid_name(v):
    """True if v looks like a person's name (not a time, number, or leave code)"""
    v = str(v).strip()
    if not v or v in ('nan','None',''): return False
    if is_shift_time(v): return False
    if is_hours_val(v): return False
    if is_leave_code(v): return False
    if re.match(r'^\d+[\.:]\d+', v): return False  # starts with number:
    if re.match(r'^\d+\.?\d*$', v): return False   # pure number
    if len(v) < 2: return False
    return True

def parse_roster(filepath):
    xl = pd.ExcelFile(filepath)

    # Staff from Source Data
    staff_map = {}
    if 'Source Data' in xl.sheet_names:
        src = pd.read_excel(xl, 'Source Data', header=0)
        for _, row in src.iterrows():
            name = str(row.get('Given Name','')).strip()
            hrs  = str(row.get('Contract Hours','')).strip()
            ctype= str(row.get('Type of Contract','')).strip()
            if name and name not in ('nan','Open Shift','Filled by Agency','None'):
                try: contract = float(hrs)
                except: contract = hrs if hrs not in ('nan','') else None
                staff_map[name] = {
                    'contract': contract,
                    'type': '' if ctype in ('nan','None') else ctype
                }

    result = {}

    for house in HOUSES:
        if house not in xl.sheet_names:
            continue
        df = pd.read_excel(xl, house, header=None)

        # Find the row containing dates (row 1 in 0-indexed)
        date_col_map = {}  # col_idx -> date_str
        date_row_idx = None

        for ri in range(min(5, len(df))):
            for ci in range(df.shape[1]):
                v = df.iloc[ri, ci]
                try:
                    if isinstance(v, str):
                        if '2026' in v or '2025' in v:
                            dt = pd.Timestamp(v)
                        else: continue
                    else:
                        dt = pd.Timestamp(v)
                    if 2020 < dt.year < 2030:
                        date_col_map[ci] = dt.strftime('%Y-%m-%d')
                        date_row_idx = ri
                except: pass

        if not date_col_map:
            continue

        # Now scan all rows after date_row for name patterns
        # Strategy: a name row has >= 1 valid person name in the date columns
        # The row immediately after is usually shift time
        # The row after that is usually hours (HH:MM:SS)
        # Optional rows after: leave code, replacement name

        house_data = defaultdict(list)  # date_str -> list of slot dicts

        # Build a list of (row_idx, {col: name}) for rows that are "name rows"
        i = date_row_idx + 1
        while i < len(df):
            row = df.iloc[i]
            names_in_row = {}
            for ci, ds in date_col_map.items():
                if ci >= df.shape[1]: continue
                v = str(row.iloc[ci]).strip() if not pd.isna(row.iloc[ci]) else ''
                if is_valid_name(v):
                    names_in_row[ci] = v

            if not names_in_row:
                i += 1
                continue

            # Check next row for shift times to confirm this is a name row
            shift_row = i + 1
            has_shifts = False
            if shift_row < len(df):
                for ci in names_in_row:
                    sv = str(df.iloc[shift_row, ci]).strip() if ci < df.shape[1] else ''
                    if is_shift_time(sv):
                        has_shifts = True; break

            if not has_shifts:
                i += 1
                continue

            # It's a name row — now extract all data
            hours_row = shift_row + 1
            # Look ahead for leave codes (within next 5 rows)
            leave_map = {}  # col -> leave_code
            repl_map  = {}  # col -> replacement_name

            for li in range(shift_row + 1, min(i + 8, len(df))):
                lrow = df.iloc[li]
                for ci in date_col_map:
                    if ci >= df.shape[1]: continue
                    lv = str(lrow.iloc[ci]).strip() if not pd.isna(lrow.iloc[ci]) else ''
                    if lv and ci not in leave_map and is_leave_code(lv) and not is_shift_time(lv):
                        leave_map[ci] = lv
                    elif lv and ci in leave_map and ci not in repl_map and is_valid_name(lv):
                        repl_map[ci] = lv

            for ci, ds in date_col_map.items():
                name = names_in_row.get(ci, '')
                if not name: continue

                shift = ''
                if shift_row < len(df) and ci < df.shape[1]:
                    sv = str(df.iloc[shift_row, ci]).strip()
                    if is_shift_time(sv) and sv not in ('nan','None'): shift = sv

                hours = 0.0
                if hours_row < len(df) and ci < df.shape[1]:
                    hours = parse_hours(df.iloc[hours_row, ci])

                leave = leave_map.get(ci, '')
                replacement = repl_map.get(ci, '')

                house_data[ds].append({
                    'name': name,
                    'shift': shift,
                    'hours': hours,
                    'leave': leave,
                    'replacement': replacement,
                    'notes': '',
                })

            i = shift_row + 1

        result[house] = dict(house_data)

    return result, staff_map


def compute_staff_hours(roster_data):
    staff = defaultdict(lambda: {'houses': defaultdict(float), 'total': 0.0, 'shifts': [], 'days': set()})
    for house, date_data in roster_data.items():
        short = HOUSE_SHORT.get(house, house)
        for date_str, slots in date_data.items():
            seen_in_day = set()
            for slot in slots:
                name = slot['name']
                if not name or name in ('Open Shift','nan','None'): continue
                key = (name, date_str, house, slot['shift'])
                if key in seen_in_day: continue
                seen_in_day.add(key)
                hrs = slot['hours']
                staff[name]['houses'][short] += hrs
                staff[name]['total'] += hrs
                staff[name]['shifts'].append({
                    'date': date_str, 'house': short,
                    'shift': slot['shift'], 'hours': hrs,
                    'leave': slot.get('leave',''),
                })
                staff[name]['days'].add(date_str)
    return dict(staff)
