import streamlit as st
import pandas as pd
import re
import uuid
from datetime import datetime, date, time
from dateutil import parser as dtparser
import pytz
from openpyxl import load_workbook

# ---------------- Utilities ----------------

def normalize_group_label(x):
    if x is None:
        return None
    try:
        if pd.isna(x):
            return None
    except Exception:
        pass
    s = str(x).strip()
    if not s:
        return None
    m = re.search(r'G\s*\.?\s*(\d+)', s, re.I)
    if m:
        return f'G {m.group(1)}'
    m2 = re.search(r'^(?:groupe)?\s*(\d+)$', s, re.I)
    if m2:
        return f'G {m2.group(1)}'
    return s


def is_time_like(x):
    if x is None:
        return False
    if isinstance(x, (pd.Timestamp, datetime, time)):
        return True
    s = str(x).strip()
    if not s:
        return False
    if re.match(r'^\d{1,2}[:hH]\d{2}(\s*[AaPp][Mm]\.?)*$', s):
        return True
    return False


def to_time(x):
    if x is None:
        return None
    if isinstance(x, time):
        return x
    if isinstance(x, pd.Timestamp):
        return x.to_pydatetime().time()
    if isinstance(x, datetime):
        return x.time()
    s = str(x).strip()
    if not s:
        return None
    s2 = s.replace('h', ':').replace('H', ':')
    try:
        dt = dtparser.parse(s2, dayfirst=True)
        return dt.time()
    except Exception:
        return None


def to_date(x):
    if x is None:
        return None
    if isinstance(x, pd.Timestamp):
        return x.to_pydatetime().date()
    if isinstance(x, datetime):
        return x.date()
    if isinstance(x, date):
        return x
    s = str(x).strip()
    if not s:
        return None
    try:
        dt = dtparser.parse(s, dayfirst=True, fuzzy=True)
        return dt.date()
    except Exception:
        return None

# ---------------- Helpers Fusion ----------------

def get_merged_map(xls_path, sheet_name):
    """Retourne un dict {(row,col): (row1,col1,row2,col2)} si la cellule est fusionnée"""
    wb = load_workbook(xls_path, data_only=True)
    ws = wb[sheet_name]
    merged_map = {}
    for merged in ws.merged_cells.ranges:
        r1, r2 = merged.min_row, merged.max_row
        c1, c2 = merged.min_col, merged.max_col
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                merged_map[(r - 1, c - 1)] = (r1 - 1, c1 - 1, r2 - 1, c2 - 1)
    return merged_map

# ---------------- Parsing ----------------

def find_week_rows(df):
    return [i for i in range(len(df)) if isinstance(df.iat[i, 0], str) and re.match(r'^\s*S\s*\d+', df.iat[i, 0].strip(), re.I)]


def find_slot_rows(df):
    return [i for i in range(len(df)) if isinstance(df.iat[i, 0], str) and re.match(r'^\s*H\s*\d+', df.iat[i, 0].strip(), re.I)]


def parse_sheet_to_events(xls_path, sheet_name):
    df = pd.read_excel(xls_path, sheet_name=sheet_name, header=None)
    nrows, ncols = df.shape
    merged_map = get_merged_map(xls_path, sheet_name)

    s_rows = find_week_rows(df)
    h_rows = find_slot_rows(df)

    raw_events = []

    for r in h_rows:
        p_candidates = [s for s in s_rows if s <= r]
        if not p_candidates:
            continue
        p = max(p_candidates)
        date_row = p + 1
        group_row = p + 2

        date_cols = [c for c in range(ncols) if date_row < nrows and to_date(df.iat[date_row, c]) is not None]

        for c in date_cols:
            for col in (c, c + 1):
                if col >= ncols:
                    continue
                summary = df.iat[r, col] if col < ncols else None
                if pd.isna(summary) or summary is None:
                    continue
                summary_str = str(summary).strip()
                if not summary_str:
                    continue

                # --- Teachers ---
                teachers = []
                if (r + 2) < nrows:
                    for off in range(2, 6):
                        idx = r + off
                        if idx >= nrows:
                            break
                        t = df.iat[idx, col]
                        if t is None or pd.isna(t):
                            continue
                        s = str(t).strip()
                        if not s:
                            continue
                        if not is_time_like(s) and to_date(s) is None:
                            teachers.append(s)
                teachers = list(dict.fromkeys(teachers))

                # --- Stop index ---
                stop_idx = None
                for off in range(1, 12):
                    idx = r + off
                    if idx >= nrows:
                        break
                    if is_time_like(df.iat[idx, col]):
                        stop_idx = idx
                        break
                if stop_idx is None:
                    stop_idx = min(r + 7, nrows)

                # --- Description ---
                desc_parts = []
                for idx in range(r + 1, stop_idx):
                    if idx >= nrows:
                        break
                    cell = df.iat[idx, col]
                    if pd.isna(cell) or cell is None:
                        continue
                    s = str(cell).strip()
                    if not s:
                        continue
                    if to_date(cell) is not None:
                        continue
                    if s in teachers or s == summary_str:
                        continue
                    desc_parts.append(s)
                desc_text = " | ".join(dict.fromkeys(desc_parts))

                # --- Heures ---
                start_val, end_val = None, None
                for off in range(1, 13):
                    idx = r + off
                    if idx >= nrows:
                        break
                    v = df.iat[idx, col]
                    if is_time_like(v):
                        if start_val is None:
                            start_val = v
                        elif end_val is None and v != start_val:
                            end_val = v
                            break
                if start_val is None or end_val is None:
                    continue
                start_t, end_t = to_time(start_val), to_time(end_val)
                if start_t is None or end_t is None:
                    continue

                # --- Date ---
                d = to_date(df.iat[date_row, c])
                if d is None:
                    continue
                dtstart, dtend = datetime.combine(d, start_t), datetime.combine(d, end_t)

                # --- Groupes ---
                gl = normalize_group_label(df.iat[group_row, col] if group_row < nrows else None)
                gl_next = normalize_group_label(df.iat[group_row, col + 1] if (col + 1) < ncols else None)
                is_left_col = (col == c)
                right_summary = df.iat[r, col + 1] if (col + 1) < ncols else None

                groups = set()
                if is_left_col:
                    # Vérif fusion réelle dans Excel
                    merged = merged_map.get((r, col))
                    if merged and (r, col + 1) in merged_map:
                        # fusion détectée → G1+G2
                        if gl:
                            groups.add(gl)
                        if gl_next:
                            groups.add(gl_next)
                    else:
                        # sinon groupe normal
                        if gl:
                            groups.add(gl)
                else:
                    if gl:
                        groups.add(gl)

                raw_events.append({
                    'summary': summary_str,
                    'teachers': set(teachers),
                    'descriptions': set([desc_text]) if desc_text else set(),
                    'start': dtstart,
                    'end': dtend,
                    'groups': groups
                })

    # --- Fusion des événements ---
    merged = {}
    for e in raw_events:
        key = (e['summary'], e['start'], e['end'])
        if key not in merged:
            merged[key] = {
                'summary': e['summary'],
                'teachers': set(),
                'descriptions': set(),
                'start': e['start'],
                'end': e['end'],
                'groups': set()
            }
        merged[key]['teachers'].update(e.get('teachers', set()))
        merged[key]['descriptions'].update(e.get('descriptions', set()))
        merged[key]['groups'].update(e.get('groups', set()))

    return [{
        'summary': v['summary'],
        'teachers': sorted(list(v['teachers'])),
        'description': " | ".join(sorted(list(v['descriptions']))) if v['descriptions'] else "",
        'start': v['start'],
        'end': v['end'],
        'groups': sorted(list(v['groups']))
    } for v in merged.values()]

# ---------------- ICS writer ----------------

def escape_ical_text(s: str) -> str:
    if s is None:
        return ""
    s = str(s)
    s = s.replace('\\', '\\\\').replace('\n', '\\n').replace(',', '\\,').replace(';', '\\;')
    return s


def events_to_ics(events, tzname='Europe/Paris'):
    tz = pytz.timezone(tzname)
    header = [
        'BEGIN:VCALENDAR',
        'VERSION:2.0',
        'PRODID:-//EDT Export//FR',
        'CALSCALE:GREGORIAN',
    ]
    body = []
    for ev in events:
        uid = str(uuid.uuid4())
        dtstart = tz.localize(ev['start']).strftime('%Y%m%dT%H%M%S')
        dtend = tz.localize(ev['end']).strftime('%Y%m%dT%H%M%S')
        summary = escape_ical_text(ev['summary'])

        desc_lines = []
        if ev.get('description'):
            desc_lines.append(ev['description'])
        if ev.get('teachers'):
            desc_lines.append('Enseignant(s): ' + ' / '.join(ev['teachers']))
        groups = ev.get('groups', [])
        if groups:
            if len(groups) == 1:
                desc_lines.append('Groupe: ' + groups[0])
            else:
                desc_lines.append('Groupes: ' + ' et '.join(groups))

        description = escape_ical_text('\n'.join(desc_lines))

        body.extend([
            'BEGIN:VEVENT',
            f'UID:{uid}',
            f'DTSTAMP:{datetime.utcnow().strftime("%Y%m%dT%H%M%SZ")}',
            f'DTSTART;TZID={tzname}:{dtstart}',
            f'DTEND;TZID={tzname}:{dtend}',
            f'SUMMARY:{summary}',
            f'DESCRIPTION:{description}',
            'END:VEVENT'
        ])

    footer = ['END:VCALENDAR']
    return '\n'.join(header + body + footer)

# ---------------- Streamlit UI ----------------

st.set_page_config(page_title='Excel → ICS (EDT)', layout='centered')
st.title('Convertisseur Emplois du Temps (Excel → .ics)')

uploaded = st.file_uploader('Choisir le fichier Excel (.xlsx)', type=['xlsx'])
if uploaded is None:
    st.info('Upload un fichier .xlsx pour commencer.')
    st.stop()

try:
    xls = pd.ExcelFile(uploaded)
    sheets = xls.sheet_names
except Exception as e:
    st.error('Impossible de lire le fichier Excel: ' + str(e))
    st.stop()

st.write('Feuilles trouvées :', sheets)

for sheet in ['EDT P1', 'EDT P2']:
    if sheet in sheets:
        st.header(sheet)
        events = parse_sheet_to_events(uploaded, sheet)
        if not events:
            st.warning(f'Aucun événement détecté dans {sheet} (vérifier la structure).')
            continue

        ics = events_to_ics(events, tzname='Europe/Paris')
        st.write(f'{len(events)} événements extraits pour {sheet}')
        st.download_button(label=f'Télécharger {sheet}.ics', data=ics, file_name=f'{sheet}.ics', mime='text/calendar')
