import streamlit as st
import pandas as pd
import re
import uuid
from datetime import datetime, date, time
from dateutil import parser as dtparser
import pytz

# ---------------- Utilities ----------------

def normalize_group_label(x):
    """Normalize different notations to 'G N' or return None."""
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
    # fallback: return cleaned string
    return s

def is_time_like(x):
    """Detect if cell looks like a time."""
    if x is None:
        return False
    if isinstance(x, (pd.Timestamp, datetime, time)):
        return True
    s = str(x).strip()
    if not s:
        return False
    # common time patterns: 08:30, 8:30, 08h30, 8h30
    if re.match(r'^\d{1,2}[:hH]\d{2}(\s*[APMapm\.]*)?$', s):
        return True
    # sometimes Excel stores fractional day as float -> pandas gives Timestamp usually; ignore floats
    return False

def to_time(x):
    """Convert a time-like cell to a datetime.time or None."""
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
    # replace 'h' with ':' for parsing convenience
    s2 = s.replace('h', ':').replace('H', ':')
    try:
        dt = dtparser.parse(s2, dayfirst=True)
        return dt.time()
    except Exception:
        return None

def to_date(x):
    """Convert a date-like cell to datetime.date or None."""
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
    # try parse with dateutil (dayfirst True)
    try:
        dt = dtparser.parse(s, dayfirst=True, fuzzy=True)
        return dt.date()
    except Exception:
        return None

# ---------------- Sheet parser ----------------

def find_week_rows(df):
    rows = []
    for i in range(len(df)):
        try:
            v = df.iat[i, 0]
        except Exception:
            v = None
        if isinstance(v, str) and re.match(r'^\s*S\s*\d+', v.strip(), re.I):
            rows.append(i)
    return rows

def find_slot_rows(df):
    rows = []
    for i in range(len(df)):
        try:
            v = df.iat[i, 0]
        except Exception:
            v = None
        if isinstance(v, str) and re.match(r'^\s*H\s*\d+', v.strip(), re.I):
            rows.append(i)
    return rows

def parse_sheet_to_events(xls, sheet_name):
    """
    Parse a timetable sheet (header=None) -> list of event dicts:
      { summary, teachers (set), description (set), start(datetime), end(datetime), groups (set) }
    """
    df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
    nrows, ncols = df.shape

    s_rows = find_week_rows(df)
    h_rows = find_slot_rows(df)

    raw_events = []

    for r in h_rows:
        # find preceding week header
        p_candidates = [s for s in s_rows if s <= r]
        if not p_candidates:
            continue
        p = max(p_candidates)
        date_row = p + 1
        group_row = p + 2

        # candidate columns are where date_row contains a date
        date_cols = [c for c in range(ncols) if date_row < nrows and to_date(df.iat[date_row, c]) is not None]

        for c in date_cols:
            for col in (c, c + 1):
                if col >= ncols:
                    continue

                # summary
                try:
                    summary = df.iat[r, col]
                except Exception:
                    summary = None
                if pd.isna(summary) or summary is None:
                    continue
                summary_str = str(summary).strip()
                if not summary_str:
                    continue

                # teacher (commonly at r+2)
                teacher = None
                if (r + 2) < nrows:
                    try:
                        t = df.iat[r + 2, col]
                        if not pd.isna(t):
                            teacher = str(t).strip()
                    except Exception:
                        teacher = None

                # try read small free-text description near the summary (e.g. 'Examen', 'Partiel', ...)
                desc_parts = []
                # scan a window below the summary but above times: r+1 .. r+6
                for idx in range(r + 1, min(r + 7, nrows)):
                    try:
                        cell = df.iat[idx, col]
                    except Exception:
                        cell = None
                    if pd.isna(cell) or cell is None:
                        continue
                    s = str(cell).strip()
                    if not s:
                        continue
                    # skip if looks like a time or a date or equals teacher/summary
                    if is_time_like(cell):
                        continue
                    if to_date(cell) is not None:
                        continue
                    if teacher and s == teacher:
                        continue
                    if s == summary_str:
                        continue
                    # otherwise it's plausible description text
                    desc_parts.append(s)
                desc_text = " | ".join(dict.fromkeys(desc_parts))  # unique preserving order

                # find start/end times (common offsets r+4/r+5) else search next rows
                start_val = None
                end_val = None
                if (r + 4) < nrows:
                    start_val = df.iat[r + 4, col]
                if (r + 5) < nrows:
                    end_val = df.iat[r + 5, col]

                # fallback heuristic scan for time-like cells
                if not is_time_like(start_val):
                    for off in range(1, 10):
                        idx = r + off
                        if idx >= nrows:
                            break
                        v = df.iat[idx, col]
                        if is_time_like(v):
                            start_val = v
                            break

                if not is_time_like(end_val):
                    for off in range(1, 12):
                        idx = r + off
                        if idx >= nrows:
                            break
                        v = df.iat[idx, col]
                        if is_time_like(v) and v != start_val:
                            end_val = v
                            break

                if not is_time_like(start_val) or not is_time_like(end_val):
                    # cannot determine times -> skip
                    continue

                start_t = to_time(start_val)
                end_t = to_time(end_val)
                if start_t is None or end_t is None:
                    continue

                # date (from date_row, column c)
                date_cell = df.iat[date_row, c]
                d = to_date(date_cell)
                if d is None:
                    continue

                dtstart = datetime.combine(d, start_t)
                dtend = datetime.combine(d, end_t)

                # read group labels
                gl = None
                gl_next = None
                if group_row < nrows:
                    try:
                        gl_raw = df.iat[group_row, col]
                        gl = normalize_group_label(gl_raw)
                    except Exception:
                        gl = None
                    if (col + 1) < ncols:
                        try:
                            gl_next_raw = df.iat[group_row, col + 1]
                            gl_next = normalize_group_label(gl_next_raw)
                        except Exception:
                            gl_next = None

                # detect merged summary spanning both group columns => "classe entière"
                is_left_col = (col == c)
                right_summary = None
                if (col + 1) < ncols:
                    try:
                        right_summary = df.iat[r, col + 1]
                    except Exception:
                        right_summary = None

                groups = set()
                if is_left_col and (pd.isna(right_summary) or right_summary is None) and gl and gl_next and gl != gl_next:
                    # left has summary, right empty, both group labels exist -> class entière
                    groups.add(gl)
                    groups.add(gl_next)
                else:
                    if gl:
                        groups.add(gl)

                # keep teacher and description as sets for merging later
                teachers = set()
                if teacher:
                    teachers.add(teacher)
                descriptions = set()
                if desc_text:
                    descriptions.add(desc_text)

                raw_events.append({
                    'summary': summary_str,
                    'teachers': teachers,
                    'descriptions': descriptions,
                    'start': dtstart,
                    'end': dtend,
                    'groups': groups
                })

    # merge events that have same summary+start+end (combine groups/teachers/descriptions)
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

    out = []
    for v in merged.values():
        out.append({
            'summary': v['summary'],
            'teachers': sorted(list(v['teachers'])),
            'description': " | ".join(sorted(list(v['descriptions']))) if v['descriptions'] else "",
            'start': v['start'],
            'end': v['end'],
            'groups': sorted(list(v['groups']))  # could be empty list
        })
    return out

# ---------------- ICS writer ----------------

def escape_ical_text(s: str) -> str:
    if s is None:
        return ""
    s = str(s)
    s = s.replace('\\', '\\\\')
    s = s.replace('\n', '\\n')
    s = s.replace(',', '\\,')
    s = s.replace(';', '\\;')
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
                # render as "Groupes: G 1 et G 2"
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
st.title('Convertisseur Emplois du Temps (Excel → .ics) — Fusion cellules & descriptions')

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
        events = parse_sheet_to_events(xls, sheet)
        if not events:
            st.warning(f'Aucun événement détecté dans {sheet} (vérifier la structure).')
            continue

        # preview table
        preview = pd.DataFrame([{
            'Résumé': ev['summary'],
            'Début': ev['start'],
            'Fin': ev['end'],
            'Enseignants': ' / '.join(ev['teachers']) if ev['teachers'] else '',
            'Groupes': ', '.join(ev['groups']) if ev['groups'] else '',
            'Description': ev['description']
        } for ev in events])
        st.write('Aperçu (10 premières lignes) :')
        st.dataframe(preview.head(10))

        ics = events_to_ics(events, tzname='Europe/Paris')
        st.download_button(label=f'Télécharger {sheet}.ics', data=ics, file_name=f'{sheet}.ics', mime='text/calendar')
