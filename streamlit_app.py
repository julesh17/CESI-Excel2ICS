import streamlit as st
import pandas as pd
import re
import uuid
from datetime import datetime, date, time
import pytz

"""
Corrected Streamlit script focused on handling merged cells ("classe entière").

How it works:
- Reads each sheet with header=None
- Locates week rows (lines starting with 'S') and slot rows (lines starting with 'H')
- For each slot, reads summary/teacher/start/end
- Detects when a summary sits in the left group column while the right group's summary is empty
  but the group labels row contains both G1 and G2 -> then treats the slot as "classe entière"
  and marks group_label = 'G 1 & G 2' (rendered in ICS as "Groupes: G 1 et G 2").

Save as a file and run with:
    streamlit run excel_to_ics_streamlit_corrected.py
"""


def normalize_group_label(x):
    if x is None:
        return None
    try:
        if pd.isna(x):
            return None
    except Exception:
        pass
    s = str(x).strip()
    m = re.search(r'G\s*\.?\s*(\d+)', s, re.I)
    if m:
        return f'G {m.group(1)}'
    m2 = re.search(r'^(?:groupe)?\s*(\d+)$', s, re.I)
    if m2:
        return f'G {m2.group(1)}'
    return s


def find_week_rows(df):
    s_rows = []
    for i in range(len(df)):
        try:
            v = df.iat[i, 0]
        except Exception:
            v = None
        if isinstance(v, str) and re.match(r'^\s*S\s*\d+', v.strip(), re.I):
            s_rows.append(i)
    return s_rows


def find_slot_rows(df):
    h_rows = []
    for i in range(len(df)):
        try:
            v = df.iat[i, 0]
        except Exception:
            v = None
        if isinstance(v, str) and re.match(r'^\s*H\s*\d+', v.strip(), re.I):
            h_rows.append(i)
    return h_rows


def is_date_cell(x):
    return isinstance(x, (pd.Timestamp, datetime, date))


def is_time_cell(x):
    return isinstance(x, time) or isinstance(x, datetime) or isinstance(x, pd.Timestamp)


def parse_sheet_to_events(xls, sheet_name):
    df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
    nrows, ncols = df.shape

    s_rows = find_week_rows(df)
    h_rows = find_slot_rows(df)

    events = []
    for r in h_rows:
        p_candidates = [s for s in s_rows if s <= r]
        if not p_candidates:
            continue
        p = max(p_candidates)
        date_row = p + 1
        group_row = p + 2

        date_cols = [c for c in range(ncols) if date_row < nrows and is_date_cell(df.iat[date_row, c])]

        for c in date_cols:
            for col in (c, c + 1):
                if col >= ncols:
                    continue
                try:
                    summary = df.iat[r, col]
                except Exception:
                    summary = None
                if pd.isna(summary) or summary is None:
                    continue

                teacher = None
                if (r + 2) < nrows:
                    try:
                        t = df.iat[r + 2, col]
                        if not pd.isna(t):
                            teacher = str(t).strip()
                    except Exception:
                        teacher = None

                start_val = None
                end_val = None
                if (r + 4) < nrows:
                    start_val = df.iat[r + 4, col]
                if (r + 5) < nrows:
                    end_val = df.iat[r + 5, col]

                if start_val is None or (isinstance(start_val, float) and pd.isna(start_val)):
                    for off in range(1, 9):
                        idx = r + off
                        if idx >= nrows:
                            break
                        v = df.iat[idx, col]
                        if is_time_cell(v):
                            start_val = v
                            break

                if end_val is None or (isinstance(end_val, float) and pd.isna(end_val)):
                    for off in range(1, 11):
                        idx = r + off
                        if idx >= nrows:
                            break
                        v = df.iat[idx, col]
                        if is_time_cell(v) and v != start_val:
                            end_val = v
                            break

                if start_val is None or end_val is None or pd.isna(start_val) or pd.isna(end_val):
                    continue

                date_cell = df.iat[date_row, c]
                if isinstance(date_cell, pd.Timestamp):
                    d = date_cell.to_pydatetime().date()
                elif isinstance(date_cell, datetime):
                    d = date_cell.date()
                elif isinstance(date_cell, date):
                    d = date_cell
                else:
                    continue

                if isinstance(start_val, pd.Timestamp) or isinstance(start_val, datetime):
                    start_t = start_val.to_pydatetime().time() if isinstance(start_val, pd.Timestamp) else start_val.time()
                elif isinstance(start_val, time):
                    start_t = start_val
                else:
                    continue

                if isinstance(end_val, pd.Timestamp) or isinstance(end_val, datetime):
                    end_t = end_val.to_pydatetime().time() if isinstance(end_val, pd.Timestamp) else end_val.time()
                elif isinstance(end_val, time):
                    end_t = end_val
                else:
                    continue

                dtstart = datetime.combine(d, start_t)
                dtend = datetime.combine(d, end_t)

                # group detection
                group_label = None
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

                is_left_col = (col == c)
                right_summary = None
                if (col + 1) < ncols:
                    try:
                        right_summary = df.iat[r, col + 1]
                    except Exception:
                        right_summary = None

                if is_left_col and (pd.isna(right_summary) or right_summary is None) and gl and gl_next and gl != gl_next:
                    group_label = f"{gl} & {gl_next}"
                else:
                    group_label = gl

                events.append({
                    'summary': str(summary).strip(),
                    'teacher': teacher,
                    'start': dtstart,
                    'end': dtend,
                    'group_label': group_label
                })
    return events


def escape_ical_text(s: str) -> str:
    if s is None:
        return ""
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
        if ev.get('teacher'):
            desc_lines.append(f"Enseignant: {ev['teacher']}")
        g = ev.get('group_label')
        if g:
            if '&' in g or ';' in g:
                g_readable = g.replace('&', ' et ').replace(';', ' et ')
                desc_lines.append(f"Groupes: {g_readable}")
            else:
                desc_lines.append(f"Groupe: {g}")
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


st.set_page_config(page_title='Excel → ICS (EDT)', layout='centered')
st.title('Convertisseur Emplois du Temps (Excel → .ics)')
st.markdown('Charge le fichier Excel (format fourni). Le script détecte automatiquement les feuilles "EDT P1" et "EDT P2" si elles existent.')

uploaded = st.file_uploader('Choisir le fichier Excel (.xlsx)', type=['xlsx'])

if uploaded is not None:
    try:
        xls = pd.ExcelFile(uploaded)
        st.write('Feuilles trouvées :', xls.sheet_names)

        for sheet in ["EDT P1", "EDT P2"]:
            if sheet in xls.sheet_names:
                events = parse_sheet_to_events(xls, sheet)
                if not events:
                    st.warning(f'Aucun événement trouvé pour la feuille {sheet} (vérifier la structure).')
                    continue
                ics = events_to_ics(events, tzname='Europe/Paris')
                st.write(f'{len(events)} événements extraits pour {sheet}')
                st.download_button(label=f'Télécharger {sheet}.ics', data=ics, file_name=f'{sheet}.ics', mime='text/calendar')
    except Exception as e:
        st.error('Erreur lors de la lecture du fichier : ' + str(e))
else:
    st.info('Upload un fichier .xlsx pour commencer.')
