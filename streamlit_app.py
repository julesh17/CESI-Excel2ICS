import streamlit as st
import pandas as pd
import re
import uuid
from datetime import datetime, date, time, timedelta
import pytz

"""
Streamlit app : Convertit les feuilles d'un fichier Excel "EDT" (format tableau horaire) en fichiers .ics

- Conçu pour le fichier fourni (ex. feuilles "EDT P1" et "EDT P2").
- Lit chaque feuille sans en-têtes (header=None) et recherche :
  * les lignes de semaine (commençant par 'S 40', 'S 41', ...)
  * les blocs de créneaux commençant par 'H1', 'H2', ...
  * pour chaque créneau : titre (ligne Hn), enseignant (Hn+2), heure début (Hn+4), heure fin (Hn+5)
  * les dates des jours (ligne juste après la ligne "S ...") et les colonnes de groupes (G 1, G2)
- Produit un .ics (un par feuille sélectionnée) comprenant tous les événements (séances promo entières et séances de groupes).

Usage :
  streamlit run excel_to_ics_streamlit.py

"""

# ---------- Parsing helpers ----------

def find_week_rows(df):
    """Return indices of rows where column 0 contains 'S ' (semaine)."""
    s_rows = []
    for i in range(len(df)):
        v = df.iat[i, 0]
        if isinstance(v, str) and re.match(r'^\s*S\s*\d+', v.strip()):
            s_rows.append(i)
    return s_rows


def find_slot_rows(df):
    """Return indices of rows where column 0 contains 'H' (créneau H1, H2...)."""
    h_rows = []
    for i in range(len(df)):
        v = df.iat[i, 0]
        if isinstance(v, str) and re.match(r'^\s*H\d+', v.strip()):
            h_rows.append(i)
    return h_rows


def is_date_cell(x):
    return isinstance(x, (pd.Timestamp, datetime, date))


def is_time_cell(x):
    return isinstance(x, time) or isinstance(x, datetime) or isinstance(x, pd.Timestamp)


def parse_sheet_to_events(xls, sheet_name):
    """Parse a timetable-style sheet (no header) and return a list of events.

    Each event is a dict {summary, teacher, start (datetime), end (datetime), group_label}
    """
    df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
    nrows, ncols = df.shape

    s_rows = find_week_rows(df)
    h_rows = find_slot_rows(df)

    events = []
    for r in h_rows:
        # nearest preceding week header
        p_candidates = [s for s in s_rows if s <= r]
        if not p_candidates:
            continue
        p = max(p_candidates)
        date_row = p + 1
        group_row = p + 2

        # find columns where date_row contains a date (these are the left columns for each weekday)
        date_cols = [c for c in range(ncols) if date_row < nrows and is_date_cell(df.iat[date_row, c])]

        # for each weekday (left column c), consider that day may have up to 2 group columns: c and c+1
        for c in date_cols:
            for col in (c, c + 1):
                if col >= ncols:
                    continue
                summary = df.iat[r, col] if r < nrows else None
                if pd.isna(summary) or summary is None:
                    continue

                teacher = df.iat[r + 2, col] if (r + 2) < nrows else None
                if pd.isna(teacher):
                    teacher = None

                # start/end times are typically at r+4 and r+5 (observed in the sheet)
                start_val = df.iat[r + 4, col] if (r + 4) < nrows else None
                end_val = df.iat[r + 5, col] if (r + 5) < nrows else None

                # fallback: scan the next few rows to find time cells if not in the expected offset
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
                    # couldn't find a valid time pair for this cell -> skip
                    continue

                # normalize date (left column c)
                date_cell = df.iat[date_row, c]
                if isinstance(date_cell, pd.Timestamp):
                    d = date_cell.to_pydatetime().date()
                elif isinstance(date_cell, datetime):
                    d = date_cell.date()
                elif isinstance(date_cell, date):
                    d = date_cell
                else:
                    continue

                # normalize start/end to time
                if isinstance(start_val, pd.Timestamp) or isinstance(start_val, datetime):
                    start_t = start_val.to_pydatetime().time() if isinstance(start_val, pd.Timestamp) else start_val.time()
                elif isinstance(start_val, time):
                    start_t = start_val
                else:
                    # unknown format
                    continue

                if isinstance(end_val, pd.Timestamp) or isinstance(end_val, datetime):
                    end_t = end_val.to_pydatetime().time() if isinstance(end_val, pd.Timestamp) else end_val.time()
                elif isinstance(end_val, time):
                    end_t = end_val
                else:
                    continue

                dtstart = datetime.combine(d, start_t)
                dtend = datetime.combine(d, end_t)
                # naive -> localize later when writing ICS

                # attempt to read group label
                group_label = None
                if group_row < nrows:
                    gl = df.iat[group_row, col]
                    if not pd.isna(gl):
                        group_label = str(gl)

                events.append({
                    'summary': str(summary).strip(),
                    'teacher': str(teacher).strip() if teacher is not None else None,
                    'start': dtstart,
                    'end': dtend,
                    'group_label': group_label
                })
    return events


# ---------- ICS writer ----------

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
        if ev.get('group_label'):
            desc_lines.append(f"Groupe: {ev['group_label']}")
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


# ---------- Streamlit UI ----------

st.set_page_config(page_title='Excel → ICS (EDT)', layout='centered')
st.title('Convertisseur Emplois du Temps (Excel → .ics)')
st.markdown('Charge le fichier Excel (format fourni). Le script détecte automatiquement les feuilles "EDT P1" et "EDT P2" si elles existent.')

uploaded = st.file_uploader('Choisir le fichier Excel (.xlsx)', type=['xlsx'])

if uploaded is not None:
    try:
        xls = pd.ExcelFile(uploaded)
        sheets = xls.sheet_names
        st.write('Feuilles trouvées :', sheets)

        default = [s for s in ['EDT P1', 'EDT P2'] if s in sheets]
        to_convert = st.multiselect('Feuilles à convertir en .ics', options=sheets, default=default)

        if st.button('Générer les fichiers .ics'):
            for sheet in to_convert:
                with st.spinner(f'Conversion {sheet} ...'):
                    events = parse_sheet_to_events(uploaded, sheet)
                    if not events:
                        st.warning(f'Aucun événement trouvé pour la feuille {sheet} (vérifier la structure).')
                        continue
                    ics = events_to_ics(events, tzname='Europe/Paris')
                    n = len(events)
                    st.success(f'{n} événements extraits pour {sheet}')
                    st.download_button(label=f'Télécharger {sheet}.ics', data=ics, file_name=f'{sheet}.ics', mime='text/calendar')
    except Exception as e:
        st.error('Erreur lors de la lecture du fichier : ' + str(e))
else:
    st.info('Upload un fichier .xlsx pour commencer.')
