import streamlit as st
import pandas as pd
import re
import uuid
from datetime import datetime, date, time
import pytz

"""
Streamlit app — conversion Excel (format fourni) → .ics + rapports

Améliorations :
- Détection et fusion des séances "classe entière" (mêmes résumé + mêmes horaires sur les deux colonnes de groupes)
  -> génère un seul VEVENT avec "Groupes: G 1 et G 2".
- Calcul des heures par matière et par groupe en se basant sur les événements extraits.
- Page de récapitulatif par matière ou par enseignant (explose les événements multi-enseignants si nécessaire).

Usage :
    streamlit run excel_to_ics_streamlit_fixed.py

Dépendances : pandas, streamlit, openpyxl, pytz
"""

# ---------------- Parsing (adapté au fichier fourni) ----------------

def find_week_rows(df):
    s_rows = []
    for i in range(len(df)):
        v = df.iat[i, 0]
        if isinstance(v, str) and re.match(r'^\s*S\s*\d+', v.strip()):
            s_rows.append(i)
    return s_rows


def find_slot_rows(df):
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


def parse_sheet_to_events_df(df):
    """Extrait la liste brute d'événements depuis la feuille (header=None attendue).

    Retourne une liste d'événements :
      {summary, teacher, start(datetime), end(datetime), group_label, r, date_col, col}
    """
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

        # colonnes contenant des dates (début de chaque jour)
        date_cols = [c for c in range(ncols) if date_row < nrows and is_date_cell(df.iat[date_row, c])]

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

                # on trouve les horaires (observé : r+4 / r+5)
                start_val = df.iat[r + 4, col] if (r + 4) < nrows else None
                end_val = df.iat[r + 5, col] if (r + 5) < nrows else None

                # fallback : recherche dans les lignes suivantes
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

                # normaliser la date du jour (col c)
                date_cell = df.iat[date_row, c]
                if isinstance(date_cell, pd.Timestamp):
                    d = date_cell.to_pydatetime().date()
                elif isinstance(date_cell, datetime):
                    d = date_cell.date()
                elif isinstance(date_cell, date):
                    d = date_cell
                else:
                    continue

                # normaliser heures
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

                group_label = None
                if group_row < nrows:
                    gl = df.iat[group_row, col]
                    if not pd.isna(gl):
                        group_label = str(gl).strip()

                events.append({
                    'summary': str(summary).strip(),
                    'teacher': str(teacher).strip() if teacher is not None else None,
                    'start': dtstart,
                    'end': dtend,
                    'group_label': group_label,
                    'r': r,
                    'date_col': c,
                    'col': col
                })
    return events


# ---------------- Fusion des doublons (classe entière) ----------------

def merge_events(events):
    """Fuse les événements qui correspondent à la même séance trouvée dans les deux colonnes G1/G2.

    Clef utilisée : (date, start, end, summary, r)
    Retourne une liste d'événements fusionnés :
      {summary, teachers (list), start, end, groups (list)}
    """
    merged = {}
    for e in events:
        key = (e['start'], e['end'], e['summary'], e['r'], e['date_col'])
        if key not in merged:
            merged[key] = {'summary': e['summary'], 'teachers': set(), 'start': e['start'], 'end': e['end'], 'groups': set()}
        if e.get('teacher'):
            merged[key]['teachers'].add(e['teacher'])
        if e.get('group_label'):
            merged[key]['groups'].add(e['group_label'])
        else:
            # si pas d'étiquette de groupe, on laisse vide (sera géré ensuite)
            pass

    out = []
    for v in merged.values():
        groups = sorted(list(v['groups']))
        teachers = sorted([t for t in v['teachers'] if t and t.lower() not in ['nan', 'none']])
        out.append({'summary': v['summary'], 'teachers': teachers, 'start': v['start'], 'end': v['end'], 'groups': groups})
    return out


# ---------------- Export ICS ----------------

def escape_ical_text(s: str) -> str:
    if s is None:
        return ""
    s = str(s)
    s = s.replace('\\', '\\\\')
    s = s.replace('\n', '\\n')
    s = s.replace(',', '\\,')
    s = s.replace(';', '\\;')
    return s


def merged_events_to_ics(merged_events, tzname='Europe/Paris'):
    tz = pytz.timezone(tzname)
    header = [
        'BEGIN:VCALENDAR',
        'VERSION:2.0',
        'PRODID:-//EDT Export//FR',
        'CALSCALE:GREGORIAN',
    ]
    body = []
    for ev in merged_events:
        uid = str(uuid.uuid4())
        dtstart = tz.localize(ev['start']).strftime('%Y%m%dT%H%M%S')
        dtend = tz.localize(ev['end']).strftime('%Y%m%dT%H%M%S')
        summary = escape_ical_text(ev['summary'])

        # description : enseignants + groupes
        desc_parts = []
        if ev.get('teachers'):
            desc_parts.append('Enseignant(s): ' + ' / '.join(ev['teachers']))
        if ev.get('groups'):
            if len(ev['groups']) == 1:
                desc_parts.append('Groupe: ' + ev['groups'][0])
            elif len(ev['groups']) > 1:
                desc_parts.append('Groupes: ' + ' et '.join(ev['groups']))

        description = escape_ical_text('\n'.join(desc_parts))

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


# ---------------- Tableaux / rapports ----------------

def merged_events_to_dataframe(merged_events):
    """Retourne un DataFrame où chaque ligne correspond à l'association (événement × groupe).

    Colonnes : group, summary, teacher, start, end, duration_h
    Si un événement a plusieurs groupes, on crée plusieurs lignes (une par groupe).
    Si un événement a plusieurs enseignants, on met l'ensemble join par ' / ' dans la colonne teacher.
    """
    rows = []
    for ev in merged_events:
        duration_h = (ev['end'] - ev['start']).total_seconds() / 3600.0
        teachers = ev['teachers'] if ev['teachers'] else []
        teacher_str = ' / '.join(teachers) if teachers else None
        groups = ev['groups'] if ev['groups'] else [None]
        if len(groups) == 0:
            groups = [None]
        for g in groups:
            rows.append({'group': g, 'summary': ev['summary'], 'teacher': teacher_str, 'start': ev['start'], 'end': ev['end'], 'duration_h': duration_h})
    df = pd.DataFrame(rows)
    return df


# ---------------- Streamlit UI ----------------

st.set_page_config(page_title='Excel → ICS (EDT)', layout='centered')
st.sidebar.title('Menu')
page = st.sidebar.radio('Aller à', ['Exporter ICS', 'Heures par groupe', 'Récapitulatif'])

uploaded = st.file_uploader('Choisir le fichier Excel (.xlsx)', type=['xlsx'])

if uploaded is None:
    st.info('Upload un fichier .xlsx (format fourni) pour commencer.')
    st.stop()

# lecture des feuilles
try:
    xls = pd.ExcelFile(uploaded)
    sheets = xls.sheet_names
except Exception as e:
    st.error('Impossible de lire le fichier Excel: ' + str(e))
    st.stop()

st.write('Feuilles trouvées :', sheets)

if page == 'Exporter ICS':
    st.title('Export .ics')
    sel = st.multiselect('Feuilles à convertir', options=sheets, default=[s for s in ['EDT P1','EDT P2'] if s in sheets])
    for sheet in sel:
        st.subheader(sheet)
        df = pd.read_excel(uploaded, sheet_name=sheet, header=None)
        events_raw = parse_sheet_to_events_df(df)
        if not events_raw:
            st.warning('Aucun événement trouvé (vérifier la structure de la feuille).')
            continue
        merged = merge_events(events_raw)
        ics = merged_events_to_ics(merged, tzname='Europe/Paris')
        st.write(f'{len(merged)} événements générés')
        st.download_button(label=f'Télécharger {sheet}.ics', data=ics, file_name=f'{sheet}.ics', mime='text/calendar')

elif page == 'Heures par groupe':
    st.title('Heures par matière et par groupe')
    sel = st.selectbox('Choisir une feuille', options=[s for s in sheets if s.startswith('EDT')])
    df = pd.read_excel(uploaded, sheet_name=sel, header=None)
    events_raw = parse_sheet_to_events_df(df)
    if not events_raw:
        st.warning('Aucun événement trouvé (vérifier la structure de la feuille).')
    else:
        merged = merge_events(events_raw)
        df_groups = merged_events_to_dataframe(merged)
        if df_groups.empty:
            st.warning('Aucun événement transformable en groupes (structure inattendue).')
        else:
            # total par groupe et par matière
            pivot = df_groups.groupby(['group', 'summary']).agg({'duration_h': 'sum'}).reset_index()
            st.write('Total heures par groupe / matière :')
            st.dataframe(pivot)
            # total par groupe
            tot = df_groups.groupby('group').agg({'duration_h': 'sum'}).reset_index().rename(columns={'duration_h':'total_heures'})
            st.write('Total heures par groupe :')
            st.dataframe(tot)

elif page == 'Récapitulatif':
    st.title('Récapitulatif (matière ou enseignant)')
    sel = st.selectbox('Choisir une feuille', options=[s for s in sheets if s.startswith('EDT')])
    df = pd.read_excel(uploaded, sheet_name=sel, header=None)
    events_raw = parse_sheet_to_events_df(df)
    if not events_raw:
        st.warning('Aucun événement trouvé (vérifier la structure de la feuille).')
    else:
        merged = merge_events(events_raw)
        df_groups = merged_events_to_dataframe(merged)
        if df_groups.empty:
            st.warning('Aucun événement transformable (structure inattendue).')
        else:
            mode = st.radio('Récapitulatif par', ['Matière', 'Enseignant'])
            if mode == 'Matière':
                recap = df_groups.groupby('summary').agg({'duration_h':'sum'}).reset_index().rename(columns={'duration_h':'total_heures'})
                st.dataframe(recap.sort_values('total_heures', ascending=False))
            else:
                # explode enseignants (peuvent être 'A / B')
                df_teach = df_groups.copy()
                df_teach['teacher'] = df_teach['teacher'].fillna('Inconnu')
                # split multiple teachers
                df_teach = df_teach.assign(teacher_split=df_teach['teacher'].str.split(' / ')).explode('teacher_split')
                recap = df_teach.groupby('teacher_split').agg({'duration_h':'sum'}).reset_index().rename(columns={'teacher_split':'Enseignant','duration_h':'total_heures'})
                st.dataframe(recap.sort_values('total_heures', ascending=False))

# ---------------- Fin ----------------
