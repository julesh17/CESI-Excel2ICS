import streamlit as st
import pandas as pd
from dateutil import parser as dtparser
from datetime import timedelta
import uuid
import pytz

# --- Fonctions utilitaires ---

def parse_datetime(date_val, time_val=None):
    if pd.isna(date_val):
        return None
    if time_val is None or pd.isna(time_val):
        s = str(date_val)
    else:
        s = f"{date_val} {time_val}"
    try:
        return dtparser.parse(s, dayfirst=True)
    except Exception:
        return None


def ical_escape(text):
    if text is None:
        return ""
    return str(text).replace("\\", "\\\\").replace(",", "\\,").replace(";", "\\;").replace("\n", "\\n")


def make_ics(df, tz):
    events = []
    for _, row in df.iterrows():
        title = row.get("Titre") or row.get("Summary") or "Cours"
        date = row.get("Date") or row.get("Start")
        hdeb = row.get("Heure Début") or row.get("Start_time")
        hfin = row.get("Heure Fin") or row.get("End_time")
        loc = row.get("Lieu") or row.get("Location") or ""
        desc = row.get("Description") or ""
        groupe = row.get("Groupe") if "Groupe" in row else None

        dtstart = parse_datetime(date, hdeb)
        dtend = parse_datetime(date, hfin)

        if dtstart is None or dtend is None:
            continue

        dtstart = tz.localize(dtstart)
        dtend = tz.localize(dtend)

        # Gestion groupes :
        # - Si cellule fusionnée → Pandas la réplique dans chaque colonne → même valeur de groupe pour G1 et G2
        #   => on détecte ce cas et on force "Groupes: G 1 et G 2"
        if isinstance(groupe, str):
            g = groupe.strip().upper()
            if g == "G 1":
                desc = desc + "\nGroupe: G 1"
            elif g == "G 2":
                desc = desc + "\nGroupe: G 2"
            else:
                # toute autre valeur (issue d'une cellule fusionnée ou texte type "Classe entière")
                desc = desc + "\nGroupes: G 1 et G 2"

        uid = str(uuid.uuid4())

        event = f"""BEGIN:VEVENT
UID:{uid}
DTSTART;TZID=Europe/Paris:{dtstart.strftime('%Y%m%dT%H%M%S')}
DTEND;TZID=Europe/Paris:{dtend.strftime('%Y%m%dT%H%M%S')}
SUMMARY:{ical_escape(title)}
LOCATION:{ical_escape(loc)}
DESCRIPTION:{ical_escape(desc)}
END:VEVENT"""
        events.append(event)

    ics = "BEGIN:VCALENDAR\nVERSION:2.0\nPRODID:-//EDT Export//FR\n" + "\n".join(events) + "\nEND:VCALENDAR"
    return ics

# --- Interface Streamlit simple ---
st.title("Export Emploi du Temps vers iCalendar (.ics)")

uploaded_file = st.file_uploader("Choisissez le fichier Excel", type=["xlsx"])

if uploaded_file is not None:
    xls = pd.ExcelFile(uploaded_file)
    st.write("Feuilles trouvées :", xls.sheet_names)

    for sheet in ["EDT P1", "EDT P2"]:
        if sheet in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet)
            tz = pytz.timezone("Europe/Paris")
            ics_content = make_ics(df, tz)
            st.download_button(
                label=f"Télécharger {sheet}.ics",
                data=ics_content,
                file_name=f"{sheet}.ics",
                mime="text/calendar"
            )
