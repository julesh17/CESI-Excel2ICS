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

        # Gestion groupes
        if isinstance(groupe, str) and groupe.strip().upper() == "G 1":
            desc = desc + "\nGroupe: G 1"
        elif isinstance(groupe, str) and groupe.strip().upper() == "G 2":
            desc = desc + "\nGroupe: G 2"
        elif isinstance(groupe, str) and "classe entière" in groupe.lower():
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


def calcul_heures(df):
    if not set(["Titre", "Date", "Heure Début", "Heure Fin", "Groupe"]).issubset(df.columns):
        return pd.DataFrame()

    df2 = df.dropna(subset=["Titre", "Date", "Heure Début", "Heure Fin"])
    df2["Début"] = pd.to_datetime(df2["Date"].astype(str) + " " + df2["Heure Début"].astype(str), errors="coerce", dayfirst=True)
    df2["Fin"] = pd.to_datetime(df2["Date"].astype(str) + " " + df2["Heure Fin"].astype(str), errors="coerce", dayfirst=True)
    df2["Durée_h"] = (df2["Fin"] - df2["Début"]).dt.total_seconds() / 3600

    return df2.groupby(["Groupe", "Titre"]).agg({"Durée_h": "sum"}).reset_index()


def recap_par(df, par="Titre"):
    if not set(["Titre", "Date", "Heure Début", "Heure Fin", "Groupe", "Enseignant"]).issubset(df.columns):
        return pd.DataFrame()

    df2 = df.dropna(subset=["Titre", "Date", "Heure Début", "Heure Fin"])
    df2["Début"] = pd.to_datetime(df2["Date"].astype(str) + " " + df2["Heure Début"].astype(str), errors="coerce", dayfirst=True)
    df2["Fin"] = pd.to_datetime(df2["Date"].astype(str) + " " + df2["Heure Fin"].astype(str), errors="coerce", dayfirst=True)
    df2["Durée_h"] = (df2["Fin"] - df2["Début"]).dt.total_seconds() / 3600

    return df2.groupby([par]).agg({"Durée_h": "sum"}).reset_index()

# --- Interface Streamlit avec multipages ---
st.sidebar.title("Menu")
page = st.sidebar.radio("Aller à", ["Exporter ICS", "Heures par groupe", "Récapitulatif"])

uploaded_file = st.file_uploader("Choisissez le fichier Excel", type=["xlsx"])

if uploaded_file is not None:
    xls = pd.ExcelFile(uploaded_file)
    
    if page == "Exporter ICS":
        st.title("Export Emploi du Temps vers iCalendar (.ics)")
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

    elif page == "Heures par groupe":
        st.title("Calcul des heures par matière et par groupe")
        for sheet in ["EDT P1", "EDT P2"]:
            if sheet in xls.sheet_names:
                st.subheader(sheet)
                df = pd.read_excel(xls, sheet_name=sheet)
                heures = calcul_heures(df)
                if not heures.empty:
                    st.dataframe(heures)
                else:
                    st.warning("Structure de feuille non conforme pour ce calcul.")

    elif page == "Récapitulatif":
        st.title("Récapitulatif par enseignant ou matière")
        choix = st.radio("Regrouper par", ["Titre", "Enseignant"])
        for sheet in ["EDT P1", "EDT P2"]:
            if sheet in xls.sheet_names:
                st.subheader(sheet)
                df = pd.read_excel(xls, sheet_name=sheet)
                recap = recap_par(df, par=choix)
                if not recap.empty:
                    st.dataframe(recap)
                else:
                    st.warning("Structure de feuille non conforme pour ce calcul.")
