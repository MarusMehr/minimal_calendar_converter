import streamlit as st
import pandas as pd
from ics import Calendar, Event
import datetime as dt
from io import BytesIO

# Funktion zur Erstellung des Kalenders
def create_minimal_ics(df):
    cal = Calendar()
    summary = df.at[1, 'Unnamed: 5']  # Zelle F2, mit tatsächlichem Pandas-Index in Zeile 1 und Spalte 5

    # Schleife über die relevanten Zeilen für die Start- und Enddaten
    for i in range(7, len(df) - 1):  # Start ab Zeile 8 (index 7) und Ende bei letzter relevanter Zeile
        start_date = df.at[i, 'Unnamed: 12']  # Spalte M für DTSTART
        end_date = df.at[i + 1, 'Unnamed: 12']  # Spalte M für DTEND
        location = df.at[i, 'Unnamed: 4']  # Spalte E für LOCATION

        # Sicherstellen, dass Start- und Enddaten vorhanden sind
        if pd.notna(start_date) and pd.notna(end_date):
            event = Event()
            event.name = summary
            event.begin = dt.datetime.strptime(str(start_date), '%Y-%m-%d')
            event.end = dt.datetime.strptime(str(end_date), '%Y-%m-%d')
            event.location = location if pd.notna(location) else "Unbekannter Ort"
            event.transparent = True  # Markiert das Ereignis als transparent (frei)
            cal.events.add(event)
    return cal

# Funktion zum Generieren der ICS-Datei
def generate_ics_file(cal):
    ics_file = BytesIO()
    ics_file.write(str(cal).encode('utf-8'))
    ics_file.seek(0)
    return ics_file

# Streamlit-Benutzeroberfläche
st.title("Minimalistic Excel to ICS Calendar Converter")

st.write("Lade eine Excel-Datei mit den Daten im korrekten Format hoch.")

# Datei-Upload-Bereich
uploaded_file = st.file_uploader("Wähle eine Excel-Datei aus", type="xlsx")

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file, engine='openpyxl')
        st.write("Vorschau der Daten:")
        st.write(df.head())

        # Kalender erstellen und Datei generieren
        cal = create_minimal_ics(df)
        ics_file = generate_ics_file(cal)
        
        st.download_button(
            label="Download ICS file",
            data=ics_file,
            file_name="minimal_calendar.ics",
            mime="text/calendar"
        )
    except ImportError:
        st.error("Fehler: Die `openpyxl`-Bibliothek ist nicht installiert.")
    except Exception as e:
        st.error(f"Es ist ein Fehler aufgetreten: {e}")
