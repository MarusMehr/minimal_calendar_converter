import streamlit as st
import pandas as pd
from ics import Calendar, Event
import datetime as dt
from io import BytesIO

def create_minimal_ics(df):
    cal = Calendar()
    for _, row in df.iterrows():
        event = Event()
        event.name = row['SUMMARY']  # Ereignisname
        event.begin = dt.datetime.strptime(str(row['DTSTART']), '%Y-%m-%d')
        event.end = dt.datetime.strptime(str(row['DTEND']), '%Y-%m-%d')
        event.location = row['LOCATION']  # Ort des Ereignisses
        event.transparent = True  # Setzt das Ereignis als transparent (freier Kalender)
        cal.events.add(event)
    return cal

def generate_ics_file(cal):
    ics_file = BytesIO()
    ics_file.write(str(cal).encode('utf-8'))
    ics_file.seek(0)
    return ics_file

st.title("Minimalistic Excel to ICS Calendar Converter")

st.write("Lade eine Excel-Datei mit den Spalten `SUMMARY`, `DTSTART`, `DTEND`, `LOCATION` hoch.")

uploaded_file = st.file_uploader("Wähle eine Excel-Datei aus", type="xlsx")

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file, engine='openpyxl')
        st.write("Vorschau der Daten:")
        st.write(df.head())

        # Überprüfen, ob die erforderlichen Spalten vorhanden sind
        if all(col in df.columns for col in ["SUMMARY", "DTSTART", "DTEND", "LOCATION"]):
            cal = create_minimal_ics(df)
            ics_file = generate_ics_file(cal)
            
            st.download_button(
                label="Download ICS file",
                data=ics_file,
                file_name="minimal_calendar.ics",
                mime="text/calendar"
            )
        else:
            st.error("Die Excel-Datei muss die Spalten `SUMMARY`, `DTSTART`, `DTEND`, `LOCATION` enthalten.")
    except ImportError:
        st.error("Fehler: Die `openpyxl`-Bibliothek ist nicht installiert.")
    except Exception as e:
        st.error(f"Es ist ein Fehler aufgetreten: {e}")

