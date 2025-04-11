import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
import io
import base64
from PIL import Image
import os
import tempfile

# Seitenkonfiguration
st.set_page_config(
    page_title="BNI Chapter Gulda - Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Titel und Einführung
st.title("BNI Chapter Gulda - Dashboard")
st.markdown("### Vergleichen Sie Mitglieder und analysieren Sie Kennzahlen")

# Funktion zum Erkennen des Dateiformats
def detect_file_format(df):
    """Erkennt das Format der Datei basierend auf den Spalten."""
    columns = set(df.columns)
    
    # Format 1: Pagisto-Format (neues Format)
    pagisto_columns = {"Datum", "Mitglied", "Platzierung", "Abwesenheit", "Empfehlungen", 
                       "Umsatzdanke", "Besucher", "121s", "Testimonials", "CTE", "Punkte"}
    
    # Format 2: PALMS-Format (altes Format)
    palms_columns = {"Vorname", "Nachname", "P", "A", "L", "M", "S", 
                     "G (Eigenbedarf)", "G (extern)", "R (Eigenbedarf)", "R (extern)",
                     "V", "1-2-1", "U", "CTE", "T"}
    
    # Überprüfe, welches Format am besten passt
    pagisto_match = len(columns.intersection(pagisto_columns)) / len(pagisto_columns)
    palms_match = len(columns.intersection(palms_columns)) / len(palms_columns)
    
    if pagisto_match > palms_match:
        return "pagisto"
    else:
        return "palms"

# Funktion zum Laden der Daten
@st.cache_data
def load_data(uploaded_file):
    try:
        # Speichere die hochgeladene Datei temporär
        with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(uploaded_file.name)[1]) as tmp_file:
            tmp_file.write(uploaded_file.getvalue())
            tmp_file_path = tmp_file.name
        
        # Versuche verschiedene Methoden zum Lesen der Datei
        df = None
        error_messages = []
        
        # 1. Wenn es eine Excel-Datei ist, versuche verschiedene Engines
        if uploaded_file.name.lower().endswith(('.xls', '.xlsx')):
            engines = ['openpyxl', 'xlrd']
            for engine in engines:
                try:
                    df = pd.read_excel(tmp_file_path, engine=engine)
                    st.success(f"Datei erfolgreich als Excel mit {engine} gelesen")
                    break
                except Exception as e:
                    error_messages.append(f"Fehler beim Lesen mit {engine}: {str(e)}")
        
        # 2. Wenn es keine Excel-Datei ist oder Excel-Lesen fehlgeschlagen ist, versuche CSV
        if df is None:
            # Versuche verschiedene Encodings und Trennzeichen
            encodings = ['utf-8', 'latin1', 'cp1252', 'iso-8859-1', 'windows-1252']
            separators = [',', ';', '\t']
            
            for encoding in encodings:
                for sep in separators:
                    try:
                        df = pd.read_csv(tmp_file_path, encoding=encoding, sep=sep)
                        st.success(f"Datei erfolgreich als CSV mit Encoding {encoding} und Trennzeichen '{sep}' gelesen")
                        break
                    except Exception as e:
                        error_messages.append(f"Fehler beim Lesen als CSV mit {encoding} und '{sep}': {str(e)}")
                if df is not None:
                    break
        
        # Lösche die temporäre Datei
        try:
            os.unlink(tmp_file_path)
        except:
            pass
        
        if df is not None:
            # Bereinige die Spaltennamen
            df.columns = df.columns.str.strip()
            
            # Erkenne das Dateiformat
            file_format = detect_file_format(df)
            
            # Verarbeite die Daten je nach Format
            if file_format == "pagisto":
                # Für Pagisto-Format: Extrahiere Vor- und Nachnamen
                if "Mitglied" in df.columns:
                    df[['Vorname', 'Nachname']] = df['Mitglied'].str.split(' ', n=1, expand=True)
                
                # Konvertiere Spalten zu numerischen Werten
                numeric_columns = ['Platzierung', 'Abwesenheit', 'Empfehlungen', 'Umsatzdanke', 
                                  'Besucher', '121s', 'Testimonials', 'CTE', 'Punkte']
                for col in numeric_columns:
                    if col in df.columns:
                        df[col] = pd.to_numeric(df[col], errors='coerce')
            
            elif file_format == "palms":
                # Für PALMS-Format: Konvertiere Spalten zu numerischen Werten
                numeric_columns = ['P', 'A', 'L', 'M', 'S', 'G (Eigenbedarf)', 'G (extern)', 
                                  'R (Eigenbedarf)', 'R (extern)', 'V', '1-2-1', 'U', 'CTE', 'T']
                for col in numeric_columns:
                    if col in df.columns:
                        df[col] = pd.to_numeric(df[col], errors='coerce')
            
            return df, file_format, None
        else:
            # Wenn alle Methoden fehlgeschlagen sind, gib die Fehlermeldungen zurück
            return None, None, "\n".join(error_messages)
    
    except Exception as e:
        return None, None, f"Fehler beim Laden der Datei: {str(e)}"

# Funktion zum Erstellen eines herunterladbaren Links
def get_download_link(df, filename="bni_data.csv"):
    csv = df.to_csv(index=False)
    b64 = base64.b64encode(csv.encode()).decode()
    href = f'<a href="data:file/csv;base64,{b64}" download="{filename}">Download als CSV</a>'
    return href

# Sidebar für Datei-Upload und Filteroptionen
with st.sidebar:
    st.header("Daten-Upload")
    uploaded_file = st.file_uploader("BNI-Bericht hochladen (CSV oder Excel)", type=['csv', 'xls', 'xlsx'])
    
    if uploaded_file is not None:
        df, file_format, error = load_data(uploaded_file)
        
        if error:
            st.error(f"Fehler beim Laden der Datei:\n{error}")
            st.info("Bitte stellen Sie sicher, dass die Datei im richtigen Format vorliegt und versuchen Sie es erneut.")
        elif df is not None:
            st.success(f"Daten erfolgreich geladen! Erkanntes Format: {file_format}")
            st.session_state['data'] = df
            st.session_state['file_format'] = file_format
            st.session_state['file_loaded'] = True
            
            # Zeige Download-Link
            st.markdown(get_download_link(df), unsafe_allow_html=True)
            
            # Filteroptionen
            st.header("Filter und Sortierung")
            
            # Sortierkriterium basierend auf dem Dateiformat
            if file_format == "pagisto":
                sort_options = {
                    'Platzierung': 'Platzierung',
                    'Abwesenheit': 'Abwesenheit',
                    'Empfehlungen': 'Empfehlungen',
                    'Umsatz': 'Umsatzdanke',
                    'Besucher': 'Besucher',
                    '1-2-1 Meetings': '121s',
                    'Testimonials': 'Testimonials',
                    'CTE': 'CTE',
                    'Punkte': 'Punkte'
                }
            else:  # palms
                sort_options = {
                    'Anwesenheit (P)': 'P',
                    'Abwesenheit (A)': 'A',
                    'Verspätung (L)': 'L',
                    'Medizinisch (M)': 'M',
                    'Vertretung (S)': 'S',
                    'Empfehlungen gegeben intern': 'G (Eigenbedarf)',
                    'Empfehlungen gegeben extern': 'G (extern)',
                    'Empfehlungen erhalten intern': 'R (Eigenbedarf)',
                    'Empfehlungen erhalten extern': 'R (extern)',
                    'Besucher (V)': 'V',
                    '1-2-1 Meetings': '1-2-1',
                    'Umsatz (U)': 'U',
                    'CTE': 'CTE',
                    'Testimonials (T)': 'T'
                }
            
            selected_sort = st.selectbox(
                "Sortieren nach:",
                options=list(sort_options.keys()),
                index=0
            )
            
            sort_ascending = st.checkbox("Aufsteigend sortieren", value=False)
            
            # Anzahl der anzuzeigenden Mitglieder
            if 'data' in st.session_state:
                max_members = len(st.session_state['data'])
                num_members = st.slider(
                    "Anzahl der Mitglieder anzeigen:",
                    min_value=1,
                    max_value=max_members,
                    value=min(10, max_members)
                )
    
    # Hilfe-Bereich
    st.header("Hilfe")
    st.markdown("""
    **Anleitung:**
    1. Laden Sie Ihre BNI-Berichtsdatei hoch
    2. Wählen Sie Sortierkriterien und Filter
    3. Vergleichen Sie Mitglieder in den Visualisierungen
    4. Wählen Sie einzelne Mitglieder für detaillierte Analysen
    
    Das Dashboard unterstützt sowohl das Pagisto-Format als auch das PALMS-Format.
    """)

# Hauptbereich für Visualisierungen
if 'file_loaded' in st.session_state and st.session_state['file_loaded']:
    df = st.session_state['data']
    file_format = st.session_state['file_format']
    
    # Sortiere Daten nach ausgewähltem Kriterium
    sort_column = sort_options[selected_sort]
    if sort_column in df.columns:
        df_sorted = df.sort_values(by=sort_column, ascending=sort_ascending)
    else:
        st.warning(f"Die Spalte '{sort_column}' wurde nicht gefunden. Die Daten werden nicht sortiert.")
        df_sorted = df
    
    # Begrenze auf ausgewählte Anzahl von Mitgliedern
    df_display = df_sorted.head(num_members)
    
    # Tabs für verschiedene Visualisierungen
    tab1, tab2, tab3, tab4 = st.tabs([
        "Mitgliedervergleich", 
        "Anwesenheit & Empfehlungen", 
        "Besucher & 1-2-1", 
        "Umsatz & Bildung"
    ])
    
    with tab1:
        st.header("Mitgliedervergleich")
        
        # Mitgliederauswahl für detaillierten Vergleich
        col1, col2 = st.columns([1, 2])
        
        with col1:
            st.subheader("Mitglieder auswählen")
            
            # Wähle die richtige Spalte für die Mitgliedernamen
            if file_format == "pagisto":
                all_members = df['Mitglied'].tolist()
            else:  # palms
                all_members = df['Nachname'].tolist()
                
            selected_members = st.multiselect(
                "Wählen Sie Mitglieder zum Vergleichen:",
                options=all_members,
                default=all_members[:min(3, len(all_members))]
            )
            
            if not selected_members:
                st.warning("Bitte wählen Sie mindestens ein Mitglied aus.")
            
            # Kennzahlen auswählen basierend auf dem Dateiformat
            if file_format == "pagisto":
                metrics_options = {
                    'Platzierung': 'Platzierung',
                    'Abwesenheit': 'Abwesenheit',
                    'Empfehlungen': 'Empfehlungen',
                    'Umsatz': 'Umsatzdanke',
                    'Besucher': 'Besucher',
                    '1-2-1 Meetings': '121s',
                    'Testimonials': 'Testimonials',
                    'CTE': 'CTE',
                    'Punkte': 'Punkte'
                }
            else:  # palms
                metrics_options = {
                    'Anwesenheit (P)': 'P',
                    'Abwesenheit (A)': 'A',
                    'Verspätung (L)': 'L',
                    'Medizinisch (M)': 'M',
                    'Vertretung (S)': 'S',
                    'Empfehlungen gegeben intern': 'G (Eigenbedarf)',
                    'Empfehlungen gegeben extern': 'G (extern)',
                    'Empfehlungen erhalten intern': 'R (Eigenbedarf)',
                    'Empfehlungen erhalten extern': 'R (extern)',
                    'Besucher (V)': 'V',
                    '1-2-1 Meetings': '1-2-1',
                    'Umsatz (U)': 'U',
                    'CTE': 'CTE',
                    'Testimonials (T)': 'T'
                }
            
            selected_metrics = st.multiselect(
                "Wählen Sie Kennzahlen zum Vergleichen:",
                options=list(metrics_options.keys()),
                default=list(metrics_options.keys())[:min(5, len(metrics_options))]
            )
        
        with col2:
            if selected_members and selected_metrics:
                st.subheader("Vergleich der ausgewählten Mitglieder")
                
                # Filtere Daten für ausgewählte Mitglieder
                if file_format == "pagisto":
                    df_selected = df[df['Mitglied'].isin(selected_members)]
                else:  # palms
                    df_selected = df[df['Nachname'].isin(selected_members)]
                
                # Erstelle Vergleichsdiagramm
                fig, ax = plt.subplots(figsize=(10, 6))
                
                # Bereite Daten für Diagramm vor
                plot_data = []
                for metric in selected_metrics:
                    column = metrics_options[metric]
                    if column in df_selected.columns:
                        for _, row in df_selected.iterrows():
                            member_name = row['Mitglied'] if file_format == "pagisto" else row['Nachname']
                            plot_data.append({
                                'Mitglied': member_name,
                                'Kennzahl': metric,
                                'Wert': row[column]
                            })
                
                if plot_data:
                    plot_df = pd.DataFrame(plot_data)
                    
                    # Erstelle Diagramm
                    sns.barplot(x='Mitglied', y='Wert', hue='Kennzahl', data=plot_df, ax=ax)
                    plt.xticks(rotation=45)
                    plt.tight_layout()
                    
                    st.pyplot(fig)
                    
                    # Zeige Tabelle mit ausgewählten Kennzahlen
                    st.subheader("Detaillierte Daten")
                    
                    # Wähle die richtigen Spalten basierend auf dem Format
                    if file_format == "pagisto":
                        columns_to_show = ['Mitglied'] + [metrics_options[m] for m in selected_metrics if metrics_options[m] in df_selected.columns]
                    else:  # palms
                        columns_to_show = ['Vorname', 'Nachname'] + [metrics_options[m] for m in selected_metrics if metrics_options[m] in df_selected.columns]
                    
                    st.dataframe(df_selected[columns_to_show])
                else:
                    st.warning("Keine Daten für die ausgewählten Kennzahlen gefunden.")
    
    with tab2:
        st.header("Anwesenheit & Empfehlungen")
        
        col1, col2 = st.columns(2)
        
        with col1:
            # Anwesenheitsstatistiken
            st.subheader("Anwesenheitsstatistiken")
            
            if file_format == "pagisto":
                # Für Pagisto-Format: Zeige Abwesenheit
                fig, ax = plt.subplots(figsize=(10, 6))
                sns.barplot(x='Mitglied', y='Abwesenheit', data=df_display, ax=ax)
                plt.xticks(rotation=45)
                plt.tight_layout()
                st.pyplot(fig)
                
                # Zeige Tabelle
                st.dataframe(df_display[['Mitglied', 'Abwesenheit']])
                
            else:  # palms
                # Für PALMS-Format: Zeige P, A, L, M, S
                fig, ax = plt.subplots(figsize=(10, 6))
                attendance_data = pd.melt(
                    df_display[['Nachname', 'P', 'A', 'L', 'M', 'S']], 
                    id_vars=['Nachname'], 
                    var_name='Status', 
                    value_name='Anzahl'
                )
                
                sns.barplot(x='Nachname', y='Anzahl', hue='Status', data=attendance_data, ax=ax)
                plt.xticks(rotation=45)
                plt.tight_layout()
                st.pyplot(fig)
                
                # Zeige Tabelle
                st.dataframe(df_display[['Vorname', 'Nachname', 'P', 'A', 'L', 'M', 'S']])
        
        with col2:
            # Empfehlungsstatistiken
            st.subheader("Empfehlungsstatistiken")
            
            if file_format == "pagisto":
                # Für Pagisto-Format: Zeige Empfehlungen
                fig, ax = plt.subplots(figsize=(10, 6))
                sns.barplot(x='Mitglied', y='Empfehlungen', data=df_display, ax=ax)
                plt.xticks(rotation=45)
                plt.tight_layout()
                st.pyplot(fig)
                
                # Zeige Tabelle
                st.dataframe(df_display[['Mitglied', 'Empfehlungen']])
                
            else:  # palms
                # Für PALMS-Format: Zeige G (Eigenbedarf), G (extern), etc.
                fig, ax = plt.subplots(figsize=(10, 6))
                referrals_given = pd.melt(
                    df_display[['Nachname', 'G (Eigenbedarf)', 'G (extern)']], 
                    id_vars=['Nachname'], 
                    var_name='Typ', 
                    value_name='Anzahl'
                )
                
                # Ersetze die Spaltennamen für bessere Lesbarkeit
                referrals_given['Typ'] = referrals_given['Typ'].replace({
                    'G (Eigenbedarf)': 'Intern gegeben',
                    'G (extern)': 'Extern gegeben'
                })
                
                sns.barplot(x='Nachname', y='Anzahl', hue='Typ', data=referrals_given, ax=ax)
                plt.xticks(rotation=45)
                plt.tight_layout()
                st.pyplot(fig)
                
                # Zeige Tabelle
                st.dataframe(df_display[['Vorname', 'Nachname', 'G (Eigenbedarf)', 'G (extern)', 'R (Eigenbedarf)', 'R (extern)']])
    
    with tab3:
        st.header("Besucher & 1-2-1 Meetings")
        
        # Besucher und 1-2-1 Vergleich
        fig, ax = plt.subplots(figsize=(12, 6))
        
        if file_format == "pagisto":
            # Für Pagisto-Format
            visitors_121 = df_display[['Mitglied', 'Besucher', '121s']]
            visitors_121_melted = pd.melt(
                visitors_121, 
                id_vars=['Mitglied'], 
                var_name='Kategorie', 
                value_name='Anzahl'
            )
            
            # Ersetze die Spaltennamen für bessere Lesbarkeit
            visitors_121_melted['Kategorie'] = visitors_121_melted['Kategorie'].replace({
                'Besucher': 'Besucher',
                '121s': '1-2-1 Meetings'
            })
            
            sns.barplot(x='Mitglied', y='Anzahl', hue='Kategorie', data=visitors_121_melted, ax=ax)
            
            # Zeige Tabelle
            st.dataframe(df_display[['Mitglied', 'Besucher', '121s']])
            
        else:  # palms
            # Für PALMS-Format
            visitors_121 = df_display[['Nachname', 'V', '1-2-1']]
            visitors_121_melted = pd.melt(
                visitors_121, 
                id_vars=['Nachname'], 
                var_name='Kategorie', 
                value_name='Anzahl'
            )
            
            # Ersetze die Spaltennamen für bessere Lesbarkeit
            visitors_121_melted['Kategorie'] = visitors_121_melted['Kategorie'].replace({
                'V': 'Besucher',
                '1-2-1': '1-2-1 Meetings'
            })
            
            sns.barplot(x='Nachname', y='Anzahl', hue='Kategorie', data=visitors_121_melted, ax=ax)
            
            # Zeige Tabelle
            st.dataframe(df_display[['Vorname', 'Nachname', 'V', '1-2-1']])
        
        plt.xticks(rotation=45)
        plt.tight_layout()
        st.pyplot(fig)
    
    with tab4:
        st.header("Umsatz & Bildung")
        
        col1, col2 = st.columns(2)
        
        with col1:
            # Umsatzverteilung
            st.subheader("Umsatz pro Mitglied")
            
            fig, ax = plt.subplots(figsize=(10, 6))
            
            if file_format == "pagisto":
                # Für Pagisto-Format
                sns.barplot(x='Mitglied', y='Umsatzdanke', data=df_display, ax=ax)
                
                # Zeige Tabelle
                st.dataframe(df_display[['Mitglied', 'Umsatzdanke']])
                
            else:  # palms
                # Für PALMS-Format
                sns.barplot(x='Nachname', y='U', data=df_display, ax=ax)
                
                # Zeige Tabelle
                st.dataframe(df_display[['Vorname', 'Nachname', 'U']])
            
            plt.xticks(rotation=45)
            plt.ylabel('Umsatz (€)')
            plt.tight_layout()
            st.pyplot(fig)
        
        with col2:
            # CTE und Testimonials
            st.subheader("CTE und Testimonials pro Mitglied")
            
            fig, ax = plt.subplots(figsize=(10, 6))
            
            if file_format == "pagisto":
                # Für Pagisto-Format
                cte_testimonials = pd.melt(
                    df_display[['Mitglied', 'CTE', 'Testimonials']], 
                    id_vars=['Mitglied'], 
                    var_name='Kategorie', 
                    value_name='Anzahl'
                )
                
                sns.barplot(x='Mitglied', y='Anzahl', hue='Kategorie', data=cte_testimonials, ax=ax)
                
                # Zeige Tabelle
                st.dataframe(df_display[['Mitglied', 'CTE', 'Testimonials']])
                
            else:  # palms
                # Für PALMS-Format
                cte_testimonials = pd.melt(
                    df_display[['Nachname', 'CTE', 'T']], 
                    id_vars=['Nachname'], 
                    var_name='Kategorie', 
                    value_name='Anzahl'
                )
                
                # Ersetze die Spaltennamen für bessere Lesbarkeit
                cte_testimonials['Kategorie'] = cte_testimonials['Kategorie'].replace({
                    'T': 'Testimonials'
                })
                
                sns.barplot(x='Nachname', y='Anzahl', hue='Kategorie', data=cte_testimonials, ax=ax)
                
                # Zeige Tabelle
                st.dataframe(df_display[['Vorname', 'Nachname', 'CTE', 'T']])
            
            plt.xticks(rotation=45)
            plt.tight_layout()
            st.pyplot(fig)
    
    # Rohdaten
    st.header("Rohdaten")
    st.dataframe(df)
    
    # Download-Button für die Daten
    st.markdown(get_download_link(df), unsafe_allow_html=True)

else:
    # Startseite, wenn noch keine Daten geladen wurden
    st.info("Bitte laden Sie eine BNI-Berichtsdatei hoch, um Visualisierungen zu sehen.")
    
    # Beispielbilder oder Erklärungen
    st.header("BNI Chapter Gulda - Dashboard")
    st.markdown("""
    Dieses Dashboard hilft Ihnen, Ihre BNI Chapter-Daten zu visualisieren und zu analysieren.
    
    **Funktionen:**
    - Einfacher Upload von CSV- oder Excel-Dateien
    - Automatische Erkennung des Dateiformats (Pagisto oder PALMS)
    - Vergleich von Mitgliedern nach verschiedenen Kennzahlen
    - Sortierung und Filterung nach beliebigen Kriterien
    - Detaillierte Visualisierungen für alle wichtigen Kennzahlen
    
    **Unterstützte Kennzahlen:**
    - Anwesenheit / Abwesenheit
    - Empfehlungen
    - Besucher und 1-2-1 Meetings
    - Umsatz, CTE und Testimonials
    - Punkte und Platzierung
    
    Laden Sie Ihre BNI-Berichtsdatei hoch, um zu beginnen!
    """)

# Footer
st.markdown("---")
st.markdown("BNI Chapter Gulda Dashboard | Erstellt für den Chapterdirektor")
