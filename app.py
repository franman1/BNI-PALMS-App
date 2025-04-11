import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
import io
import base64
from PIL import Image

# Seitenkonfiguration
st.set_page_config(
    page_title="BNI Chapter Gulda - PALMS Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Titel und Einführung
st.title("BNI Chapter Gulda - PALMS Dashboard")
st.markdown("### Vergleichen Sie Mitglieder und analysieren Sie Kennzahlen")

# Funktion zum Laden der Daten
@st.cache_data
def load_data(uploaded_file):
    try:
        # Versuche verschiedene Encodings und Trennzeichen
        encodings = ['utf-8', 'latin1', 'cp1252', 'iso-8859-1']
        separators = [',', ';', '\t']
        
        for encoding in encodings:
            for sep in separators:
                try:
                    df = pd.read_csv(uploaded_file, encoding=encoding, sep=sep)
                    # Wenn erfolgreich, bereinige die Spaltennamen
                    df.columns = df.columns.str.strip()
                    return df, None
                except Exception as e:
                    continue
        
        # Wenn CSV-Lesen fehlschlägt, versuche Excel
        try:
            df = pd.read_excel(uploaded_file)
            df.columns = df.columns.str.strip()
            return df, None
        except Exception as e:
            return None, f"Konnte die Datei nicht lesen: {str(e)}"
    
    except Exception as e:
        return None, f"Fehler beim Laden der Datei: {str(e)}"

# Funktion zum Erstellen eines herunterladbaren Links
def get_download_link(df, filename="bni_data.csv"):
    csv = df.to_csv(index=False)
    b64 = base64.b64encode(csv.encode()).decode()
    href = f'<a href="data:file/csv;base64,{b64}" download="{filename}">Download als CSV</a>'
    return href

# Sidebar für Datei-Upload und Filteroptionen
with st.sidebar:
    st.header("Daten-Upload")
    uploaded_file = st.file_uploader("PALMS-Bericht hochladen (CSV oder Excel)", type=['csv', 'xls', 'xlsx'])
    
    if uploaded_file is not None:
        df, error = load_data(uploaded_file)
        
        if error:
            st.error(error)
        elif df is not None:
            st.success("Daten erfolgreich geladen!")
            st.session_state['data'] = df
            st.session_state['file_loaded'] = True
            
            # Zeige Download-Link
            st.markdown(get_download_link(df), unsafe_allow_html=True)
            
            # Filteroptionen
            st.header("Filter und Sortierung")
            
            # Sortierkriterium
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
    1. Laden Sie Ihre PALMS-Berichtsdatei hoch
    2. Wählen Sie Sortierkriterien und Filter
    3. Vergleichen Sie Mitglieder in den Visualisierungen
    4. Wählen Sie einzelne Mitglieder für detaillierte Analysen
    
    Bei Fragen wenden Sie sich an den Support.
    """)

# Hauptbereich für Visualisierungen
if 'file_loaded' in st.session_state and st.session_state['file_loaded']:
    df = st.session_state['data']
    
    # Sortiere Daten nach ausgewähltem Kriterium
    sort_column = sort_options[selected_sort]
    df_sorted = df.sort_values(by=sort_column, ascending=sort_ascending)
    
    # Begrenze auf ausgewählte Anzahl von Mitgliedern
    df_display = df_sorted.head(num_members)
    
    # Tabs für verschiedene Visualisierungen
    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "Mitgliedervergleich", 
        "Anwesenheit", 
        "Empfehlungen", 
        "Besucher & 1-2-1", 
        "Umsatz & Bildung"
    ])
    
    with tab1:
        st.header("Mitgliedervergleich")
        
        # Mitgliederauswahl für detaillierten Vergleich
        col1, col2 = st.columns([1, 2])
        
        with col1:
            st.subheader("Mitglieder auswählen")
            all_members = df['Nachname'].tolist()
            selected_members = st.multiselect(
                "Wählen Sie Mitglieder zum Vergleichen:",
                options=all_members,
                default=all_members[:min(3, len(all_members))]
            )
            
            if not selected_members:
                st.warning("Bitte wählen Sie mindestens ein Mitglied aus.")
            
            # Kennzahlen auswählen
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
                default=list(metrics_options.keys())[:5]
            )
        
        with col2:
            if selected_members and selected_metrics:
                st.subheader("Vergleich der ausgewählten Mitglieder")
                
                # Filtere Daten für ausgewählte Mitglieder
                df_selected = df[df['Nachname'].isin(selected_members)]
                
                # Erstelle Vergleichsdiagramm
                fig, ax = plt.subplots(figsize=(10, 6))
                
                # Bereite Daten für Diagramm vor
                plot_data = []
                for metric in selected_metrics:
                    column = metrics_options[metric]
                    for _, row in df_selected.iterrows():
                        plot_data.append({
                            'Mitglied': row['Nachname'],
                            'Kennzahl': metric,
                            'Wert': row[column]
                        })
                
                plot_df = pd.DataFrame(plot_data)
                
                # Erstelle Diagramm
                sns.barplot(x='Mitglied', y='Wert', hue='Kennzahl', data=plot_df, ax=ax)
                plt.xticks(rotation=45)
                plt.tight_layout()
                
                st.pyplot(fig)
                
                # Zeige Tabelle mit ausgewählten Kennzahlen
                st.subheader("Detaillierte Daten")
                columns_to_show = ['Vorname', 'Nachname'] + [metrics_options[m] for m in selected_metrics]
                st.dataframe(df_selected[columns_to_show])
    
    with tab2:
        st.header("Anwesenheitsstatistiken")
        
        col1, col2 = st.columns(2)
        
        with col1:
            # Anwesenheitsverteilung nach Mitglied
            st.subheader(f"Top {num_members} Mitglieder nach {selected_sort}")
            
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
        
        with col2:
            # Gesamte Anwesenheitsverteilung
            st.subheader("Gesamte Anwesenheitsverteilung")
            
            fig, ax = plt.subplots(figsize=(10, 6))
            attendance_totals = [
                df['P'].sum(),
                df['A'].sum(),
                df['L'].sum(),
                df['M'].sum(),
                df['S'].sum()
            ]
            
            attendance_labels = ['Present', 'Absent', 'Late', 'Medical', 'Substitute']
            plt.bar(attendance_labels, attendance_totals, color=sns.color_palette("viridis", 5))
            plt.ylabel('Anzahl')
            plt.tight_layout()
            
            st.pyplot(fig)
        
        # Anwesenheitsstatistiken als Tabelle
        st.subheader("Anwesenheitsstatistiken pro Mitglied")
        st.dataframe(df_display[['Vorname', 'Nachname', 'P', 'A', 'L', 'M', 'S']])
    
    with tab3:
        st.header("Empfehlungsstatistiken")
        
        col1, col2 = st.columns(2)
        
        with col1:
            # Empfehlungen gegeben
            st.subheader("Empfehlungen gegeben (intern vs. extern)")
            
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
        
        with col2:
            # Empfehlungen erhalten
            st.subheader("Empfehlungen erhalten (intern vs. extern)")
            
            fig, ax = plt.subplots(figsize=(10, 6))
            referrals_received = pd.melt(
                df_display[['Nachname', 'R (Eigenbedarf)', 'R (extern)']], 
                id_vars=['Nachname'], 
                var_name='Typ', 
                value_name='Anzahl'
            )
            
            # Ersetze die Spaltennamen für bessere Lesbarkeit
            referrals_received['Typ'] = referrals_received['Typ'].replace({
                'R (Eigenbedarf)': 'Intern erhalten',
                'R (extern)': 'Extern erhalten'
            })
            
            sns.barplot(x='Nachname', y='Anzahl', hue='Typ', data=referrals_received, ax=ax)
            plt.xticks(rotation=45)
            plt.tight_layout()
            
            st.pyplot(fig)
        
        # Empfehlungsstatistiken als Tabelle
        st.subheader("Empfehlungsstatistiken pro Mitglied")
        
        # Berechne Gesamtempfehlungen
        df_display['Empfehlungen_gegeben'] = df_display['G (Eigenbedarf)'] + df_display['G (extern)']
        df_display['Empfehlungen_erhalten'] = df_display['R (Eigenbedarf)'] + df_display['R (extern)']
        
        st.dataframe(df_display[[
            'Vorname', 'Nachname', 
            'G (Eigenbedarf)', 'G (extern)', 'Empfehlungen_gegeben',
            'R (Eigenbedarf)', 'R (extern)', 'Empfehlungen_erhalten'
        ]])
    
    with tab4:
        st.header("Besucher und 1-2-1 Meetings")
        
        # Besucher und 1-2-1 Vergleich
        fig, ax = plt.subplots(figsize=(12, 6))
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
        plt.xticks(rotation=45)
        plt.tight_layout()
        
        st.pyplot(fig)
        
        # Besucher- und 1-2-1-Statistiken als Tabelle
        st.subheader("Besucher und 1-2-1 Meetings pro Mitglied")
        st.dataframe(df_display[['Vorname', 'Nachname', 'V', '1-2-1']])
    
    with tab5:
        st.header("Umsatz und Bildung")
        
        col1, col2 = st.columns(2)
        
        with col1:
            # Umsatzverteilung
            st.subheader("Umsatz pro Mitglied")
            
            fig, ax = plt.subplots(figsize=(10, 6))
            sns.barplot(x='Nachname', y='U', data=df_display, ax=ax)
            plt.xticks(rotation=45)
            plt.ylabel('Umsatz (€)')
            plt.tight_layout()
            
            st.pyplot(fig)
        
        with col2:
            # CTE und Testimonials
            st.subheader("CTE und Testimonials pro Mitglied")
            
            fig, ax = plt.subplots(figsize=(10, 6))
            cte_testimonials = pd.melt(
                df_display[['Nachname', 'CTE', 'T']], 
                id_vars=['Nachname'], 
                var_name='Kategorie', 
                value_name='Anzahl'
            )
            
            sns.barplot(x='Nachname', y='Anzahl', hue='Kategorie', data=cte_testimonials, ax=ax)
            plt.xticks(rotation=45)
            plt.tight_layout()
            
            st.pyplot(fig)
        
        # Umsatz- und Bildungsstatistiken als Tabelle
        st.subheader("Umsatz, CTE und Testimonials pro Mitglied")
        st.dataframe(df_display[['Vorname', 'Nachname', 'U', 'CTE', 'T']])
    
    # Rohdaten
    st.header("Rohdaten")
    st.dataframe(df)
    
    # Download-Button für die Daten
    st.markdown(get_download_link(df), unsafe_allow_html=True)

else:
    # Startseite, wenn noch keine Daten geladen wurden
    st.info("Bitte laden Sie eine PALMS-Berichtsdatei hoch, um Visualisierungen zu sehen.")
    
    # Beispielbilder oder Erklärungen
    st.header("BNI PALMS-Bericht Dashboard")
    st.markdown("""
    Dieses Dashboard hilft Ihnen, Ihre BNI Chapter-Daten zu visualisieren und zu analysieren.
    
    **Funktionen:**
    - Einfacher Upload von CSV- oder Excel-Dateien
    - Automatische Erkennung des Dateiformats und Encodings
    - Vergleich von Mitgliedern nach verschiedenen Kennzahlen
    - Sortierung und Filterung nach beliebigen Kriterien
    - Detaillierte Visualisierungen für alle wichtigen Kennzahlen
    
    **Unterstützte Kennzahlen:**
    - Anwesenheit (P, A, L, M, S)
    - Empfehlungen (gegeben und erhalten, intern und extern)
    - Besucher und 1-2-1 Meetings
    - Umsatz, CTE und Testimonials
    
    Laden Sie Ihre PALMS-Berichtsdatei hoch, um zu beginnen!
    """)

# Footer
st.markdown("---")
st.markdown("BNI Chapter Gulda PALMS Dashboard | Erstellt für den Chapterdirektor")
