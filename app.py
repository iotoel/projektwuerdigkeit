import streamlit as st
import pandas as pd

def main():
    st.set_page_config(
        page_title="Projektwürdigkeitsanalyse - ICT - Empa",
        page_icon="📊",
        layout="wide"
    )
    
    st.title("📊 Projektwürdigkeitsanalyse - ICT - Empa")
    st.markdown("---")
    
    # Define criteria and options with exact Excel points (12 Kriterien aus Excel A1:Z30)
    criteria_data = {
        "Anzahl der involvierten Bereiche (Abteilungen, Teams usw.)": {
            "options": ["1 Bereich", "2 Bereiche", ">= 3 Bereiche"],
            "points": {"Change": 2, "Kleinprojekt": 2, "Projekt": 3},
            "help": "Beschreibt, wie viele organisatorische Einheiten aktiv an Planung oder Umsetzung beteiligt sind.\n\nWarum wichtig:\nJe mehr Bereiche involviert sind, desto höher sind Koordinationsaufwand, Abstimmungsbedarf und Governance."
        },
        "Grösse des gesamten Projektteams": {
            "options": ["< 2 Personen", "3-5 Personen", ">= 6 Personen"],
            "points": {"Change": 1, "Kleinprojekt": 2, "Projekt": 2},
            "help": "Anzahl Personen, die aktiv an Planung oder Umsetzung beteiligt sind.\n\nWarum wichtig:\nGrössere Teams erhöhen Koordinationsaufwand und Kommunikationsbedarf."
        },
        "Ressourcenmix": {
            "options": ["einfach (1 Rolle)", "mittel (2 Rollen)", "hoch (>3 Rollen)"],
            "points": {"Change": 1, "Kleinprojekt": 2, "Projekt": 3},
            "help": "Beschreibt die Vielfalt der benötigten Kompetenzen oder Rollen.\n\nWarum wichtig:\nJe mehr unterschiedliche Rollen (z. B. Infrastruktur, Security, Entwicklung, Betrieb) beteiligt sind, desto komplexer wird das Vorhaben."
        },
        "Personalaufwand": {
            "options": ["< 10 Tage", "11-30 Tage", ">= 31 Tage"],
            "points": {"Change": 1, "Kleinprojekt": 1, "Projekt": 1},
            "help": "Geschätzter Gesamtaufwand in Personentagen über alle Beteiligten.\n\nWarum wichtig:\nDer Aufwand ist einer der wichtigsten Indikatoren für die Projektgrösse."
        },
        "Investitionen": {
            "options": ["< 10k CHF", "11-30k CHF", ">= 31k CHF"],
            "points": {"Change": 1, "Kleinprojekt": 2, "Projekt": 3},
            "help": "Finanzielle Mittel, die für das Vorhaben benötigt werden (Software, Hardware, externe Leistungen).\n\nWarum wichtig:\nHöhere Investitionen erfordern stärkere Governance und Entscheidungsprozesse."
        },
        "Dauer": {
            "options": ["< 2 Wochen", "2-4 Wochen", ">= 1 Monat"],
            "points": {"Change": 1, "Kleinprojekt": 1, "Projekt": 1},
            "help": "Zeitspanne vom Start bis zum Abschluss des Vorhabens.\n\nWarum wichtig:\nLängere Vorhaben erfordern strukturierte Planung und Steuerung."
        },
        "Inhaltliche Komplexität": {
            "options": ["gering", "mittel", "hoch"],
            "points": {"Change": 1, "Kleinprojekt": 1, "Projekt": 1},
            "help": "Beschreibt, wie komplex die fachliche oder technische Lösung ist.\n\nWarum wichtig:\nKomplexe Lösungen erhöhen Risiko und Abstimmungsbedarf."
        },
        "Technologisches Risiko": {
            "options": ["gering", "mittel", "hoch"],
            "points": {"Change": 1, "Kleinprojekt": 1, "Projekt": 1},
            "help": "Grad der technischen Unsicherheit oder Neuartigkeit.\n\nWarum wichtig:\nNeue Technologien oder unbekannte Lösungen erhöhen das Projektrisiko."
        },
        "Strategische Bedeutung": {
            "options": ["gering (lokale Verbesserung)", "mittel (Verbesserung eines Services)", "hoch (strategisches IT-Thema)"],
            "points": {"Change": 1, "Kleinprojekt": 2, "Projekt": 3},
            "help": "Beitrag des Vorhabens zur strategischen Entwicklung der ICT oder der Organisation.\nBeispiele:\n- lokale Verbesserung\n- Verbesserung eines bestehenden Services\n- strategisches IT-Thema"
        },
        "Auswirkungen auf Betrieb": {
            "options": ["gering (keine Betriebsänderung)", "mittel (kleine Anpassung)", "hoch (neue Betriebsprozesse)"],
            "points": {"Change": 1, "Kleinprojekt": 1, "Projekt": 2},
            "help": "Beschreibt, wie stark der laufende IT-Betrieb betroffen ist.\n\nWarum wichtig:\nÄnderungen im Betrieb können Auswirkungen auf Support, Prozesse oder Infrastruktur haben."
        },
        "Sicherheits- / Compliance-Risiko": {
            "options": ["gering (keine sensiblen Daten)", "mittel (interne Daten)", "hoch (personenbezogene Daten)"],
            "points": {"Change": 1, "Kleinprojekt": 2, "Projekt": 3},
            "help": "Risiken im Zusammenhang mit Datenschutz, Informationssicherheit oder regulatorischen Anforderungen.\n\nWarum wichtig:\nWir sind in der Forschung und im Public Sector"
        },
        "Anzahl betroffene User": {
            "options": ["gering (<10)", "mittel (10-100)", "hoch (>200)"],
            "points": {"Change": 2, "Kleinprojekt": 2, "Projekt": 3},
            "help": "Wie viele Nutzerinnen und Nutzer von der Lösung betroffen sind.\n\nWarum wichtig:\nJe mehr User betroffen sind, desto grösser ist die Wirkung und das Risiko."
        }
    }
    
    st.subheader("📝 Kriterien bewerten")
    st.markdown("Bitte wählen Sie für jedes Kriterium die passende Option aus:")
    
    # Store user selections
    user_selections = {}
    
    # Create columns for better layout
    col1, col2 = st.columns(2)
    
    # Get all criteria items
    criteria_items = list(criteria_data.items())
    total_criteria = len(criteria_items)
    
    # Calculate split point (first half in left column, second half in right column)
    split_point = (total_criteria + 1) // 2  # Round up for odd numbers
    
    # Left column: first half of criteria
    for i in range(split_point):
        criterion, data = criteria_items[i]
        with col1:
            selection = st.radio(
                f"**{i+1}. {criterion}**",
                options=data["options"],
                key=f"criteria_left_{i}",
                index=0,
                help=data.get("help", "")
            )
            user_selections[criterion] = selection
    
    # Right column: second half of criteria
    for i in range(split_point, total_criteria):
        criterion, data = criteria_items[i]
        with col2:
            selection = st.radio(
                f"**{i+1}. {criterion}**",
                options=data["options"],
                key=f"criteria_right_{i}",
                index=0,
                help=data.get("help", "")
            )
            user_selections[criterion] = selection
    
    # st.markdown("---")
    
    # Calculate scores
    if st.button("🔍 Analyse durchführen", type="primary"):
        scores = {"Change": 0, "Kleinprojekt": 0, "Projekt": 0}
        
        for criterion, selection in user_selections.items():
            option_index = criteria_data[criterion]["options"].index(selection)
            points = criteria_data[criterion]["points"]
            
            # Excel logic: If selection matches option, add points for that category only
            if option_index == 0:  # Change option selected
                scores["Change"] += points["Change"]
            elif option_index == 1:  # Kleinprojekt option selected
                scores["Kleinprojekt"] += points["Kleinprojekt"]
            else:  # Projekt option selected
                scores["Projekt"] += points["Projekt"]
        
        # Display results
        # st.subheader("📈 Ergebnisse")
        
        # col_results1, col_results2, col_results3 = st.columns(3)
        
        # with col_results1:
        #     st.metric(
        #         label="🔄 Change",
        #         value=scores["Change"],
        #         delta=None
        #     )
        
        # with col_results2:
        #     st.metric(
        #         label="📋 Kleinprojekt",
        #         value=scores["Kleinprojekt"],
        #         delta=None
        #     )
        
        # with col_results3:
        #     st.metric(
        #         label="🚀 Projekt",
        #         value=scores["Projekt"],
        #         delta=None
        #     )
        
        # Determine recommendation using exact Excel logic
        # =IF(J7="","", (IF((AND(E19>G19,E19>I19)),D5,IF(AND(G19>E19,G19>I19),F5,H5))))
        change_score = scores["Change"]
        kleinprojekt_score = scores["Kleinprojekt"]
        projekt_score = scores["Projekt"]
        
        # Excel logic: Use strict > comparisons, default to Projekt
        if change_score > kleinprojekt_score and change_score > projekt_score:
            recommendation = "Change"
        elif kleinprojekt_score > change_score and kleinprojekt_score > projekt_score:
            recommendation = "Kleinprojekt"
        else:
            recommendation = "Projekt"  # Default case (including ties)
        
        st.markdown("---")
        st.subheader("🎯 Empfehlung")
        
        if recommendation == "Change":
            st.success(f"**Empfehlung: {recommendation}**")
            st.info("Es handelt sich um eine normale Betriebs- oder Verbesserungsarbeit.")
        elif recommendation == "Kleinprojekt":
            st.warning(f"**Empfehlung: {recommendation}**")
            st.info("Es handelt sich um ein begrenztes Vorhaben mit Struktur.")
        else:  # Projekt
            st.error(f"**Empfehlung: {recommendation}**")
            st.info("Es handelt sich um ein komplexes Vorhaben mit Governance.")
        
        # Show tie explanation if applicable
        if (change_score == kleinprojekt_score and change_score > projekt_score) or \
           (change_score == projekt_score and change_score > kleinprojekt_score) or \
           (kleinprojekt_score == projekt_score and kleinprojekt_score > change_score):
            st.info("ℹ️ **Hinweis:** Bei Gleichständen wird gemäss Excel-Logik 'Projekt' als Standard empfohlen.")

if __name__ == "__main__":
    main()
