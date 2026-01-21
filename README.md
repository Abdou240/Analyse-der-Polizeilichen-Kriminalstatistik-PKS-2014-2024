# Analyse der Polizeilichen Kriminalstatistik (PKS) 2014-2024

Dieses Projekt analysiert die polizeiliche Kriminalstatistik (PKS) des Bundes sowie der Länder für die Jahre 2014 bis 2024. Es umfasst Python-Skripte zum Einlesen, Bereinigen und Auswerten der Rohdaten (Excel-Dateien), zur Erstellung von Grafiken und Präsentationen sowie zur Generierung eines formatierten Berichts im PDF-Format.

## Projektstruktur

Um Daten und Ausgaben übersichtlich zu halten, wurden die Dateien in Unterverzeichnisse einsortiert. Symbolische Links im Projekt-Root erhalten die ursprünglichen Dateinamen, sodass die Python-Skripte weiterhin funktionieren, ohne angepasst werden zu müssen.

Analyse-der-Polizeilichen-Kriminalstatistik-PKS-2014-2024-main/
├── data_converted/ # Alle rohen Excel-Tabellen für 'improved_main.py' und 'pks_analysis_finale_1.py'
│ ├── LA-F-02-T01-Laender-Faelle-HZ_xls.xlsx
│ ├── LA-F-02-T01-Laender-Faelle-HZ_xls_2014.xlsx
│ ├── ... (2015-2024)
│ └── T01-ZR-Bund-Fälle_xls.xlsx
├── data/
│ └── processed/ # Bereinigte bzw. intermediäre Daten, die von den Skripten erzeugt werden
│ ├── laender_relevant_cases_2012_2024.csv
│ ├── laender_relevant_aq_2012_2024.csv
│ ├── time_series_cases_clean_2014_2024.csv
│ ├── time_series_clearance_clean_2014_2024.csv
│ ├── tree_all_diebstahl.txt
│ └── tree_relevant_diebstahl.txt
├── figures/ # Alle generierten Visualisierungen
│ ├── figure1_cases.png
│ ├── figure2_clearance.png
│ ├── figure3_bar_states.png
│ └── figure4_scatter_cases_clearance.png
├── docs/ # Dokumentationen und Präsentationen
│ ├── PKS_Analysis_Paper_YYYYMMDD_HHMMSS.pdf
│ ├── PKS_presentation_professional.pptx
│ ├── Projektkonzept_PKS_abderrahmen.pdf
│ └── Projektkonzept_PKS_abderrahmen.pptx
├── LICENSE
├── README.md # Diese Datei
├── improved_main.py
├── pks_analysis_finale_1.py
├── generate_graphs_and_presentation.py
└── generate_paper.py

## Symlinks für Abwärtskompatibilität

Die Python-Skripte erwarten bestimmte Dateien im Projekt-Root. Damit sie unverändert weiterlaufen können, gibt es symbolische Links:

*   **Rohdaten**: Alle Excel-Tabellen aus `data_converted/` sind im Root verlinkt (z.B. `LA-F-02-T01-Laender-Faelle-HZ_xls_2019.xlsx` → `data_converted/LA-F-02-T01-Laender-Faelle-HZ_xls_2019.xlsx`).
*   **Verarbeitete Daten**: Die bereinigten CSVs und Textdateien in `data/processed/` besitzen Links im Root (z.B. `laender_relevant_cases_2012_2024.csv` → `data/processed/laender_relevant_cases_2012_2024.csv`).
*   **Grafiken**: Die generierten Abbildungen unter `figures/` sind über Links im Root erreichbar (z.B. `figure1_cases.png` → `figures/figure1_cases.png`).
*   **Bundesdaten**: Für die Datei `T01-ZR-Bund-Fälle_xls_xlsx` existiert ebenfalls ein Link von `data_converted/`.

## Python-Skripte

Die wichtigsten Python-Skripte in diesem Repository bleiben in ihrer Arbeitsweise unverändert. Sie gehen davon aus, dass die oben genannten Dateien entweder im Projekt-Root oder als Symlink vorhanden sind.

### `pks_analysis_finale_1.py`
Dieses Skript bereinigt die bundesweite Zeitreihen-Excel-Datei (`T01-ZR-Bund-Fälle_xls_xlsx`) und extrahiert Fallzahlen und Aufklärungsquoten für drei Deliktgruppen (Diebstahl insgesamt, Wohnungseinbruchdiebstahl und Kraftfahrzeugdiebstahl) in den Jahren 2014-2024.
*   **Funktion**: Konvertiert deutsche Zahlenformate in numerische Werte.
*   **Ausgabe**: Speichert die Ergebnisse als CSV-Dateien in `data/processed/` (`time_series_cases_clean_2014_2024.csv` und `time_series_clearance_clean_2014_2024.csv`). Zusätzlich werden die Baumstrukturen der PKS-Positionen in `tree_all_diebstahl.txt` und `tree_relevant_diebstahl.txt` abgelegt.

### `improved_main.py`
Dieses Skript liest die Länder-Tabellen `LA-F-02-T01-Laender-Faelle-HZ` für die Jahre 2014-2024 aus dem Verzeichnis `data_converted/`.
*   **Funktion**: Normalisiert die Spaltenüberschriften, erkennt Länder, Delikt-Schlüssel und die Spalten für Fallzahlen bzw. Aufklärungsquoten anhand von Synonymlisten und entfernt Zeilen, die sich auf den Bund beziehen.
*   **Ausgabe**: Zwei CSV-Dateien (`laender_relevant_cases_2012_2024.csv` und `laender_relevant_aq_2012_2024.csv`), die in `data/processed/` abgelegt werden.

### `generate_graphs_and_presentation.py`
Dieses Skript erstellt professionelle Grafiken mit `matplotlib` und `seaborn` und baut daraus eine PowerPoint-Präsentation.
*   **Eingabe**: Liest die bereinigten CSV-Dateien aus `data/processed/`.
*   **Ausgabe**: Speichert die generierten Abbildungen in `figures/` (z.B. `figure1_cases.png`, `figure2_clearance.png` usw.). Erstellt anschließend eine Präsentation im `docs/`-Ordner (`PKS_presentation_professional.pptx`), in die die Grafiken eingebettet werden.

### `generate_paper.py`
Dieses Skript erzeugt mit `reportlab` einen formatierten PDF-Bericht.
*   **Funktion**: Definiert verschiedene Absatzstile und setzt einen "Paper" mit Titelseite, Zusammenfassung und mehreren Abschnitten zusammen.
*   **Ausgabe**: Der Bericht wird im `docs/`-Verzeichnis mit einem Zeitstempel im Dateinamen ausgegeben (z.B. `PKS_Analysis_Paper_YYYYMMDD_HHMMSS.pdf`).

## Verwendung

1.  **Vorbereitung**: Stelle sicher, dass alle benötigten Excel-Dateien im Ordner `data_converted/` liegen. Die Datei `T01-ZR-Bund-Fälle_xls.xlsx` ist für die Bundeszeitreihen erforderlich, die Dateien `LA-F-02-T01-Laender-Faelle-HZ_xls_*.xlsx` für die Länderanalyse.
2.  **Zeitreihen bereinigen**:
    ```bash
    python pks_analysis_finale_1.py
    ```
    Erzeugt die CSVs `time_series_cases_clean_2014_2024.csv` und `time_series_clearance_clean_2014_2024.csv` sowie die Baumstrukturen unter `data/processed/`.
3.  **Länder-Daten extrahieren**:
    ```bash
    python improved_main.py
    ```
    Liest alle `LA-F-02`-Excel-Dateien aus `data_converted/` und speichert die länderbezogenen Fallzahlen und Aufklärungsquoten in `data/processed/`.
4.  **Grafiken und Präsentation erstellen**:
    ```bash
    python generate_graphs_and_presentation.py
    ```
    Erzeugt die Abbildungen in `figures/` und eine PowerPoint-Präsentation im Ordner `docs/`.
5.  **PDF-Bericht erzeugen**:
    ```bash
    python generate_paper.py
    ```
    Erstellt einen formatierten PDF-Bericht in `docs/`.

Durch die oben beschriebenen Pfade ist sichergestellt, dass alle Skripte ohne zusätzliche Anpassungen korrekt auf die Daten zugreifen und ihre Ergebnisse strukturiert ablegen.
