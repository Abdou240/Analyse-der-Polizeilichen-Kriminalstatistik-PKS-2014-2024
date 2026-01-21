#!/bin/bash
# setup_repo.sh
# Dieses Skript erstellt die Verzeichnisstruktur und symbolische Links für das PKS-Analyseprojekt.
# Es sollte im Root-Verzeichnis des Projekts ausgeführt werden.

echo "Erstelle Verzeichnisstruktur für das PKS-Analyseprojekt..."

# Hauptverzeichnisse erstellen
mkdir -p data_converted
mkdir -p data/processed
mkdir -p figures
mkdir -p docs

echo "Verzeichnisse erstellt: data_converted/, data/processed/, figures/, docs/"

# HINWEIS: Dieses Skript erstellt nur die Struktur und beispielhafte Symlinks.
# Die eigentlichen Daten-Dateien (Excel, CSV, PNG, PDF, PPTX) müssen manuell in die
# entsprechenden Zielverzeichnisse (data_converted/, data/processed/, figures/, docs/) kopiert werden.
# Anschließend können die folgenden Befehle angepasst und ausgeführt werden, um die Symlinks zu erstellen.

echo ""
echo "================================================================"
echo "WICHTIG: Dieses Skript erstellt nur die leere Ordnerstruktur."
echo "Führen Sie die folgenden Schritte manuell durch:"
echo "1.  Kopieren Sie Ihre Roh-Excel-Dateien in 'data_converted/'."
echo "2.  Führen Sie die Python-Skripte aus, um die verarbeiteten Daten, Grafiken und Berichte zu generieren."
echo "3.  Erstellen Sie symbolische Links für die Abwärtskompatibilität."
echo "    Beispielbefehle (im Projekt-Root-Verzeichnis ausführen):"
echo "    ln -s data_converted/LA-F-02-T01-Laender-Faelle-HZ_xls_2019.xlsx LA-F-02-T01-Laender-Faelle-HZ_xls_2019.xlsx"
echo "    ln -s data/processed/laender_relevant_cases_2012_2024.csv laender_relevant_cases_2012_2024.csv"
echo "    ln -s figures/figure1_cases.png figure1_cases.png"
echo "================================================================"
