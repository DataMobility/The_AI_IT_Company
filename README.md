
# Storingsanalyse Demo

Dit is een demo-project voor een storingsanalyse van voetbalstadions in Nederland, inclusief een kansberekening en voorspelling voor 2025. 
Het project bevat een Python-script dat willekeurige storingen genereert, een analyse uitvoert en een PowerPoint-rapport maakt met de resultaten.

## Inhoud

- `scripts/generate_report.py`: Python-script om data te genereren en een PowerPoint te maken
- `data/historische_storingen.csv`: Historische data van gegenereerde storingen
- `output/Storingsanalyse_Voorspelling_2025_Final_Branded_Presentatie.pptx`: PowerPoint-rapport met de analyse
- `output/voorspelde_storingen_2025.csv`: CSV met de kansberekening voor 2025

## Vereisten

- Python 3.x
- Virtuele omgeving (`venv`)

## Installatie

1. Clone deze repository.
2. Maak een virtuele omgeving aan:
    ```bash
    python3 -m venv venv
    source venv/bin/activate  # Op Windows: venv\Scripts\activate
    ```
3. Installeer vereiste packages:
    ```bash
    pip install -r requirements.txt
    ```

## Uitvoeren

Draai het script om de CSV-bestanden en PowerPoint te genereren:
```bash
python scripts/generate_report.py
```

De gegenereerde bestanden worden in de `output`-map geplaatst.
