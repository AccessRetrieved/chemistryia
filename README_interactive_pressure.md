# Interactive Pressure-vs-Time Analyzer

This includes `interactive_pressure_analysis.py`, which reads `Data/Data.xlsx` and serves a modern browser UI for the **Raw Data** sheet.

## What it does

- Extracts all 40 trials from Raw Data.
- Shows scatter points connected by lines (pressure vs time).
- Uses a modern **Vue + Plotly** interface (served by Python stdlib).
- Lets you:
  - hover points for exact values,
  - drag-select points directly on the graph,
  - or manually enter a time range.
- Calculates pressure rise rate (linear-fit slope, kPa/s) for selected range.

## Run

```bash
python interactive_pressure_analysis.py
```

Open: `http://127.0.0.1:8050`

## Optional flags

```bash
python interactive_pressure_analysis.py --host 0.0.0.0 --port 9000 --no-browser
```

- `--host`: server interface (default `0.0.0.0`)
- `--port`: server port
- `--no-browser`: do not auto-open browser
