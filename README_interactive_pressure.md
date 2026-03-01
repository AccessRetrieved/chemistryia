# Interactive Pressure-vs-Time Analyzer

This includes `interactive_pressure_analysis.py`, which reads `Data/Data.xlsx` and serves a modern browser UI for the **Raw Data** sheet.

## What it does

- Extracts all 40 trials from Raw Data.
- Shows scatter points connected by lines (pressure vs time).
- Uses a modern **Vue + Plotly** interface (served by Python stdlib).
- Lets you:
  - hover points for exact values,
  - drag-select points directly on the graph,
  - click any point to auto-analyze the following 15-second window,
  - or manually enter a time range.
- Calculates pressure rise rate (linear-fit slope, kPa/s) for selected ranges.

## How the initial rate is calculated

For any selected time range (drag selection, manual range, or 15-second click window), the app now shows two rate calculations:

- **Linear-fit pressure rise rate (primary)**: slope `m` from least-squares fit of pressure vs time,
  - model: `P = m·t + b`
  - shown in kPa/s, and the UI also reports `R²` for fit quality.
- **Endpoint slope**: simple two-point estimate `ΔP/Δt` from first to last selected point.

The **initial pressure rise rate** should be taken from the **linear-fit slope** over your chosen early-time interval.

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
