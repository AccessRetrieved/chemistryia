# Interactive Pressure-vs-Time Analyzer

This repository now includes `interactive_pressure_analysis.py`, a no-dependency Python script (standard library only) that:

- Loads `Data/Data.xlsx`
- Reads the **Raw Data** sheet
- Extracts all 40 trials (8 temperature levels × 5 trials)
- Starts a local web app with an interactive chart (scatter points connected as a line)
- Lets you drag-select a time range and calculates pressure rise rate (slope in kPa/s)

## Run

```bash
python interactive_pressure_analysis.py
```

Then open: `http://127.0.0.1:8050`

## Controls

- Use the dropdown to switch between trials.
- Use **Box Select** (plot toolbar) and drag across data points.
- The panel below the graph shows:
  - selected time range
  - selected pressure range
  - linear-fit pressure rise rate (kPa/s)

## Optional flags

```bash
python interactive_pressure_analysis.py --host 0.0.0.0 --port 9000 --no-browser
```

- `--host`: set server host/interface (default `0.0.0.0`).
- `--port`: set web server port.
- `--no-browser`: don't auto-open your browser.
