#!/usr/bin/env python3
"""Interactive pressure-vs-time analysis from Data/Data.xlsx raw data sheet.

Run:
    python interactive_pressure_analysis.py
Then open the printed URL in your browser.
"""

from __future__ import annotations

import argparse
import json
import re
import threading
import webbrowser
from dataclasses import dataclass
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from tempfile import TemporaryDirectory
from typing import Any
from zipfile import ZipFile
import xml.etree.ElementTree as ET

NS_MAIN = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
NS_REL = {"r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}


@dataclass
class TrialSeries:
    level: str
    trial: str
    time_s: list[float]
    pressure_kpa: list[float]
    temperature_c: list[float]


def col_to_index(col: str) -> int:
    value = 0
    for char in col:
        value = value * 26 + (ord(char) - ord("A") + 1)
    return value


def split_ref(cell_ref: str) -> tuple[str, int]:
    match = re.match(r"([A-Z]+)(\d+)", cell_ref)
    if not match:
        raise ValueError(f"Invalid cell reference: {cell_ref}")
    return match.group(1), int(match.group(2))


def read_shared_strings(xlsx_path: Path) -> list[str]:
    with ZipFile(xlsx_path) as zf:
        if "xl/sharedStrings.xml" not in zf.namelist():
            return []
        root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
    shared: list[str] = []
    for si in root.findall("m:si", NS_MAIN):
        text = "".join(t.text or "" for t in si.findall(".//m:t", NS_MAIN))
        shared.append(text)
    return shared


def get_raw_sheet_path(xlsx_path: Path) -> str:
    with ZipFile(xlsx_path) as zf:
        workbook = ET.fromstring(zf.read("xl/workbook.xml"))
        rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))

    rel_map = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels}
    for sheet in workbook.find("m:sheets", NS_MAIN):
        name = sheet.attrib.get("name", "").strip().lower()
        if name == "raw data":
            rid = sheet.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
            target = rel_map.get(rid, "")
            if not target:
                break
            return f"xl/{target}"
    raise ValueError("Could not find a sheet named 'Raw Data' in the workbook.")


def parse_raw_data_trials(xlsx_path: Path) -> list[TrialSeries]:
    shared = read_shared_strings(xlsx_path)
    sheet_path = get_raw_sheet_path(xlsx_path)

    with ZipFile(xlsx_path) as zf:
        sheet_xml = ET.fromstring(zf.read(sheet_path))

    rows = {}
    for row in sheet_xml.find("m:sheetData", NS_MAIN).findall("m:row", NS_MAIN):
        row_num = int(row.attrib["r"])
        cells: dict[str, str] = {}
        for cell in row.findall("m:c", NS_MAIN):
            ref = cell.attrib.get("r", "")
            col, _ = split_ref(ref)
            c_type = cell.attrib.get("t")
            value_node = cell.find("m:v", NS_MAIN)
            if value_node is None:
                continue
            value = value_node.text or ""
            if c_type == "s" and value:
                value = shared[int(value)]
            cells[col] = value
        rows[row_num] = cells

    level_row = rows.get(2, {})
    trial_row = rows.get(3, {})
    header_row = rows.get(4, {})
    data_row_numbers = sorted(r for r in rows if r >= 5)

    time_values: list[float] = []
    for r in data_row_numbers:
        time_raw = rows[r].get("A")
        if time_raw is None:
            continue
        try:
            time_values.append(float(time_raw))
        except ValueError:
            continue

    all_cols = sorted((c for c in header_row if c != "A"), key=col_to_index)
    temp_cols = [c for c in all_cols if "temperature" in header_row.get(c, "").lower()]

    trials: list[TrialSeries] = []
    current_level = "Unknown level"

    for temp_col in temp_cols:
        col_index = col_to_index(temp_col)
        pressure_col = ""
        for candidate in all_cols:
            if col_to_index(candidate) == col_index + 1 and "pressure" in header_row.get(candidate, "").lower():
                pressure_col = candidate
                break
        if not pressure_col:
            continue

        if temp_col in level_row and level_row[temp_col].strip():
            current_level = level_row[temp_col].strip()
        trial_name = trial_row.get(temp_col, f"Trial @ {temp_col}").strip()

        temperatures: list[float] = []
        pressures: list[float] = []
        times: list[float] = []

        for r in data_row_numbers:
            row = rows[r]
            t_raw = row.get("A")
            temp_raw = row.get(temp_col)
            p_raw = row.get(pressure_col)
            if t_raw is None or temp_raw is None or p_raw is None:
                continue
            try:
                times.append(float(t_raw))
                temperatures.append(float(temp_raw))
                pressures.append(float(p_raw))
            except ValueError:
                continue

        if times and pressures:
            trials.append(
                TrialSeries(
                    level=current_level,
                    trial=trial_name,
                    time_s=times,
                    pressure_kpa=pressures,
                    temperature_c=temperatures,
                )
            )

    return trials


def build_html(trials: list[TrialSeries]) -> str:
    payload = [
        {
            "label": f"{t.level} • {t.trial}",
            "level": t.level,
            "trial": t.trial,
            "time_s": t.time_s,
            "pressure_kpa": t.pressure_kpa,
            "temperature_c": t.temperature_c,
        }
        for t in trials
    ]

    return f"""<!doctype html>
<html lang=\"en\">
<head>
  <meta charset=\"utf-8\" />
  <meta name=\"viewport\" content=\"width=device-width,initial-scale=1\" />
  <title>Pressure vs Time Analyzer</title>
  <script src=\"https://cdn.plot.ly/plotly-2.35.2.min.js\"></script>
  <style>
    body {{ font-family: Arial, sans-serif; margin: 1rem; }}
    .panel {{ display: grid; grid-template-columns: 1fr; gap: 0.8rem; max-width: 1100px; }}
    #stats {{ border: 1px solid #ddd; border-radius: 8px; padding: 0.8rem; background: #fafafa; }}
    .muted {{ color: #666; font-size: 0.9rem; }}
    code {{ background: #f0f0f0; padding: 0.1rem 0.3rem; border-radius: 4px; }}
  </style>
</head>
<body>
  <h2>Raw Data Pressure vs Time Analyzer</h2>
  <div class=\"panel\">
    <label for=\"trialSelect\"><strong>Select trial:</strong></label>
    <select id=\"trialSelect\"></select>
    <div id=\"plot\" style=\"height: 620px;\"></div>
    <div id=\"stats\">
      <div><strong>Selected range stats</strong></div>
      <div id=\"statsContent\" class=\"muted\">Use box/lasso select on points to compute slope (pressure rise rate).</div>
      <p class=\"muted\">Tip: In the plot toolbar, choose <code>Box Select</code>, drag across a time range, and the slope is shown below in kPa/s.</p>
    </div>
  </div>

<script>
const trials = {json.dumps(payload)};
const selectEl = document.getElementById('trialSelect');
const plotEl = document.getElementById('plot');
const statsEl = document.getElementById('statsContent');

function linearRegression(x, y) {{
  const n = x.length;
  const xMean = x.reduce((a,b)=>a+b,0)/n;
  const yMean = y.reduce((a,b)=>a+b,0)/n;
  let num = 0;
  let den = 0;
  for (let i=0; i<n; i++) {{
    num += (x[i]-xMean)*(y[i]-yMean);
    den += (x[i]-xMean)*(x[i]-xMean);
  }}
  const slope = den === 0 ? NaN : num/den;
  const intercept = yMean - slope*xMean;
  return {{ slope, intercept }};
}}

function renderPlot(idx) {{
  const t = trials[idx];
  const trace = {{
    x: t.time_s,
    y: t.pressure_kpa,
    mode: 'lines+markers',
    type: 'scatter',
    marker: {{ size: 6 }},
    name: `${{t.level}} ${{t.trial}}`,
    customdata: t.temperature_c,
    hovertemplate: 'Time: %{{x:.2f}} s<br>Pressure: %{{y:.3f}} kPa<br>Temp: %{{customdata:.2f}} °C<extra></extra>'
  }};

  const layout = {{
    title: `${{t.level}} — ${{t.trial}}`,
    xaxis: {{ title: 'Time (s)' }},
    yaxis: {{ title: 'Pressure (kPa)' }},
    dragmode: 'select',
    hovermode: 'closest'
  }};

  Plotly.newPlot(plotEl, [trace], layout, {{responsive: true}});
  statsEl.innerHTML = 'Use box/lasso select on points to compute slope (pressure rise rate).';

  plotEl.on('plotly_selected', (eventData) => {{
    if (!eventData || !eventData.points || eventData.points.length < 2) {{
      statsEl.textContent = 'Need at least 2 selected points.';
      return;
    }}

    const points = eventData.points
      .map(p => ({{x: p.x, y: p.y}}))
      .sort((a,b) => a.x - b.x);

    const xs = points.map(p => p.x);
    const ys = points.map(p => p.y);
    const reg = linearRegression(xs, ys);
    const deltaP = ys[ys.length-1] - ys[0];
    const deltaT = xs[xs.length-1] - xs[0];

    statsEl.innerHTML = `
      Points selected: <strong>${{points.length}}</strong><br>
      Time range: <strong>${{xs[0].toFixed(2)}} to ${{xs[xs.length-1].toFixed(2)}} s</strong> (Δt=${{deltaT.toFixed(2)}} s)<br>
      Pressure range: <strong>${{ys[0].toFixed(3)}} to ${{ys[ys.length-1].toFixed(3)}} kPa</strong> (ΔP=${{deltaP.toFixed(3)}} kPa)<br>
      Linear fit pressure rise rate: <strong>${{Number.isFinite(reg.slope) ? reg.slope.toFixed(5) : 'N/A'}} kPa/s</strong>
    `;
  }});
}}

trials.forEach((t, idx) => {{
  const opt = document.createElement('option');
  opt.value = String(idx);
  opt.textContent = `${{String(idx+1).padStart(2, '0')}}. ${{t.label}}`;
  selectEl.appendChild(opt);
}});

selectEl.addEventListener('change', () => renderPlot(Number(selectEl.value)));
renderPlot(0);
</script>
</body>
</html>
"""


def run_server(html_content: str, host: str, port: int, open_browser: bool) -> None:
    with TemporaryDirectory() as tmp:
        root = Path(tmp)
        (root / "index.html").write_text(html_content, encoding="utf-8")

        class Handler(SimpleHTTPRequestHandler):
            def __init__(self, *args: Any, **kwargs: Any):
                super().__init__(*args, directory=str(root), **kwargs)

            def end_headers(self):
                self.send_header("Cache-Control", "no-cache, no-store, must-revalidate")
                self.send_header("Pragma", "no-cache")
                self.send_header("Expires", "0")
                super().end_headers()

        server = ThreadingHTTPServer((host, port), Handler)
        display_host = "127.0.0.1" if host == "0.0.0.0" else host
        url = f"http://{display_host}:{port}"
        print(f"Serving interactive analyzer at: {url}")
        print("Press Ctrl+C to stop.")

        if open_browser:
            threading.Timer(0.8, lambda: webbrowser.open(url)).start()

        try:
            server.serve_forever()
        except KeyboardInterrupt:
            print("\nServer stopped.")
        finally:
            server.server_close()


def main() -> None:
    parser = argparse.ArgumentParser(description="Interactive pressure-vs-time analyzer for Raw Data sheet.")
    parser.add_argument("--xlsx", default="Data/Data.xlsx", help="Path to the Excel workbook.")
    parser.add_argument("--host", default="0.0.0.0", help="Host interface for web server.")
    parser.add_argument("--port", type=int, default=5000, help="Local port for web server.")
    parser.add_argument("--no-browser", action="store_true", help="Do not auto-open browser.")
    args = parser.parse_args()

    xlsx_path = Path(args.xlsx)
    if not xlsx_path.exists():
        raise SystemExit(f"Workbook not found: {xlsx_path}")

    trials = parse_raw_data_trials(xlsx_path)
    if not trials:
        raise SystemExit("No trial data found in Raw Data sheet.")

    html = build_html(trials)
    run_server(html, host=args.host, port=args.port, open_browser=not args.no_browser)


if __name__ == "__main__":
    main()
