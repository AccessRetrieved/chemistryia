#!/usr/bin/env python3
"""Modern interactive pressure-vs-time analyzer for Data/Data.xlsx Raw Data sheet."""

from __future__ import annotations

import argparse
import json
import re
import threading
import webbrowser
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from tempfile import TemporaryDirectory
from typing import Any
from zipfile import ZipFile

NS_MAIN = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


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
    values: list[str] = []
    for si in root.findall("m:si", NS_MAIN):
        values.append("".join(t.text or "" for t in si.findall(".//m:t", NS_MAIN)))
    return values


def get_raw_sheet_path(xlsx_path: Path) -> str:
    with ZipFile(xlsx_path) as zf:
        workbook = ET.fromstring(zf.read("xl/workbook.xml"))
        rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))

    rel_map = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels}
    for sheet in workbook.find("m:sheets", NS_MAIN):
        name = sheet.attrib.get("name", "").strip().lower()
        if name != "raw data":
            continue
        rel_id = sheet.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id", "")
        target = rel_map.get(rel_id, "")
        if target:
            return f"xl/{target}"
    raise ValueError("Could not find a sheet named 'Raw Data' in the workbook.")


def parse_raw_data_trials(xlsx_path: Path) -> list[TrialSeries]:
    shared_strings = read_shared_strings(xlsx_path)
    sheet_path = get_raw_sheet_path(xlsx_path)

    with ZipFile(xlsx_path) as zf:
        sheet = ET.fromstring(zf.read(sheet_path))

    rows: dict[int, dict[str, str]] = {}
    for row in sheet.find("m:sheetData", NS_MAIN).findall("m:row", NS_MAIN):
        row_num = int(row.attrib["r"])
        row_cells: dict[str, str] = {}
        for cell in row.findall("m:c", NS_MAIN):
            ref = cell.attrib.get("r", "")
            col, _ = split_ref(ref)
            value_node = cell.find("m:v", NS_MAIN)
            if value_node is None:
                continue
            raw_value = value_node.text or ""
            if cell.attrib.get("t") == "s" and raw_value:
                raw_value = shared_strings[int(raw_value)]
            row_cells[col] = raw_value
        rows[row_num] = row_cells

    level_row = rows.get(2, {})
    trial_row = rows.get(3, {})
    header_row = rows.get(4, {})
    data_rows = sorted(r for r in rows if r >= 5)

    all_cols = sorted((c for c in header_row if c != "A"), key=col_to_index)
    temp_cols = [c for c in all_cols if "temperature" in header_row.get(c, "").lower()]

    trials: list[TrialSeries] = []
    current_level = "Unknown level"

    for temp_col in temp_cols:
        temp_idx = col_to_index(temp_col)
        pressure_col = next(
            (
                col
                for col in all_cols
                if col_to_index(col) == temp_idx + 1 and "pressure" in header_row.get(col, "").lower()
            ),
            "",
        )
        if not pressure_col:
            continue

        if level_row.get(temp_col, "").strip():
            current_level = level_row[temp_col].strip()
        trial_name = trial_row.get(temp_col, f"Trial @ {temp_col}").strip()

        times: list[float] = []
        temperatures: list[float] = []
        pressures: list[float] = []

        for row_num in data_rows:
            row = rows[row_num]
            if "A" not in row or temp_col not in row or pressure_col not in row:
                continue
            try:
                times.append(float(row["A"]))
                temperatures.append(float(row[temp_col]))
                pressures.append(float(row[pressure_col]))
            except ValueError:
                continue

        if times:
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
            "id": i + 1,
            "label": f"{t.level} • {t.trial}",
            "level": t.level,
            "trial": t.trial,
            "time_s": t.time_s,
            "pressure_kpa": t.pressure_kpa,
            "temperature_c": t.temperature_c,
        }
        for i, t in enumerate(trials)
    ]

    return """<!doctype html>
<html lang=\"en\">
<head>
  <meta charset=\"utf-8\" />
  <meta name=\"viewport\" content=\"width=device-width,initial-scale=1\" />
  <title>Pressure Analyzer</title>
  <script src=\"https://cdn.jsdelivr.net/npm/vue@3/dist/vue.global.prod.js\"></script>
  <script src=\"https://cdn.plot.ly/plotly-2.35.2.min.js\"></script>
  <style>
    :root { --bg:#f4f6fb; --card:#fff; --ink:#182236; --muted:#5d6980; --accent:#405cf5; --accent-soft:#eef1ff; }
    * { box-sizing:border-box; }
    body { margin:0; background:var(--bg); color:var(--ink); font-family:Inter,Segoe UI,Arial,sans-serif; }
    .wrap { max-width:1300px; margin:1.25rem auto; padding:0 1rem; }
    .grid { display:grid; grid-template-columns:320px 1fr; gap:1rem; }
    .card { background:var(--card); border-radius:12px; box-shadow:0 6px 20px rgba(18,25,38,.08); padding:1rem; }
    h1 { margin:0 0 .75rem; font-size:1.35rem; }
    .sub { color:var(--muted); font-size:.9rem; margin-bottom:1rem; }
    label { display:block; font-size:.85rem; color:var(--muted); margin:.6rem 0 .25rem; }
    select,input { width:100%; padding:.55rem .65rem; border:1px solid #d7ddea; border-radius:8px; }
    .row2 { display:grid; grid-template-columns:1fr 1fr; gap:.5rem; }
    .btnrow { display:grid; grid-template-columns:1fr 1fr; gap:.5rem; }
    button { padding:.55rem .7rem;border:none;border-radius:8px;background:var(--accent);color:white;cursor:pointer; margin-top:.5rem; width:100%; }
    #plot { height:620px; }
    .bottom-panel { margin-top:.75rem; border:1px solid #dbe3f6; border-radius:10px; padding:.8rem; background:var(--accent-soft); }
    .bottom-title { font-weight:600; margin-bottom:.45rem; }
    .bottom-grid { display:grid; grid-template-columns:repeat(2,minmax(220px,1fr)); gap:.45rem .7rem; font-size:.9rem; }
    .muted { color:var(--muted); }
    @media (max-width: 1000px) { .grid { grid-template-columns:1fr; } #plot { height:520px; } .bottom-grid { grid-template-columns:1fr; } }
  </style>
</head>
<body>
<div id=\"app\" class=\"wrap\">
  <div class=\"grid\">
    <section class=\"card\">
      <h1>Pressure Rise Explorer</h1>
      <div class=\"sub\">Drag horizontally to select a time window. All points in that time range are included.</div>
      <label>Temperature level</label>
      <select v-model=\"selectedLevel\">
        <option v-for=\"lv in levels\" :key=\"lv\" :value=\"lv\">{{ lv }}</option>
      </select>
      <label>Trial</label>
      <select v-model.number=\"selectedTrialId\">
        <option v-for=\"t in filteredTrials\" :key=\"t.id\" :value=\"t.id\">{{ t.label }}</option>
      </select>
      <label>Manual range (seconds)</label>
      <div class=\"row2\">
        <input type=\"number\" step=\"0.2\" v-model.number=\"manualStart\" />
        <input type=\"number\" step=\"0.2\" v-model.number=\"manualEnd\" />
      </div>
      <button @click=\"applyManualRange\">Calculate from range</button>
      <div class=\"btnrow\">
        <button @click=\"copyWindowResults\">Copy window results</button>
        <button @click=\"downloadCsv\">Download CSV</button>
      </div>
      <div class=\"sub\" style=\"margin-top:1rem\">Headspace volume is fixed at 225 mL. Use equation: Δn = (V_gas / RT) × (ΔP/Δt).</div>
    </section>
    <section class=\"card\">
      <div id=\"plot\"></div>
      <div class=\"bottom-panel\">
        <div class=\"bottom-title\">Selected window stats</div>
        <div class=\"bottom-grid\">
          <div><strong>Start time:</strong> {{ stats.startTime }}</div>
          <div><strong>End time:</strong> {{ stats.endTime }}</div>
          <div><strong>ΔP/Δt:</strong> {{ stats.dpdt }}</div>
          <div><strong>Δt:</strong> {{ stats.deltaT }}</div>
          <div><strong>ΔP:</strong> {{ stats.deltaP }}</div>
          <div><strong>Pressure range:</strong> {{ stats.pressureRange }}</div>
          <div><strong>ΔT:</strong> {{ stats.deltaTemp }}</div>
          <div><strong>Temperature range:</strong> {{ stats.tempRange }}</div>
          <div><strong>Equation:</strong> {{ stats.equation }}</div>
          <div><strong>Slope uncertainty (SE):</strong> {{ stats.slopeSE }}</div>
          <div><strong>Relative slope uncertainty:</strong> {{ stats.slopeRelUnc }}</div>
          <div class=\"muted\">Use drag or manual range to update values.</div>
        </div>
      </div>
    </section>
  </div>
</div>
<script>
const trialData = __PAYLOAD__;

const blankStats = () => ({
  startTime: '—',
  endTime: '—',
  dpdt: '—',
  deltaT: '—',
  deltaP: '—',
  pressureRange: '—',
  deltaTemp: '—',
  tempRange: '—',
  equation: '—',
  slopeSE: '—',
  slopeRelUnc: '—'
});

const App = {
  data() {
    const levels = [...new Set(trialData.map(t => t.level))];
    const firstLevel = levels[0];
    const firstTrial = trialData.find(t => t.level === firstLevel)?.id || trialData[0]?.id || 1;
    return {
      trials: trialData,
      levels,
      selectedLevel: firstLevel,
      selectedTrialId: firstTrial,
      manualStart: 0,
      manualEnd: 10,
      stats: blankStats(),
      selectedIndices: [],
      latestExportRow: null
    };
  },
  computed: {
    filteredTrials() {
      return this.trials.filter(t => t.level === this.selectedLevel);
    },
    activeTrial() {
      return this.trials.find(t => t.id === this.selectedTrialId) || this.filteredTrials[0];
    }
  },
  watch: {
    selectedLevel() {
      if (!this.filteredTrials.some(t => t.id === this.selectedTrialId)) this.selectedTrialId = this.filteredTrials[0]?.id;
      this.$nextTick(this.drawPlot);
    },
    selectedTrialId() {
      this.$nextTick(this.drawPlot);
    }
  },
  methods: {
    formatNumber(v, d=3) {
      return Number.isFinite(v) ? v.toFixed(d) : 'N/A';
    },
    linearRegression(x, y) {
      const n = x.length;
      if (n < 2) return { slope: NaN, intercept: NaN, r2: NaN, slope_stderr: NaN, n_points: n, residual_std: NaN };
      const xMean = x.reduce((a,b)=>a+b,0)/n;
      const yMean = y.reduce((a,b)=>a+b,0)/n;
      let sxx = 0;
      let sxy = 0;
      let sst = 0;
      for (let i=0; i<n; i++) {
        const dx = x[i]-xMean;
        const dy = y[i]-yMean;
        sxx += dx*dx;
        sxy += dx*dy;
        sst += dy*dy;
      }
      if (sxx === 0) return { slope: NaN, intercept: NaN, r2: NaN, slope_stderr: NaN, n_points: n, residual_std: NaN };
      const slope = sxy / sxx;
      const intercept = yMean - slope*xMean;
      let sse = 0;
      for (let i=0; i<n; i++) {
        const yHat = slope*x[i] + intercept;
        const e = y[i] - yHat;
        sse += e*e;
      }
      const r2 = sst === 0 ? NaN : 1 - (sse / sst);
      const residualStd = n > 2 ? Math.sqrt(sse / (n - 2)) : NaN;
      const slopeStderr = n > 2 ? residualStd / Math.sqrt(sxx) : NaN;
      return { slope, intercept, r2, slope_stderr: slopeStderr, n_points: n, residual_std: residualStd };
    },
    getIndicesForRange(start, end) {
      const t = this.activeTrial;
      if (!t) return [];
      const lo = Math.min(start, end);
      const hi = Math.max(start, end);
      const indices = [];
      for (let i=0; i<t.time_s.length; i++) {
        if (t.time_s[i] >= lo && t.time_s[i] <= hi) indices.push(i);
      }
      return indices;
    },
    applySelection(indices) {
      const t = this.activeTrial;
      if (!t || indices.length === 0) return;
      const xs = indices.map(i => t.time_s[i]);
      const ys = indices.map(i => t.pressure_kpa[i]);
      const ts = indices.map(i => t.temperature_c[i]);
      const x0 = xs[0];
      const x1 = xs[xs.length - 1];
      const p0 = ys[0];
      const p1 = ys[ys.length - 1];
      const t0 = ts[0];
      const t1 = ts[ts.length - 1];
      const reg = this.linearRegression(xs, ys);
      const slopeRelUnc = Number.isFinite(reg.slope_stderr) && Number.isFinite(reg.slope) && reg.slope !== 0
        ? Math.abs(reg.slope_stderr / reg.slope)
        : NaN;
      this.selectedIndices = indices;
      this.stats = {
        startTime: `${this.formatNumber(x0, 2)} s`,
        endTime: `${this.formatNumber(x1, 2)} s`,
        dpdt: (x1 - x0) !== 0 ? `${this.formatNumber((p1 - p0) / (x1 - x0), 5)} kPa/s` : 'N/A',
        deltaT: `${this.formatNumber(x1 - x0, 2)} s`,
        deltaP: `${this.formatNumber(p1 - p0, 3)} kPa`,
        pressureRange: `${this.formatNumber(p0, 3)} → ${this.formatNumber(p1, 3)} kPa`,
        deltaTemp: `${this.formatNumber(t1 - t0, 3)} °C`,
        tempRange: `${this.formatNumber(t0, 3)} → ${this.formatNumber(t1, 3)} °C`,
        equation: Number.isFinite(reg.slope) ? `P = ${reg.slope.toFixed(5)}·t + ${reg.intercept.toFixed(5)}` : 'N/A',
        slopeSE: `${this.formatNumber(reg.slope_stderr, 6)} kPa/s`,
        slopeRelUnc: Number.isFinite(slopeRelUnc) ? `${this.formatNumber(slopeRelUnc * 100, 3)} %` : 'N/A'
      };
      this.latestExportRow = {
        level: t.level,
        trial: t.trial,
        startTime_s: x0,
        endTime_s: x1,
        slope_kPa_s: reg.slope,
        slopeSE_kPa_s: reg.slope_stderr,
        slope_rel_unc: slopeRelUnc,
        r2: reg.r2,
        n_points: reg.n_points,
        residual_std_kPa: reg.residual_std,
        deltaP_kPa: p1 - p0,
        deltaT_s: x1 - x0
      };
      Plotly.restyle('plot', { selectedpoints: [indices] }, [0]);
      Plotly.relayout('plot', {
        shapes: [{
          type: 'rect', xref: 'x', yref: 'paper',
          x0: Math.min(x0, x1), x1: Math.max(x0, x1), y0: 0, y1: 1,
          fillcolor: 'rgba(64, 92, 245, 0.09)',
          line: { width: 0 }
        }]
      });
    },
    applyManualRange() {
      const indices = this.getIndicesForRange(this.manualStart, this.manualEnd);
      this.applySelection(indices);
    },
    copyWindowResults() {
      if (!this.latestExportRow) return;
      const r = this.latestExportRow;
      const line = [
        r.level,
        r.trial,
        this.formatNumber(r.startTime_s, 4),
        this.formatNumber(r.endTime_s, 4),
        this.formatNumber(r.slope_kPa_s, 8),
        this.formatNumber(r.slopeSE_kPa_s, 8),
        this.formatNumber(r.slope_rel_unc, 8),
        this.formatNumber(r.r2, 6),
        r.n_points,
        this.formatNumber(r.residual_std_kPa, 8),
        this.formatNumber(r.deltaP_kPa, 6),
        this.formatNumber(r.deltaT_s, 6)
      ].map(v => String(v).replaceAll(',', ';')).join(',');
      navigator.clipboard.writeText(line);
    },
    downloadCsv() {
      if (!this.latestExportRow) return;
      const r = this.latestExportRow;
      const headers = [
        'level','trial','startTime_s','endTime_s','slope_kPa_s','slopeSE_kPa_s','slope_rel_unc','r2','n_points','residual_std_kPa','deltaP_kPa','deltaT_s'
      ];
      const row = [
        r.level,
        r.trial,
        this.formatNumber(r.startTime_s, 6),
        this.formatNumber(r.endTime_s, 6),
        this.formatNumber(r.slope_kPa_s, 10),
        this.formatNumber(r.slopeSE_kPa_s, 10),
        this.formatNumber(r.slope_rel_unc, 10),
        this.formatNumber(r.r2, 10),
        r.n_points,
        this.formatNumber(r.residual_std_kPa, 10),
        this.formatNumber(r.deltaP_kPa, 10),
        this.formatNumber(r.deltaT_s, 10)
      ].map(v => `"${String(v).replaceAll('"','""')}"`);
      const csv = `${headers.join(',')}\n${row.join(',')}\n`;
      const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'window_results.csv';
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    },
    drawPlot() {
      const t = this.activeTrial;
      if (!t) return;
      const trace = {
        x: t.time_s,
        y: t.pressure_kpa,
        type: 'scatter',
        mode: 'lines+markers',
        marker: { size: 6, color: '#405cf5' },
        line: { width: 2, color: '#405cf5' },
        customdata: t.temperature_c,
        selected: { marker: { color: '#0d1fb5', size: 7 } },
        unselected: { marker: { opacity: 0.35 } },
        hovertemplate: 'Time: %{x:.2f} s<br>Pressure: %{y:.3f} kPa<br>Temperature: %{customdata:.2f} °C<extra></extra>'
      };
      const layout = {
        title: `${t.level} — ${t.trial}`,
        dragmode: 'select',
        selectdirection: 'h',
        hovermode: 'closest',
        margin: { l: 65, r: 20, t: 50, b: 55 },
        xaxis: { title: 'Time (s)', showgrid: true, gridcolor: '#e9edf6' },
        yaxis: { title: 'Pressure (kPa)', showgrid: true, gridcolor: '#e9edf6' },
        paper_bgcolor: '#ffffff',
        plot_bgcolor: '#ffffff',
        shapes: []
      };
      Plotly.newPlot('plot', [trace], layout, { responsive: true });
      const plotEl = document.getElementById('plot');
      plotEl.on('plotly_selected', (eventData) => {
        if (!eventData || !eventData.range || !eventData.range.x) return;
        const x0 = Math.min(eventData.range.x[0], eventData.range.x[1]);
        const x1 = Math.max(eventData.range.x[0], eventData.range.x[1]);
        const indices = this.getIndicesForRange(x0, x1);
        this.applySelection(indices);
      });
      this.stats = blankStats();
      this.selectedIndices = [];
      this.latestExportRow = null;
    }
  },
  mounted() {
    this.drawPlot();
  }
};

Vue.createApp(App).mount('#app');
</script>
</body>
</html>
""".replace("__PAYLOAD__", json.dumps(payload))


def run_server(html_content: str, host: str, port: int, open_browser: bool) -> None:
    with TemporaryDirectory() as tmp_dir:
        root = Path(tmp_dir)
        (root / "index.html").write_text(html_content, encoding="utf-8")

        class Handler(SimpleHTTPRequestHandler):
            def __init__(self, *args: Any, **kwargs: Any):
                super().__init__(*args, directory=str(root), **kwargs)

        server = ThreadingHTTPServer((host, port), Handler)
        display_host = "127.0.0.1" if host == "0.0.0.0" else host
        url = f"http://{display_host}:{port}"
        print(f"Serving interactive analyzer at: {url}")
        print("Press Ctrl+C to stop.")

        if open_browser:
            threading.Timer(0.7, lambda: webbrowser.open(url)).start()

        try:
            server.serve_forever()
        except KeyboardInterrupt:
            print("\nServer stopped.")
        finally:
            server.server_close()


def main() -> None:
    parser = argparse.ArgumentParser(description="Interactive pressure-vs-time analyzer for Raw Data sheet")
    parser.add_argument("--xlsx", default="Data/Data.xlsx", help="Path to workbook")
    parser.add_argument("--host", default="0.0.0.0", help="Server host/interface")
    parser.add_argument("--port", type=int, default=8050, help="Server port")
    parser.add_argument("--no-browser", action="store_true", help="Disable browser auto-open")
    args = parser.parse_args()

    workbook = Path(args.xlsx)
    if not workbook.exists():
        raise SystemExit(f"Workbook not found: {workbook}")

    trials = parse_raw_data_trials(workbook)
    if not trials:
        raise SystemExit("No trial data found in Raw Data sheet.")

    run_server(build_html(trials), host=args.host, port=args.port, open_browser=not args.no_browser)


if __name__ == "__main__":
    main()
