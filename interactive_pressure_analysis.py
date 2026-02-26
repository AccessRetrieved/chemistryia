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

    return f"""<!doctype html>
<html lang=\"en\">
<head>
  <meta charset=\"utf-8\" />
  <meta name=\"viewport\" content=\"width=device-width,initial-scale=1\" />
  <title>Pressure Analyzer</title>
  <script src=\"https://cdn.jsdelivr.net/npm/vue@3/dist/vue.global.prod.js\"></script>
  <script src=\"https://cdn.plot.ly/plotly-2.35.2.min.js\"></script>
  <style>
    :root {{ --bg:#f4f6fb; --card:#fff; --ink:#182236; --muted:#5d6980; --accent:#405cf5; --accent-soft:#eef1ff; --warn:#a15a00; }}
    * {{ box-sizing:border-box; }}
    body {{ margin:0; background:var(--bg); color:var(--ink); font-family:Inter,Segoe UI,Arial,sans-serif; }}
    .wrap {{ max-width:1300px; margin:1.25rem auto; padding:0 1rem; }}
    .grid {{ display:grid; grid-template-columns:320px 1fr; gap:1rem; }}
    .card {{ background:var(--card); border-radius:12px; box-shadow:0 6px 20px rgba(18,25,38,.08); padding:1rem; }}
    h1 {{ margin:0 0 .75rem; font-size:1.35rem; }}
    .sub {{ color:var(--muted); font-size:.9rem; margin-bottom:1rem; }}
    .warn {{ color:var(--warn); font-size:.85rem; margin-top:.45rem; }}
    label {{ display:block; font-size:.85rem; color:var(--muted); margin:.6rem 0 .25rem; }}
    select,input {{ width:100%; padding:.55rem .65rem; border:1px solid #d7ddea; border-radius:8px; }}
    .row2 {{ display:grid; grid-template-columns:1fr 1fr; gap:.5rem; }}
    .btnrow {{ display:grid; grid-template-columns:1fr 1fr; gap:.5rem; margin-top:.5rem; }}
    button {{ padding:.55rem .7rem;border:none;border-radius:8px;background:var(--accent);color:white;cursor:pointer; }}
    .ghost {{ background:#667084; }}
    .metric {{ border:1px solid #e8edf7; border-radius:10px; padding:.55rem .65rem; margin-top:.5rem; }}
    .metric b {{ display:block; font-size:1.05rem; word-break:break-word; }}
    .metric span {{ font-size:.8rem; color:var(--muted); }}
    #plot {{ height:620px; }}
    .bottom-panel {{ margin-top:.75rem; border:1px solid #dbe3f6; border-radius:10px; padding:.8rem; background:var(--accent-soft); }}
    .bottom-title {{ font-weight:600; margin-bottom:.45rem; }}
    .bottom-grid {{ display:grid; grid-template-columns:repeat(2,minmax(220px,1fr)); gap:.45rem .7rem; font-size:.9rem; }}
    .muted {{ color:var(--muted); }}
    @media (max-width: 1000px) {{ .grid {{ grid-template-columns:1fr; }} #plot {{ height:520px; }} .bottom-grid {{ grid-template-columns:1fr; }} }}
  </style>
</head>
<body>
<div id=\"app\" class=\"wrap\">
  <div class=\"grid\">
    <section class=\"card\">
      <h1>Pressure Rise Explorer</h1>
      <div class=\"sub\">Vue + Plotly interactive analysis for Raw Data trials.</div>

      <label>Temperature level</label>
      <select v-model=\"selectedLevel\">
        <option v-for=\"lv in levels\" :key=\"lv\" :value=\"lv\">{{{{ lv }}}}</option>
      </select>

      <label>Trial</label>
      <select v-model.number=\"selectedTrialId\">
        <option v-for=\"t in filteredTrials\" :key=\"t.id\" :value=\"t.id\">{{{{ t.label }}}}</option>
      </select>

      <label>Manual range (seconds)</label>
      <div class=\"row2\">
        <input type=\"number\" step=\"0.2\" v-model.number=\"manualStart\" />
        <input type=\"number\" step=\"0.2\" v-model.number=\"manualEnd\" />
      </div>
      <button @click=\"applyManualRange\" style=\"margin-top:.5rem;\">Calculate from range</button>

      <label>Headspace gas volume V_gas (mL)</label>
      <div class=\"row2\">
        <input type=\"number\" step=\"0.1\" v-model.number=\"gasVolumeMl\" />
        <input type=\"number\" step=\"0.1\" v-model.number=\"gasVolumeUncMl\" placeholder=\"Uncertainty (mL)\" />
      </div>
      <div v-if=\"!isPositive(gasVolumeMl)\" class=\"warn\">Enter V_gas > 0 mL to compute molar rate.</div>

      <label>Click-window length (s)</label>
      <input type=\"number\" step=\"0.5\" min=\"0.5\" v-model.number=\"clickWindowSeconds\" />

      <label>Temperature setpoint T_set (°C)</label>
      <input type=\"number\" step=\"0.1\" v-model.number=\"tempSetC\" />

      <label>Temperature uncertainty ΔT (°C)</label>
      <div class=\"row2\">
        <input type=\"number\" step=\"0.01\" v-model.number=\"tempUncC\" />
        <button @click=\"useAutoTempUnc\" class=\"ghost\">Use auto</button>
      </div>

      <div class=\"btnrow\">
        <button @click=\"copyWindowResults\">Copy window results</button>
        <button @click=\"downloadCsv\" class=\"ghost\">Download CSV</button>
      </div>

      <div class=\"metric\"><b>{{{{ stats.points }}}}</b><span>Selected points</span></div>
      <div class=\"metric\"><b>{{{{ stats.timeRange }}}}</b><span>Time range</span></div>
      <div class=\"metric\"><b>{{{{ stats.pressureRange }}}}</b><span>Pressure range</span></div>
      <div class=\"metric\"><b>{{{{ stats.slope }}}} ± {{{{ stats.slopeStdErr }}}} kPa/s</b><span>dP/dt from linear fit</span></div>
      <div class=\"metric\"><b>{{{{ stats.r0mol }}}} ± {{{{ stats.r0uncMol }}}} mol/s</b><span>Initial rate r0 (ideal gas law)</span></div>
      <div class=\"metric\"><b>{{{{ stats.r0mmol }}}} ± {{{{ stats.r0uncMmol }}}} mmol/s</b><span>Initial rate r0 (readable units)</span></div>
      <p class=\"sub\" style=\"margin-top:.7rem\">Tip: choose Box Select in chart toolbar and drag over points.</p>
    </section>

    <section class=\"card\">
      <div id=\"plot\"></div>
      <div class=\"bottom-panel\">
        <div class=\"bottom-title\">Window stats</div>
        <div class=\"bottom-grid\">
          <div><strong>Mode:</strong> {{{{ clickStats.mode }}}}</div>
          <div><strong>Start time:</strong> {{{{ clickStats.startTime }}}}</div>
          <div><strong>End time:</strong> {{{{ clickStats.endTime }}}}</div>
          <div><strong>Points:</strong> {{{{ clickStats.points }}}}</div>
          <div><strong>Δt:</strong> {{{{ clickStats.deltaT }}}}</div>
          <div><strong>Pressure range:</strong> {{{{ clickStats.pressureRange }}}}</div>
          <div><strong>ΔP:</strong> {{{{ clickStats.deltaP }}}}</div>
          <div><strong>Fit slope:</strong> {{{{ clickStats.slope }}}} ± {{{{ clickStats.slopeStdErr }}}} kPa/s</div>
          <div><strong>Fit equation:</strong> {{{{ clickStats.fit }}}}</div>
          <div><strong>T_set:</strong> {{{{ clickStats.tSet }}}}</div>
          <div><strong>T range:</strong> {{{{ clickStats.tempRange }}}}</div>
          <div><strong>ΔT used:</strong> {{{{ clickStats.tempUnc }}}}</div>
          <div><strong>r0:</strong> {{{{ clickStats.r0mol }}}} ± {{{{ clickStats.r0uncMol }}}} mol/s</div>
          <div><strong>r0:</strong> {{{{ clickStats.r0mmol }}}} ± {{{{ clickStats.r0uncMmol }}}} mmol/s</div>
          <div><strong>ln(r0):</strong> {{{{ clickStats.lnR0 }}}}</div>
          <div><strong>1/T:</strong> {{{{ clickStats.invT }}}} K^-1</div>
          <div class=\"muted\">Window length setting: {{{{ formatNumber(clickWindowSeconds, 2) }}}} s</div>
        </div>
      </div>
    </section>
  </div>
</div>

<script>
const trialData = {json.dumps(payload)};
const R_GAS = 8.314462618;

const blankStats = () => ({{
  points: 0,
  timeRange: '—',
  pressureRange: '—',
  slope: '—',
  slopeStdErr: '—',
  r0mol: '—',
  r0uncMol: '—',
  r0mmol: '—',
  r0uncMmol: '—'
}});

const blankClickStats = () => ({{
  mode: '—',
  startTime: '—',
  endTime: '—',
  points: 0,
  deltaT: '—',
  pressureRange: '—',
  deltaP: '—',
  slope: '—',
  slopeStdErr: '—',
  fit: '—',
  tSet: '—',
  tempRange: '—',
  tempUnc: '—',
  r0mol: '—',
  r0uncMol: '—',
  r0mmol: '—',
  r0uncMmol: '—',
  lnR0: '—',
  invT: '—'
}});

const App = {{
  data() {{
    const levels = [...new Set(trialData.map(t => t.level))];
    const firstLevel = levels[0];
    const firstTrial = trialData.find(t => t.level === firstLevel)?.id || trialData[0]?.id || 1;
    const initSetpoint = this.extractTempSetpoint(firstLevel);
    return {{
      trials: trialData,
      levels,
      selectedLevel: firstLevel,
      selectedTrialId: firstTrial,
      manualStart: 0,
      manualEnd: 10,
      gasVolumeMl: 0,
      gasVolumeUncMl: 0,
      clickWindowSeconds: 15,
      tempSetC: initSetpoint,
      tempUncC: 0,
      tempUncOverride: false,
      stats: blankStats(),
      clickStats: blankClickStats(),
      latestWindow: null,
      latestCsvRow: null,
      copying: false
    }};
  }},
  computed: {{
    filteredTrials() {{
      return this.trials.filter(t => t.level === this.selectedLevel);
    }},
    activeTrial() {{
      return this.trials.find(t => t.id === this.selectedTrialId) || this.filteredTrials[0];
    }}
  }},
  watch: {{
    selectedLevel() {{
      if (!this.filteredTrials.some(t => t.id === this.selectedTrialId)) this.selectedTrialId = this.filteredTrials[0]?.id;
      const guessed = this.extractTempSetpoint(this.selectedLevel);
      this.tempSetC = guessed;
      this.$nextTick(this.drawPlot);
    }},
    selectedTrialId() {{
      this.$nextTick(this.drawPlot);
    }},
    clickWindowSeconds() {{
      if (this.latestWindow?.mode === 'click') this.recomputeLatestWindow();
    }},
    gasVolumeMl() {{ this.recomputeLatestWindow(); }},
    gasVolumeUncMl() {{ this.recomputeLatestWindow(); }},
    tempSetC() {{ this.recomputeLatestWindow(); }},
    tempUncC() {{
      this.tempUncOverride = true;
      this.recomputeLatestWindow();
    }}
  }},
  methods: {{
    isPositive(v) {{ return Number.isFinite(v) && v > 0; }},
    formatNumber(v, d=5) {{ return Number.isFinite(v) ? v.toFixed(d) : 'N/A'; }},
    extractTempSetpoint(levelText) {{
      const m = String(levelText || '').match(/-?\d+(\.\d+)?/);
      return m ? Number(m[0]) : NaN;
    }},
    useAutoTempUnc() {{
      this.tempUncOverride = false;
      this.recomputeLatestWindow();
    }},
    linearRegression(x, y) {{
      const n = x.length;
      if (n < 2) return {{ slope: NaN, intercept: NaN, slopeStdErr: NaN, r2: NaN }};
      const xMean = x.reduce((a,b)=>a+b,0)/n;
      const yMean = y.reduce((a,b)=>a+b,0)/n;
      let sxx = 0;
      let sxy = 0;
      let syy = 0;
      for (let i=0; i<n; i++) {{
        const dx = x[i]-xMean;
        const dy = y[i]-yMean;
        sxx += dx*dx;
        sxy += dx*dy;
        syy += dy*dy;
      }}
      if (sxx === 0) return {{ slope: NaN, intercept: NaN, slopeStdErr: NaN, r2: NaN }};
      const slope = sxy / sxx;
      const intercept = yMean - slope*xMean;
      let sse = 0;
      for (let i=0; i<n; i++) {{
        const yHat = slope*x[i] + intercept;
        const e = y[i] - yHat;
        sse += e*e;
      }}
      const syx = n > 2 ? Math.sqrt(sse / (n - 2)) : NaN;
      const slopeStdErr = n > 2 ? syx / Math.sqrt(sxx) : NaN;
      const r2 = syy === 0 ? 1 : 1 - (sse / syy);
      return {{ slope, intercept, slopeStdErr, r2 }};
    }},
    getIndicesForManualRange() {{
      const t = this.activeTrial;
      if (!t) return [];
      const lo = Math.min(this.manualStart, this.manualEnd);
      const hi = Math.max(this.manualStart, this.manualEnd);
      const indices = [];
      for (let i=0; i<t.time_s.length; i++) {{
        if (t.time_s[i] >= lo && t.time_s[i] <= hi) indices.push(i);
      }}
      return indices;
    }},
    getIndicesForClick(startIndex) {{
      const t = this.activeTrial;
      if (!t) return {{ indices: [], x0: NaN, x1: NaN }};
      const x0 = t.time_s[startIndex];
      const windowS = this.isPositive(this.clickWindowSeconds) ? this.clickWindowSeconds : 15;
      const x1 = x0 + windowS;
      const indices = [];
      for (let i=startIndex; i<t.time_s.length; i++) {{
        if (t.time_s[i] <= x1) indices.push(i);
        else break;
      }}
      return {{ indices, x0, x1 }};
    }},
    computeWindowFromIndices(indices, modeLabel, x0Hint=null, x1Hint=null) {{
      const t = this.activeTrial;
      if (!t || !indices || indices.length === 0) return null;
      const xs = indices.map(i => t.time_s[i]);
      const ys = indices.map(i => t.pressure_kpa[i]);
      const temps = indices.map(i => t.temperature_c[i]);
      const startT = Number.isFinite(x0Hint) ? x0Hint : xs[0];
      const endTForStats = xs[xs.length - 1];
      const endTHighlight = Number.isFinite(x1Hint) ? x1Hint : endTForStats;
      const pMin = Math.min(...ys);
      const pMax = Math.max(...ys);
      const tMin = Math.min(...temps);
      const tMax = Math.max(...temps);
      const tHalfRange = (tMax - tMin) / 2;
      const tempUncUsedC = this.tempUncOverride && Number.isFinite(this.tempUncC) ? Math.abs(this.tempUncC) : tHalfRange;
      if (!this.tempUncOverride) this.tempUncC = tHalfRange;

      const reg = this.linearRegression(xs, ys);
      const dt = endTForStats - xs[0];
      const dp = ys[ys.length - 1] - ys[0];
      const vMl = this.gasVolumeMl;
      const vMlUnc = this.gasVolumeUncMl;
      const tSetC = this.tempSetC;
      const tSetK = Number.isFinite(tSetC) ? tSetC + 273.15 : NaN;

      const slopeKpaS = reg.slope;
      const slopeStdErrKpaS = reg.slopeStdErr;
      const slopePaS = Number.isFinite(slopeKpaS) ? slopeKpaS * 1000 : NaN;
      const vM3 = Number.isFinite(vMl) ? vMl * 1e-6 : NaN;
      const r0 = Number.isFinite(vM3) && Number.isFinite(tSetK) && tSetK > 0 && Number.isFinite(slopePaS)
        ? (vM3 / (R_GAS * tSetK)) * slopePaS
        : NaN;

      const relV = this.isPositive(vMl) && Number.isFinite(vMlUnc) ? Math.abs(vMlUnc / vMl) : NaN;
      const relSlope = Number.isFinite(slopeKpaS) && slopeKpaS !== 0 && Number.isFinite(slopeStdErrKpaS)
        ? Math.abs(slopeStdErrKpaS / slopeKpaS)
        : NaN;
      const relT = Number.isFinite(tSetK) && tSetK > 0 && Number.isFinite(tempUncUsedC)
        ? Math.abs(tempUncUsedC / tSetK)
        : NaN;

      let r0Unc = NaN;
      if (Number.isFinite(r0) && Number.isFinite(relV) && Number.isFinite(relSlope) && Number.isFinite(relT)) {{
        r0Unc = Math.abs(r0) * Math.sqrt(relV*relV + relSlope*relSlope + relT*relT);
      }}

      const lnR0 = Number.isFinite(r0) && r0 > 0 ? Math.log(r0) : NaN;
      const invT = Number.isFinite(tSetK) && tSetK > 0 ? (1 / tSetK) : NaN;

      const row = {{
        level: t.level,
        trial: t.trial,
        startTime: xs[0],
        endTime: endTForStats,
        nPoints: indices.length,
        slope_kPa_s: slopeKpaS,
        slopeStdErr_kPa_s: slopeStdErrKpaS,
        V_gas_mL: vMl,
        T_set_C: tSetC,
        T_unc_C: tempUncUsedC,
        r0_mol_s: r0,
        r0_unc_mol_s: r0Unc,
        ln_r0: lnR0,
        inv_T_K_inv: invT
      }};

      return {{
        mode: modeLabel,
        indices,
        x0: startT,
        x1: endTHighlight,
        row,
        display: {{
          mode: modeLabel,
          startTime: `${{xs[0].toFixed(2)}} s`,
          endTime: `${{endTForStats.toFixed(2)}} s`,
          points: indices.length,
          deltaT: Number.isFinite(dt) ? `${{dt.toFixed(2)}} s` : 'N/A',
          pressureRange: `${{pMin.toFixed(3)}} → ${{pMax.toFixed(3)}} kPa`,
          deltaP: Number.isFinite(dp) ? `${{dp.toFixed(3)}} kPa` : 'N/A',
          slope: this.formatNumber(slopeKpaS, 5),
          slopeStdErr: this.formatNumber(slopeStdErrKpaS, 5),
          fit: Number.isFinite(reg.slope) ? `P = ${{reg.slope.toFixed(5)}}·t + ${{reg.intercept.toFixed(5)}} (R²=${{this.formatNumber(reg.r2, 4)}})` : 'N/A',
          tSet: Number.isFinite(tSetC) ? `${{tSetC.toFixed(2)}} °C / ${{(tSetK).toFixed(2)}} K` : 'N/A',
          tempRange: `${{tMin.toFixed(2)}} → ${{tMax.toFixed(2)}} °C`,
          tempUnc: Number.isFinite(tempUncUsedC) ? `${{tempUncUsedC.toFixed(3)}} °C` : 'N/A',
          r0mol: this.formatNumber(r0, 8),
          r0uncMol: this.formatNumber(r0Unc, 8),
          r0mmol: this.formatNumber(Number.isFinite(r0) ? r0 * 1000 : NaN, 5),
          r0uncMmol: this.formatNumber(Number.isFinite(r0Unc) ? r0Unc * 1000 : NaN, 5),
          lnR0: this.formatNumber(lnR0, 6),
          invT: this.formatNumber(invT, 8)
        }}
      }};
    }},
    applyWindowResult(result) {{
      if (!result) return;
      this.latestWindow = result;
      this.latestCsvRow = result.row;
      this.clickStats = result.display;
      this.stats = {{
        points: result.display.points,
        timeRange: `${{result.display.startTime}} → ${{result.display.endTime}}`,
        pressureRange: result.display.pressureRange,
        slope: result.display.slope,
        slopeStdErr: result.display.slopeStdErr,
        r0mol: result.display.r0mol,
        r0uncMol: result.display.r0uncMol,
        r0mmol: result.display.r0mmol,
        r0uncMmol: result.display.r0uncMmol
      }};
      Plotly.restyle('plot', {{ selectedpoints: [result.indices] }}, [0]);
      Plotly.relayout('plot', {{
        shapes: [{{
          type: 'rect', xref: 'x', yref: 'paper',
          x0: result.x0, x1: result.x1, y0: 0, y1: 1,
          fillcolor: 'rgba(64, 92, 245, 0.09)',
          line: {{ width: 0 }}
        }}]
      }});
    }},
    recomputeLatestWindow() {{
      if (!this.latestWindow) return;
      if (this.latestWindow.mode === 'manual') {{
        const indices = this.getIndicesForManualRange();
        this.applyWindowResult(this.computeWindowFromIndices(indices, 'manual'));
      }} else if (this.latestWindow.mode === 'click') {{
        const first = this.latestWindow.indices?.[0];
        if (first === undefined) return;
        const m = this.getIndicesForClick(first);
        this.applyWindowResult(this.computeWindowFromIndices(m.indices, 'click', m.x0, m.x1));
      }} else if (this.latestWindow.mode === 'box') {{
        const indices = this.latestWindow.indices || [];
        this.applyWindowResult(this.computeWindowFromIndices(indices, 'box'));
      }}
    }},
    applyManualRange() {{
      const indices = this.getIndicesForManualRange();
      this.applyWindowResult(this.computeWindowFromIndices(indices, 'manual'));
    }},
    copyWindowResults() {{
      if (!this.latestCsvRow) return;
      const r = this.latestCsvRow;
      const fields = [
        r.level,
        r.trial,
        this.formatNumber(r.startTime, 3),
        this.formatNumber(r.endTime, 3),
        r.nPoints,
        this.formatNumber(r.slope_kPa_s, 8),
        this.formatNumber(r.slopeStdErr_kPa_s, 8),
        this.formatNumber(r.V_gas_mL, 4),
        this.formatNumber(r.T_set_C, 4),
        this.formatNumber(r.T_unc_C, 4),
        this.formatNumber(r.r0_mol_s, 10),
        this.formatNumber(r.r0_unc_mol_s, 10),
        Number.isFinite(r.ln_r0) ? r.ln_r0.toFixed(10) : 'N/A',
        this.formatNumber(r.inv_T_K_inv, 10)
      ].map(v => String(v).replaceAll(',', ';'));
      const line = fields.join(',');
      navigator.clipboard.writeText(line);
    }},
    downloadCsv() {{
      if (!this.latestCsvRow) return;
      const r = this.latestCsvRow;
      const headers = [
        'level','trial','startTime','endTime','nPoints','slope_kPa_s','slopeStdErr_kPa_s','V_gas_mL','T_set_C','T_unc_C','r0_mol_s','r0_unc_mol_s','ln(r0)','1/T(K^-1)'
      ];
      const row = [
        r.level,
        r.trial,
        this.formatNumber(r.startTime, 6),
        this.formatNumber(r.endTime, 6),
        r.nPoints,
        this.formatNumber(r.slope_kPa_s, 10),
        this.formatNumber(r.slopeStdErr_kPa_s, 10),
        this.formatNumber(r.V_gas_mL, 6),
        this.formatNumber(r.T_set_C, 6),
        this.formatNumber(r.T_unc_C, 6),
        this.formatNumber(r.r0_mol_s, 12),
        this.formatNumber(r.r0_unc_mol_s, 12),
        Number.isFinite(r.ln_r0) ? r.ln_r0.toFixed(12) : 'N/A',
        this.formatNumber(r.inv_T_K_inv, 12)
      ].map(v => `"${{String(v).replaceAll('"','""')}}"`);
      const csv = `${{headers.join(',')}}\n${{row.join(',')}}\n`;
      const blob = new Blob([csv], {{ type: 'text/csv;charset=utf-8;' }});
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'window_results.csv';
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    }},
    drawPlot() {{
      const t = this.activeTrial;
      if (!t) return;
      const trace = {{
        x: t.time_s,
        y: t.pressure_kpa,
        type: 'scatter',
        mode: 'lines+markers',
        marker: {{ size: 6, color: '#405cf5' }},
        line: {{ width: 2, color: '#405cf5' }},
        customdata: t.temperature_c,
        selected: {{ marker: {{ color: '#0d1fb5', size: 7 }} }},
        unselected: {{ marker: {{ opacity: 0.35 }} }},
        hovertemplate: 'Time: %{{x:.2f}} s<br>Pressure: %{{y:.3f}} kPa<br>Temperature: %{{customdata:.2f}} °C<extra></extra>'
      }};
      const layout = {{
        title: `${{t.level}} — ${{t.trial}}`,
        dragmode: 'select',
        hovermode: 'closest',
        margin: {{ l: 65, r: 20, t: 50, b: 55 }},
        xaxis: {{ title: 'Time (s)', showgrid: true, gridcolor: '#e9edf6' }},
        yaxis: {{ title: 'Pressure (kPa)', showgrid: true, gridcolor: '#e9edf6' }},
        paper_bgcolor: '#ffffff',
        plot_bgcolor: '#ffffff',
        shapes: []
      }};
      Plotly.newPlot('plot', [trace], layout, {{ responsive: true }});
      const plotEl = document.getElementById('plot');

      plotEl.on('plotly_selected', (eventData) => {{
        if (!eventData || !eventData.points || eventData.points.length === 0) return;
        const indices = [...new Set(eventData.points.map(p => p.pointNumber))].sort((a,b)=>a-b);
        this.applyWindowResult(this.computeWindowFromIndices(indices, 'box'));
      }});

      plotEl.on('plotly_click', (eventData) => {{
        const point = eventData?.points?.[0];
        if (!point || point.pointNumber === undefined) return;
        const m = this.getIndicesForClick(point.pointNumber);
        this.applyWindowResult(this.computeWindowFromIndices(m.indices, 'click', m.x0, m.x1));
      }});

      this.stats = blankStats();
      this.clickStats = blankClickStats();
      this.latestWindow = null;
      this.latestCsvRow = null;
      this.tempUncOverride = false;
      this.tempUncC = 0;
    }}
  }},
  mounted() {{
    this.drawPlot();
  }}
}};
Vue.createApp(App).mount('#app');
</script>
</body>
</html>
"""


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
