# Interactive Pressure-vs-Time Analyzer

## Overview
A Python web application that reads experimental data from `Data/Data.xlsx` and serves an interactive Plotly-based chart for analyzing pressure vs time across 40 trials (8 temperature levels x 5 trials). Users can select trials and drag-select time ranges to compute pressure rise rates.

## Project Architecture
- **Language**: Python 3.12 (standard library only, no external dependencies)
- **Entry point**: `interactive_pressure_analysis.py`
- **Data**: `Data/Data.xlsx` (Excel workbook with "Raw Data" sheet)
- **Frontend**: Single-page HTML served by Python's built-in HTTP server, uses Plotly.js via CDN
- **Port**: 5000 (bound to 0.0.0.0)

## Key Files
- `interactive_pressure_analysis.py` - Main application: parses Excel, builds HTML, serves web app
- `Data/Data.xlsx` - Source experimental data
- `Documents/` - LaTeX report and related resources
- `Resources/` - IA guidance documents

## Recent Changes
- 2026-02-07: Configured for Replit environment (port 5000, cache-control headers)
