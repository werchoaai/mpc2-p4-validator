# MPC² Corrosion Ray — DL-EPR Legacy Validator (Project 4)

Streamlit app that parses DL-EPR ASC files using the proven A1 parser core
and exports in the **legacy format** used in Project 4 reports (overview
workbook, detail workbook, raw workbook with `Vorlage` sheet).

## Relation to A1
- **Parser & DL-EPR analysis**: 100% reused from `werchoaai/mpc2-parser`
- **UI, charts, integrity score, password gate**: 100% reused from A1
- **Excel output schema**: adapted to match legacy ON2024-0013 format

## Deploy
- Streamlit Cloud: https://mpc2-p4-validator.streamlit.app
- Embedded at: https://customer.werchota.ai/mpc2/p4-validator
- Secrets: `password = "mpc2demo2026"` (set in Streamlit Cloud Settings → Secrets)

## Standards applied
- `runtime.txt` pinning Python 3.11
- Loose dependency pins (`numpy>=2.1`, `scipy>=1.14`)
- Daily keep-alive via GitHub Actions (`.github/workflows/keep-alive.yml`)

