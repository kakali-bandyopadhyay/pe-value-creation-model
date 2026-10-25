# PE Value Creation Model

Open `model/PE_Value_Creation_Model_synthetic.xlsx`, start on **Start Here**, then review the **Dashboard**. Use the scenario selector and initiative switches on **Controls**. The forecast recalculates from the selected assumptions. Blue-font cells with pale yellow fill are editable. Black-font formulas are on the same sheet; green-font formulas link to other tabs.

> **Synthetic case:** NimbleBridge Services is an invented company. Every company figure, plan, sample actual, threshold, role and outcome in this project is synthetic or illustrative. The workbook does not describe a real deal and does not claim realized value.

## Headline returns

| Case | Entry equity ($m) | Exit EBITDA ($m) | Exit multiple | Exit equity ($m) | MOIC | IRR |
|---|---:|---:|---:|---:|---:|---:|
| Base | 86.80 | 22.76 | 10.0x | 210.09 | 2.4204x | 19.33% |
| Downside | 86.80 | 17.24 | 8.5x | 116.76 | 1.3451x | 6.11% |
| Upside | 86.80 | 25.67 | 11.0x | 271.72 | 3.1305x | 25.62% |

![Synthetic dashboard preview](docs/dashboard_preview.png)

## What the workbook contains

- Three synthetic historical years and a five-year ownership forecast.
- Base, Downside and Upside assumptions that vary operating growth, margin improvement, initiative timing, implementation cost, realization, adoption ramp and exit multiple.
- Three toggleable levers: pricing and mix after a churn offset, commercial efficiency, and AI-enabled finance automation. The AI lever calculates capacity released first and converts only the explicitly avoidable and realized share into EBITDA.
- A cash flow and debt schedule with cash tax, interest on opening debt, capital expenditure, working capital, mandatory amortization, cash sweep and retained cash.
- Dated sponsor cash flows with MOIC and XIRR formulas.
- Formula-based sensitivities, a quarterly plan tracker with synthetic actuals, visible RAG thresholds, an AI expand/fix/pause gate, and model checks.

## Rebuild and verify

Requires Python 3.11 or later, `openpyxl`, `pytest`, and LibreOffice for recalculation.

```powershell
python -m pip install -r requirements-dev.txt
python build/build_model.py
python -m pytest -q
python check/verify_returns.py
```

The build script regenerates the workbook with live formulas. The verification script builds isolated scenario workbooks from the same generator, recalculates each with LibreOffice headless, independently recomputes sponsor returns from the workbook's assumptions using Python standard-library math, checks toggle-off behavior, scans for formula errors, and refreshes the saved workbook with calculated Base-case values.

Windows LibreOffice headless example:

```powershell
New-Item -ItemType Directory -Force .\out | Out-Null
& 'C:\Program Files\LibreOffice\program\soffice.com' --headless --convert-to xlsx --outdir .\out .\model\PE_Value_Creation_Model_synthetic.xlsx
```

Open and recalculate the workbook in Excel or LibreOffice after changing controls. The annual forecast uses one active scenario selector and one live operating model. The sensitivities are formula grids; the second sensitivity recalculates the annual cash and debt schedule for each AI realization point.

## Project files

- `build/build_model.py` creates the workbook.
- `check/verify_returns.py` performs the independent scenario and toggle checks.
- `docs/CASE_STUDY.md` is a two-page narrative with a worked assumption change.
- [Architecture](docs/ARCHITECTURE.md) shows the tab-level calculation flow.
- `docs/dashboard_preview.png` and `docs/vcp_tracker_preview.png` show workbook views.
- `tests/` contains workbook, control, and content checks.

No macros, network calls, external data or real-person names are used.
