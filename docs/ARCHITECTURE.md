# Workbook calculation flow

All names, data and outcomes in the diagrams and workbook are synthetic and illustrative.

```mermaid
flowchart LR
  Controls[Controls<br/>scenario, toggles, assumptions] --> Op[Operating Model<br/>revenue, margin, initiative EBITDA]
  History[Historical and baseline<br/>synthetic FY2023 to FY2025] --> Op
  Initiatives[Initiatives<br/>scenario ramps, costs, realization] --> Op
  Op --> Cash[Cash Flow and Debt<br/>tax, capex, working capital, amortization, sweep]
  Cash --> Returns[Returns<br/>entry, exit, MOIC, dated XIRR]
  Op --> Dashboard[Dashboard<br/>trends, EBITDA and EV bridges]
  Returns --> Dashboard
  Sensitivity[Sensitivity<br/>exit multiple, EBITDA and AI realization] --> Dashboard
  Tracker[VCP Tracker<br/>quarterly plan and synthetic actuals] --> Dashboard
  Checks[Checks and Definitions<br/>terminal diagnostics and KPI references]
  Controls -.-> Checks
  Op -.-> Checks
  Cash -.-> Checks
  Returns -.-> Checks
```

**Legend:** solid arrows show formula dependencies. Dashed arrows show read-only checks. Checks are terminal diagnostics and do not feed assumptions, operating calculations or outputs.

The workbook uses one active scenario selector. The operating model and cash flow schedule recalculate when the selector changes. The return sensitivities use visible formula grids. The AI realization grid has a separate visible calculation block that recalculates annual taxes, debt paydown and retained cash for each realization point.
