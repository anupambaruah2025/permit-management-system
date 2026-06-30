# PermitAnalytics — Power BI report (PBIP)

A real **Power BI Project** (PBIP) for the CHN1A permit data: a TMDL semantic model + a PBIR report with native slicers and cross-filtering. Open it in **Power BI Desktop**.

## What's here

```
powerbi/
├─ PermitAnalytics.pbip                 ← open THIS in Power BI Desktop
├─ PermitAnalytics.SemanticModel/       ← the data model (TMDL)
│   └─ definition/tables/Permits.tmdl   ← 23 permits + DAX measures
└─ PermitAnalytics.Report/              ← the report (PBIR)
    └─ definition/pages/…/visuals/      ← KPI cards, donut, bar, table, slicers
```

The model is **self-contained** — the 23 permits are embedded as a DAX calculated table (`DATATABLE`), so there is no external data source to refresh and nothing to connect.

## How to open

1. **Install Power BI Desktop (free):**
   - Microsoft Store → search "Power BI Desktop" → Install, **or**
   - https://www.microsoft.com/download/details.aspx?id=58494
2. Double-click **`PermitAnalytics.pbip`** (or File → Open in Power BI Desktop).
3. The model loads with all data; the report opens on the **Overview** page.

## What you get (native Power BI interactivity)

- **KPI cards:** Permits, Completed, In Progress, Budget (THB)
- **Donut:** status mix · **Column:** permits by phase · **Bar:** budget by phase
- **Slicers:** Phase, Status, Risk
- **Table:** full permit register (ID, Permit, Authority, Phase, Lead Days, Cost, Risk, Status)

Click any slicer, donut slice, or bar and **every visual cross-filters** — the real Power BI behaviour. Edit visuals, add pages, or publish to the Power BI service / Fabric from Desktop.

## Measures (DAX)

`Permit Count`, `Completed`, `In Progress`, `Pending`, `Total Lead Days`, `Total Budget`, `% Complete` — defined in `PermitAnalytics.SemanticModel/definition/tables/Permits.tmdl`.

## Regenerate the model

If you change the permit data, edit `generate_model.py` and run:

```bash
python generate_model.py
```

Then validate the report with [pbir-cli](https://pypi.org/project/pbir-cli/):

```bash
pip install pbir-cli
pbir validate "PermitAnalytics.Report" --fields
```

Built with the [power-bi-agentic-development](https://github.com/data-goblin/power-bi-agentic-development) skills (PBIP/TMDL/PBIR).
