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

## Three views — click a heading to switch

Every page has a **Page Navigator** strip at the top (Timeline · Dependencies · Authority).
Click a heading and Power BI jumps to that view. Within a view, click any bar, row, or
slicer and the other visuals cross-filter.

- **Timeline** — Gantt schedule (each permit's start month → finish, sorted by start) + a
  detail table (permit, phase, lead days, start month, duration).
- **Dependencies** — Blocker → Blocked table, a "most-blocking permits" bar, and a Blocker
  slicer. Shows which permits gate others (EIA and the Building Permit block the most).
- **Authority** — permits grouped by authority (matrix), a "permits per authority" bar, and
  a contacts table (authority, phone, website).

## Model (DAX measures + tables)

Tables: **Permits** (23 rows), **Authorities** (contacts, related to Permits), **Dependencies**
(Blocker→Blocked edges). Measures: `Permit Count`, `Total Lead Days`, `Total Budget`,
`Start (month)`, `Duration (months)`, `Dependency Count` — in
`PermitAnalytics.SemanticModel/definition/tables/`.

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
