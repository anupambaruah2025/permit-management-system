"""Generate the PermitAnalytics semantic model (TMDL) for the thick PBIP."""
import os, json, uuid

ROOT = os.path.dirname(os.path.abspath(__file__))
SM = os.path.join(ROOT, "PermitAnalytics.SemanticModel")
DEF = os.path.join(SM, "definition")
TBL = os.path.join(DEF, "tables")
os.makedirs(TBL, exist_ok=True)

def w(path, text):
    # UTF-8 without BOM, CRLF to match Power BI Desktop
    with open(path, "w", encoding="utf-8", newline="\r\n") as f:
        f.write(text)

# id, name, authority, phase label, phaseNo, leadDays, cost, risk, status, startMonth, durMonths
P = [
 (1,"Company Registration","DBD","1. Development",1,30,5000,"Low","Completed",0,1),
 (2,"BOI Promotion Certificate","BOI","1. Development",1,120,0,"Critical","Completed",0.5,4),
 (3,"Foreign Business License","DBD","1. Development",1,60,10000,"Low","Completed",0.5,2),
 (4,"Land Use & Zoning","DPW","1. Development",1,30,2000,"Low","Completed",0,1),
 (5,"Environmental Impact Assessment","ONEP","1. Development",1,180,350000,"Critical","Completed",0.5,6),
 (6,"Building Construction Permit","DPW","2. Design & Planning",2,60,15000,"Critical","Completed",6,2),
 (7,"Road Access / Highway","DPW","2. Design & Planning",2,45,8000,"High","In Progress",6,1.5),
 (8,"Grid Connection Approval","PEA/MEA","2. Design & Planning",2,60,75000,"High","Completed",2,2),
 (9,"Power Availability (ERC)","ERC","2. Design & Planning",2,30,0,"Critical","Completed",0.5,1),
 (10,"HV Substation Approval","PEA/MEA","2. Design & Planning",2,90,150000,"High","Completed",6,3),
 (11,"Water Supply Connection","MWA/PWA","3. Construction",3,45,30000,"Low","Completed",8,1.5),
 (12,"Wastewater / Drainage","MWA","3. Construction",3,45,37500,"Low","Completed",8,1.5),
 (13,"Fire Protection System","DPW","3. Construction",3,30,115000,"High","Completed",8,1),
 (14,"Generator & Fuel Storage","DOEB","3. Construction",3,60,15000,"High","In Progress",8,2),
 (15,"Telecom / Fiber","NBTC","3. Construction",3,45,10000,"Low","In Progress",8,1.5),
 (16,"Electrical Inspection","PEA/MEA","3. Construction",3,14,5000,"Low","In Progress",10,0.5),
 (17,"Fire Safety Inspection","Fire Authority","4. Commissioning",4,7,0,"Low","In Progress",11,0.4),
 (18,"Building Completion","DPW","4. Commissioning",4,7,3000,"Low","In Progress",11,0.4),
 (19,"Occupancy Certificate","DPW","4. Commissioning",4,14,2000,"High","In Progress",11.3,0.6),
 (20,"Commissioning & Utility Acceptance","Utilities","4. Commissioning",4,21,22500,"Low","Pending",11.5,0.8),
 (21,"Factory License","DIW","5. Operations",5,30,5000,"Low","Pending",12,1),
 (22,"PDPA Compliance","PDPC","5. Operations",5,30,0,"Low","Pending",12,1),
 (23,"Cybersecurity Compliance","NCSA","5. Operations",5,30,0,"Low","Pending",12,1),
]

def esc(s): return s.replace('"','""')
rows = []
for r in P:
    i,name,auth,ph,phn,days,cost,risk,status,gs,gd = r
    rows.append('\t\t\t\t\t{ %d, "%s", "%s", "%s", %d, %d, %d, "%s", "%s", %s, %s }'
                % (i, esc(name), esc(auth), esc(ph), phn, days, cost, risk, status, float(gs), float(gd)))
rows_dax = ",\n".join(rows)

# ---- Authorities (one row per distinct authority, with contacts) ----
CONTACTS = {
 "DBD":("0-2141-4271-9","www.dbd.go.th"), "BOI":("+66-2-553-8214","www.boi.go.th"),
 "DPW":("0-2243-9200","www.dlt.go.th"), "ONEP":("0-2298-3600","www.onep.go.th"),
 "PEA/MEA":("0-2399-9999","www.pea.co.th"), "ERC":("via PEA/MEA","www.erc.or.th"),
 "MWA/PWA":("0-2299-3000","www.mwa.co.th"), "MWA":("0-2299-3000","www.mwa.co.th"),
 "DOEB":("0-2202-8300-4","www.doeb.go.th"), "NBTC":("0-2279-7800","www.nbtc.go.th"),
 "Fire Authority":("199 / local","www.dlt.go.th"), "Utilities":("PEA/MEA/MWA","-"),
 "DIW":("0-2201-3000","www.diw.go.th"), "PDPC":("0-2141-9000","www.pdpa.go.th"),
 "NCSA":("0-2578-2000","www.ncsa.go.th"),
}
auth_order = []
for r in P:
    if r[2] not in auth_order: auth_order.append(r[2])
auth_rows = []
for a in auth_order:
    ph_, web = CONTACTS.get(a, ("-","-"))
    auth_rows.append('\t\t\t\t\t{ "%s", "%s", "%s" }' % (esc(a), esc(ph_), esc(web)))
auth_dax = ",\n".join(auth_rows)

# ---- Dependencies (Blocker -> Blocked edges) ----
DEP = [
 ("Company Registration","BOI Promotion Certificate"),
 ("Company Registration","Foreign Business License"),
 ("Land Use & Zoning","Environmental Impact Assessment"),
 ("Power Availability (ERC)","BOI Promotion Certificate"),
 ("Power Availability (ERC)","Grid Connection Approval"),
 ("Environmental Impact Assessment","Building Construction Permit"),
 ("Environmental Impact Assessment","Water Supply Connection"),
 ("Environmental Impact Assessment","Wastewater / Drainage"),
 ("Environmental Impact Assessment","Generator & Fuel Storage"),
 ("Environmental Impact Assessment","Factory License"),
 ("Building Construction Permit","Road Access / Highway"),
 ("Building Construction Permit","Water Supply Connection"),
 ("Building Construction Permit","Wastewater / Drainage"),
 ("Building Construction Permit","Fire Protection System"),
 ("Building Construction Permit","Generator & Fuel Storage"),
 ("Building Construction Permit","Telecom / Fiber"),
 ("Building Construction Permit","Building Completion"),
 ("Grid Connection Approval","HV Substation Approval"),
 ("Grid Connection Approval","Electrical Inspection"),
 ("Fire Protection System","Fire Safety Inspection"),
 ("Building Completion","Occupancy Certificate"),
 ("Fire Safety Inspection","Occupancy Certificate"),
 ("Electrical Inspection","Occupancy Certificate"),
 ("Occupancy Certificate","Commissioning & Utility Acceptance"),
 ("Occupancy Certificate","Factory License"),
]
dep_rows = ['\t\t\t\t\t{ "%s", "%s" }' % (esc(b), esc(t)) for b,t in DEP]
dep_dax = ",\n".join(dep_rows)

g = lambda: str(uuid.uuid4())

permits_tmdl = f"""/// All 23 permits for the CHN1A data-center project, with phase, lead time, cost, risk and status.
table Permits
\tlineageTag: {g()}

\t/// Total number of permits in the current filter context.
\tmeasure 'Permit Count' = COUNTROWS ( Permits )
\t\tformatString: #,##0
\t\tlineageTag: {g()}

\t/// Permits with status Completed.
\tmeasure Completed = CALCULATE ( [Permit Count], Permits[Status] = "Completed" )
\t\tformatString: #,##0
\t\tlineageTag: {g()}

\t/// Permits currently In Progress.
\tmeasure 'In Progress' = CALCULATE ( [Permit Count], Permits[Status] = "In Progress" )
\t\tformatString: #,##0
\t\tlineageTag: {g()}

\t/// Permits not yet started.
\tmeasure Pending = CALCULATE ( [Permit Count], Permits[Status] = "Pending" )
\t\tformatString: #,##0
\t\tlineageTag: {g()}

\t/// Sum of lead-time days across permits.
\tmeasure 'Total Lead Days' = SUM ( Permits[LeadDays] )
\t\tformatString: #,##0
\t\tlineageTag: {g()}

\t/// Sum of estimated permit cost (THB).
\tmeasure 'Total Budget' = SUM ( Permits[Cost] )
\t\tformatString: #,##0
\t\tlineageTag: {g()}

\t/// Share of permits completed.
\tmeasure '% Complete' = DIVIDE ( [Completed], [Permit Count] )
\t\tformatString: 0%
\t\tlineageTag: {g()}

\t/// Gantt offset: project month each permit starts (transparent segment).
\tmeasure 'Start (month)' = SUM ( Permits[StartMonth] )
\t\tformatString: 0.0
\t\tlineageTag: {g()}

\t/// Gantt bar length in months.
\tmeasure 'Duration (months)' = SUM ( Permits[DurationMonths] )
\t\tformatString: 0.0
\t\tlineageTag: {g()}

\tcolumn ID
\t\tdataType: int64
\t\tisKey
\t\tlineageTag: {g()}
\t\tsummarizeBy: none
\t\tsourceColumn: ID

\tcolumn Permit
\t\tdataType: string
\t\tlineageTag: {g()}
\t\tsummarizeBy: none
\t\tsourceColumn: Permit

\tcolumn Authority
\t\tdataType: string
\t\tlineageTag: {g()}
\t\tsummarizeBy: none
\t\tsourceColumn: Authority

\tcolumn Phase
\t\tdataType: string
\t\tlineageTag: {g()}
\t\tsummarizeBy: none
\t\tsourceColumn: Phase
\t\tsortByColumn: PhaseNo

\tcolumn PhaseNo
\t\tdataType: int64
\t\tisHidden
\t\tlineageTag: {g()}
\t\tsummarizeBy: none
\t\tsourceColumn: PhaseNo

\tcolumn LeadDays
\t\tdataType: int64
\t\tlineageTag: {g()}
\t\tsummarizeBy: sum
\t\tsourceColumn: LeadDays

\tcolumn Cost
\t\tdataType: int64
\t\tformatString: #,##0
\t\tlineageTag: {g()}
\t\tsummarizeBy: sum
\t\tsourceColumn: Cost

\tcolumn Risk
\t\tdataType: string
\t\tlineageTag: {g()}
\t\tsummarizeBy: none
\t\tsourceColumn: Risk

\tcolumn Status
\t\tdataType: string
\t\tlineageTag: {g()}
\t\tsummarizeBy: none
\t\tsourceColumn: Status

\tcolumn StartMonth
\t\tdataType: double
\t\tisHidden
\t\tlineageTag: {g()}
\t\tsummarizeBy: none
\t\tsourceColumn: StartMonth

\tcolumn DurationMonths
\t\tdataType: double
\t\tlineageTag: {g()}
\t\tsummarizeBy: sum
\t\tsourceColumn: DurationMonths

\tpartition Permits = calculated
\t\tmode: import
\t\tsource =
\t\t\t\tDATATABLE (
\t\t\t\t\t"ID", INTEGER,
\t\t\t\t\t"Permit", STRING,
\t\t\t\t\t"Authority", STRING,
\t\t\t\t\t"Phase", STRING,
\t\t\t\t\t"PhaseNo", INTEGER,
\t\t\t\t\t"LeadDays", INTEGER,
\t\t\t\t\t"Cost", INTEGER,
\t\t\t\t\t"Risk", STRING,
\t\t\t\t\t"Status", STRING,
\t\t\t\t\t"StartMonth", DOUBLE,
\t\t\t\t\t"DurationMonths", DOUBLE,
\t\t\t\t\t{{
{rows_dax}
\t\t\t\t\t}}
\t\t\t\t)

\tannotation PBI_ResultType = Table
"""

authorities_tmdl = f"""/// One row per government authority, with contact details.
table Authorities
\tlineageTag: {g()}

\tcolumn Authority
\t\tdataType: string
\t\tisKey
\t\tlineageTag: {g()}
\t\tsummarizeBy: none
\t\tsourceColumn: Authority

\tcolumn Phone
\t\tdataType: string
\t\tlineageTag: {g()}
\t\tsummarizeBy: none
\t\tsourceColumn: Phone

\tcolumn Website
\t\tdataType: string
\t\tdataCategory: WebUrl
\t\tlineageTag: {g()}
\t\tsummarizeBy: none
\t\tsourceColumn: Website

\tpartition Authorities = calculated
\t\tmode: import
\t\tsource =
\t\t\t\tDATATABLE (
\t\t\t\t\t"Authority", STRING,
\t\t\t\t\t"Phone", STRING,
\t\t\t\t\t"Website", STRING,
\t\t\t\t\t{{
{auth_dax}
\t\t\t\t\t}}
\t\t\t\t)
"""

dependencies_tmdl = f"""/// Blocker -> Blocked permit dependency edges (which permit must finish before another can start).
table Dependencies
\tlineageTag: {g()}

\t/// Number of dependency links in context.
\tmeasure 'Dependency Count' = COUNTROWS ( Dependencies )
\t\tformatString: #,##0
\t\tlineageTag: {g()}

\tcolumn Blocker
\t\tdataType: string
\t\tlineageTag: {g()}
\t\tsummarizeBy: none
\t\tsourceColumn: Blocker

\tcolumn Blocked
\t\tdataType: string
\t\tlineageTag: {g()}
\t\tsummarizeBy: none
\t\tsourceColumn: Blocked

\tpartition Dependencies = calculated
\t\tmode: import
\t\tsource =
\t\t\t\tDATATABLE (
\t\t\t\t\t"Blocker", STRING,
\t\t\t\t\t"Blocked", STRING,
\t\t\t\t\t{{
{dep_dax}
\t\t\t\t\t}}
\t\t\t\t)
"""

relationships_tmdl = f"""relationship {g()}
\tfromColumn: Permits.Authority
\ttoColumn: Authorities.Authority
"""

model_tmdl = """model PermitAnalytics
\tculture: en-US
\tdefaultPowerBIDataSourceVersion: powerBI_V3
\tdiscourageImplicitMeasures
\tsourceQueryCulture: en-US
\tdataAccessOptions
\t\tlegacyRedirects
\t\treturnErrorValuesAsNull

\tannotation PBI_QueryOrder = ["Permits","Authorities","Dependencies"]

\tannotation __PBI_TimeIntelligenceEnabled = 0
"""

database_tmdl = f"""database PermitAnalytics
\tid: {g()}
\tcompatibilityLevel: 1567
"""

pbism = json.dumps({
  "$schema":"https://developer.microsoft.com/json-schemas/fabric/item/semanticModel/definitionProperties/1.0.0/schema.json",
  "version":"4.2","settings":{}
}, indent=2)

platform = json.dumps({
  "$schema":"https://developer.microsoft.com/json-schemas/fabric/gitIntegration/platformProperties/2.0.0/schema.json",
  "metadata":{"type":"SemanticModel","displayName":"PermitAnalytics"},
  "config":{"version":"2.0","logicalId":str(uuid.uuid4())}
}, indent=2)

w(os.path.join(TBL,"Permits.tmdl"), permits_tmdl)
w(os.path.join(TBL,"Authorities.tmdl"), authorities_tmdl)
w(os.path.join(TBL,"Dependencies.tmdl"), dependencies_tmdl)
w(os.path.join(DEF,"relationships.tmdl"), relationships_tmdl)
w(os.path.join(DEF,"model.tmdl"), model_tmdl)
w(os.path.join(DEF,"database.tmdl"), database_tmdl)
w(os.path.join(SM,"definition.pbism"), pbism)
w(os.path.join(SM,".platform"), platform)

# Wire the report to this model (thick, byPath)
pbir_path = os.path.join(ROOT,"PermitAnalytics.Report","definition.pbir")
pbir = {
  "$schema":"https://developer.microsoft.com/json-schemas/fabric/item/report/definitionProperties/2.0.0/schema.json",
  "version":"4.0",
  "datasetReference":{"byPath":{"path":"../PermitAnalytics.SemanticModel"}}
}
w(pbir_path, json.dumps(pbir, indent=2))

print("Semantic model written to", SM)
print("Report wired byPath ->", pbir_path)
