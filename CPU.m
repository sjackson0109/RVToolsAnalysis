// =============================================================================
// RVTools VM Right-Sizing — CPU Analysis
// =============================================================================
// Author:        Simon Jackson (@sjackson0109)
// Created:       2026-03-07
//
// Description:
//   Standalone CPU right-sizing query. Reads sheet data from the
//   "RVTools_Source" connector query and analyses vInfo, vCPU and vHost.
//
//   Outputs only VMs whose CPU allocation should change (Decrease / Increase).
//
// Prerequisites:
//   Import RVTools-Source.m as a query named "RVTools-Source" (Connection Only).
//
// Configuration:
//   Edit the constants in the "CONFIGURATION" section below.
//
// CPU Sizing Notes:
//   RVTools vCPU "Max" = CPU entitlement (CPUs × Host Speed), NOT peak demand.
//   "Overall" = actual current CPU demand in MHz.
//   A 50 %-per-step downsize cap compensates for point-in-time snapshot data.
// =============================================================================
let
    // =========================================================================
    // CONFIGURATION — edit these values to suit your environment
    // =========================================================================

    // Scope filters
    INCLUDE_POWERED_OFF             = false,   // Include powered-off VMs?
    INCLUDE_TEMPLATES               = false,   // Include VM templates?
    EXCLUDE_VM_NAME_CONTAINS        = "",      // Comma-separated substrings (e.g. "backup,test")
    EXCLUDE_FOLDER_CONTAINS         = "",      // Comma-separated folder substrings
    EXCLUDE_RESOURCE_GROUP_CONTAINS = "",      // Comma-separated resource-group substrings

    // CPU sizing thresholds
    CPU_TARGET_UTIL_PCT       = 70,   // Target CPU utilisation %
    CPU_HEADROOM_PCT          = 20,   // Additional headroom % on top of demand
    CPU_MIN_VCPU              = 2,    // Minimum vCPU count
    CPU_MAX_DOWNSIZE_UTIL_PCT = 50,   // Only decrease if util < this %
    CPU_MIN_UPSIZE_UTIL_PCT   = 85,   // Only increase if util >= this %
    CPU_MAX_INCREASE_PCT      = 25,   // Max % increase per sizing pass

    // =========================================================================
    // END OF CONFIGURATION
    // =========================================================================

    // -------------------------------------------------------------------------
    // RVTools data source  (references the "RVTools-Source" connector query)
    // -------------------------------------------------------------------------
    GetSheet = (sheetName as text) as table =>
        let
            Row  = try #"RVTools-Source"{[Item=sheetName, Kind="Sheet"]} otherwise null,
            Data = if Row = null then #table({}, {}) else Row[Data],
            First = if Table.RowCount(Data) = 0 then ""
                    else List.First(Table.ColumnNames(Data), "")
        in
            if Table.RowCount(Data) = 0 then Data
            else if Text.StartsWith(First, "Column") or First = "" then
                Table.PromoteHeaders(Data, [PromoteAllScalars=true])
            else Data,

    // -------------------------------------------------------------------------
    // Helpers
    // -------------------------------------------------------------------------
    SafeTransformColumnTypes = (tbl as table, transforms as list) as table =>
        let
            cols  = Table.ColumnNames(tbl),
            valid = List.Select(transforms, each List.Contains(cols, _{0}))
        in if List.Count(valid) > 0
           then try Table.TransformColumnTypes(tbl, valid) otherwise tbl
           else tbl,

    SafeGetField = (rec as record, fld as text) as any =>
        try Record.Field(rec, fld) otherwise null,

    SafeSelectColumns = (tbl as table, names as list) as table =>
        let
            cols  = Table.ColumnNames(tbl),
            valid = List.Select(names, each List.Contains(cols, _))
        in if List.Count(valid) > 0 then Table.SelectColumns(tbl, valid) else tbl,

    SafeFilterRows = (tbl as table, col as text, cond as function) as table =>
        if List.Contains(Table.ColumnNames(tbl), col)
        then Table.SelectRows(tbl, cond)
        else tbl,

    SafeNestedJoin = (lt as table, lk as list, rt as table, rk as list,
                      nc as text, jk as number) as table =>
        let
            lok = List.AllTrue(List.Transform(lk, each List.Contains(Table.ColumnNames(lt), _))),
            rok = List.AllTrue(List.Transform(rk, each List.Contains(Table.ColumnNames(rt), _)))
        in if lok and rok then Table.NestedJoin(lt, lk, rt, rk, nc, jk) else lt,

    SafeExpandTableColumn = (tbl as table, col as text, src as list, dst as list) as table =>
        try Table.ExpandTableColumn(tbl, col, src, dst) otherwise tbl,

    // -------------------------------------------------------------------------
    // Load & type the sheets we need
    // -------------------------------------------------------------------------
    vInfo = SafeTransformColumnTypes(GetSheet("vInfo"), {
        {"VM", type text}, {"Powerstate", type text}, {"Template", type text},
        {"CPUs", Int64.Type}, {"Memory", Int64.Type}, {"Host", type text},
        {"Folder", type any}, {"Resource Group", type any}
    }),

    vCPU = SafeTransformColumnTypes(GetSheet("vCPU"), {
        {"VM", type text}, {"CPUs", Int64.Type},
        {"Overall", type number}, {"Host", type text}
    }),

    vHost = SafeTransformColumnTypes(GetSheet("vHost"), {
        {"Host", type text}, {"Speed", type number}
    }),

    // -------------------------------------------------------------------------
    // VM base list + filtering
    // -------------------------------------------------------------------------
    Base = SafeSelectColumns(vInfo,
        {"VM", "Powerstate", "Template", "CPUs", "Memory", "Host", "Folder", "Resource Group"}),

    ExcludeVMList     = List.Select(List.Transform(Text.Split(Text.From(EXCLUDE_VM_NAME_CONTAINS ?? ""), ","), Text.Trim), each _ <> ""),
    ExcludeFolderList = List.Select(List.Transform(Text.Split(Text.From(EXCLUDE_FOLDER_CONTAINS ?? ""), ","), Text.Trim), each _ <> ""),
    ExcludeRGList     = List.Select(List.Transform(Text.Split(Text.From(EXCLUDE_RESOURCE_GROUP_CONTAINS ?? ""), ","), Text.Trim), each _ <> ""),

    F1 = if INCLUDE_TEMPLATES = true then Base
         else SafeFilterRows(Base, "Template", each try ([Template] <> true) otherwise true),

    F2 = if INCLUDE_POWERED_OFF = true then F1
         else SafeFilterRows(F1, "Powerstate", each try Text.Lower([Powerstate]) = "poweredon" otherwise false),

    F3 = if List.Count(ExcludeVMList) = 0 then F2
         else SafeFilterRows(F2, "VM", each
             let n = try Text.Lower(Text.From([VM])) otherwise ""
             in not List.AnyTrue(List.Transform(ExcludeVMList, each Text.Contains(n, Text.Lower(_))))),

    F4 = if List.Count(ExcludeFolderList) = 0 then F3
         else SafeFilterRows(F3, "Folder", each
             let f = try Text.Lower(Text.From([Folder])) otherwise ""
             in not List.AnyTrue(List.Transform(ExcludeFolderList, each Text.Contains(f, Text.Lower(_))))),

    Filtered = if List.Count(ExcludeRGList) = 0 then F4
               else SafeFilterRows(F4, "Resource Group", each
                   let rg = try Text.Lower(Text.From([Resource Group])) otherwise ""
                   in not List.AnyTrue(List.Transform(ExcludeRGList, each Text.Contains(rg, Text.Lower(_))))),

    // -------------------------------------------------------------------------
    // CPU joins  (vInfo → vCPU → vHost)
    // -------------------------------------------------------------------------
    J1 = SafeNestedJoin(Filtered, {"VM"}, vCPU, {"VM"}, "CPU", JoinKind.LeftOuter),
    J2 = SafeExpandTableColumn(J1, "CPU", {"Overall"}, {"CPU_Overall_MHz"}),
    J3 = SafeNestedJoin(J2, {"Host"}, vHost, {"Host"}, "HOST", JoinKind.LeftOuter),
    J4 = SafeExpandTableColumn(J3, "HOST", {"Speed"}, {"Host_Speed_MHz"}),

    // -------------------------------------------------------------------------
    // CPU demand and utilisation
    // -------------------------------------------------------------------------
    // "Overall" = actual demand in MHz.  Demand_Cores = Overall / Host_Speed.
    // Util_Pct = Demand_Cores / CPUs × 100.
    // -------------------------------------------------------------------------
    AddDemandMHz = Table.AddColumn(J4, "CPU_Demand_MHz",
        each SafeGetField(_, "CPU_Overall_MHz"), type number),

    AddDemandCores = Table.AddColumn(AddDemandMHz, "CPU_Demand_Cores",
        each let spd = SafeGetField(_, "Host_Speed_MHz"), mhz = _[CPU_Demand_MHz]
             in if spd = null or spd = 0 or mhz = null then null else mhz / spd,
        type number),

    AddUtilPct = Table.AddColumn(AddDemandCores, "CPU_Util_Pct",
        each let c = _[CPUs], cores = _[CPU_Demand_Cores]
             in if c = null or c = 0 or cores = null then null else (cores / c) * 100,
        type number),

    // -------------------------------------------------------------------------
    // CPU recommendation
    // -------------------------------------------------------------------------
    // ideal  = ceil( demandCores × headroom / target )
    // Downsize: clamped to max 50 % reduction per step.
    // Upsize:   clamped to CPU_MAX_INCREASE_PCT per step.
    // -------------------------------------------------------------------------
    AddRec = Table.AddColumn(AddUtilPct, "CPU_Recommended_vCPU_Raw",
        each
            let
                current     = _[CPUs],
                demandCores = _[CPU_Demand_Cores],
                target      = CPU_TARGET_UTIL_PCT / 100,
                headroom    = 1 + (CPU_HEADROOM_PCT / 100),
                ideal       = if demandCores = null or target = 0 then null
                              else Number.RoundUp((demandCores * headroom) / target),
                idealClamped = if ideal = null then null
                               else List.Max({CPU_MIN_VCPU, ideal}),
                minDown     = if current = null then null
                              else List.Max({CPU_MIN_VCPU, Number.RoundUp(current / 2)}),
                maxUp       = if current = null then null
                              else Number.RoundUp(current * (1 + CPU_MAX_INCREASE_PCT / 100)),
                raw = if current = null then CPU_MIN_VCPU
                      else if idealClamped = null then current
                      else if idealClamped < current then List.Max({idealClamped, minDown})
                      else if idealClamped > current then List.Min({idealClamped, maxUp})
                      else current
            in raw,
        Int64.Type),

    // -------------------------------------------------------------------------
    // CPU action (threshold-gated)
    // -------------------------------------------------------------------------
    AddAction = Table.AddColumn(AddRec, "CPU_Action",
        each
            let cur = _[CPUs], rec = _[CPU_Recommended_vCPU_Raw], util = _[CPU_Util_Pct]
            in  if cur = null or rec = null then "No data"
                else if rec < cur and util <> null and util < CPU_MAX_DOWNSIZE_UTIL_PCT then "Decrease vCPU"
                else if rec > cur and util <> null and util >= CPU_MIN_UPSIZE_UTIL_PCT   then "Increase vCPU"
                else "Keep vCPU",
        type text),

    AddFinal = Table.AddColumn(AddAction, "CPU_Recommended_vCPU",
        each let cur = _[CPUs], raw = _[CPU_Recommended_vCPU_Raw], act = _[CPU_Action]
             in if act = "Keep vCPU" then cur else if act = "No data" then null else raw,
        Int64.Type),

    // -------------------------------------------------------------------------
    // Filter to actionable recommendations only
    // -------------------------------------------------------------------------
    Actionable = Table.SelectRows(AddFinal, each
        _[CPU_Action] = "Decrease vCPU" or _[CPU_Action] = "Increase vCPU"),

    // -------------------------------------------------------------------------
    // Output
    // -------------------------------------------------------------------------
    Output = SafeSelectColumns(Actionable, {
        "VM", "CPUs",
        "CPU_Demand_MHz", "Host_Speed_MHz", "CPU_Demand_Cores", "CPU_Util_Pct",
        "CPU_Recommended_vCPU", "CPU_Action"
    })
in
    Output
