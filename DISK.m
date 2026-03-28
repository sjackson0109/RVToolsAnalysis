// =============================================================================
// RVTools VM Right-Sizing — Disk / Free-Space Analysis
// =============================================================================
// Author:        Simon Jackson (@sjackson0109)
// Created:       2026-03-07
//
// Description:
//   Standalone Disk right-sizing query. Reads sheet data from the
//   "RVTools_Source" connector query and analyses vInfo, vDisk and vPartition.
//
//   For each VM it calculates:
//   - Total provisioned disk capacity (from vDisk).
//   - Minimum free-space percentage across all partitions (from vPartition).
//   - Whether any partition falls below DISK_MIN_FREE_PCT.
//   - How many additional MiB are needed to bring every partition up to the
//     threshold (Disk_Shortfall_MiB).
//
//   Outputs only VMs flagged "Expand disk" (i.e. at least one partition is
//   below the free-space threshold).
//
// Prerequisites:
//   Import RVTools-Source.m as a query named "RVTools-Source" (Connection Only).
//
// Configuration:
//   Edit the constants in the "CONFIGURATION" section below.
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

    // Disk thresholds
    DISK_MIN_FREE_PCT = 20,   // Flag VMs with any partition below this % free

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

    SafeGroup = (tbl as table, keyCol as text, newCol as text,
                 aggCol as text, aggFunc as function) as table =>
        if not List.Contains(Table.ColumnNames(tbl), keyCol)
           or not List.Contains(Table.ColumnNames(tbl), aggCol)
        then #table({keyCol, newCol}, {})
        else Table.Group(tbl, {keyCol}, {{newCol, each aggFunc(Table.Column(_, aggCol)), type number}}),

    // -------------------------------------------------------------------------
    // Load & type the sheets we need
    // -------------------------------------------------------------------------
    vInfo = SafeTransformColumnTypes(GetSheet("vInfo"), {
        {"VM", type text}, {"Powerstate", type text}, {"Template", type text},
        {"CPUs", Int64.Type}, {"Memory", Int64.Type}, {"Host", type text},
        {"Folder", type any}, {"Resource Group", type any}
    }),

    vDisk = SafeTransformColumnTypes(GetSheet("vDisk"), {
        {"VM", type text}, {"Capacity MiB", type number}
    }),

    vPartitionRaw = GetSheet("vPartition"),
    vPartitionCols = Table.ColumnNames(vPartitionRaw),
    
    // Find the free percentage column - try common variations
    FreePctCol = List.First(List.Select(vPartitionCols, each 
        Text.Contains(Text.Lower(_), "free") and 
        (Text.Contains(Text.Lower(_), "%") or Text.Contains(Text.Lower(_), "pct") or Text.Contains(Text.Lower(_), "percent"))
    ), ""),
    
    vPartition = SafeTransformColumnTypes(vPartitionRaw, {
        {"VM", type text}, {"Capacity MiB", type number}, {"Free Space MiB", type number}
    } & (if FreePctCol <> "" then {{FreePctCol, type number}} else {})),

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
    // Disk aggregation — total provisioned capacity per VM (vDisk)
    // -------------------------------------------------------------------------
    DiskTotals = SafeGroup(vDisk, "VM", "Disk_Total_Capacity_MiB", "Capacity MiB", List.Sum),

    // -------------------------------------------------------------------------
    // Partition aggregation — minimum free % per VM (vPartition)
    // -------------------------------------------------------------------------
    // Also compute the total shortfall MiB: how much additional space is needed
    // across all partitions below the threshold.
    // -------------------------------------------------------------------------
    // Calculate minimum free percentage per VM (handle missing column gracefully)
    PartMinFree = if FreePctCol = "" then 
        #table({"VM", "Partition_Min_Free_Pct"}, {})
    else
        SafeGroup(vPartition, "VM", "Partition_Min_Free_Pct", FreePctCol, List.Min),

    // Compute per-partition shortfall, then sum per VM
    PartShortfall =
        let
            hasNeeded = List.Contains(Table.ColumnNames(vPartition), "Capacity MiB")
                         and List.Contains(Table.ColumnNames(vPartition), "Free Space MiB"),
            annotated = if not hasNeeded then Table.AddColumn(vPartition, "Shortfall_MiB", each 0, type number)
                        else Table.AddColumn(vPartition, "Shortfall_MiB", each
                            let
                                cap   = _[Capacity MiB],
                                free  = _[Free Space MiB],
                                need  = if cap = null or cap = 0 then 0
                                        else (DISK_MIN_FREE_PCT / 100) * cap,
                                short = if free = null then 0
                                        else List.Max({0, need - free})
                            in short, type number),
            grouped = SafeGroup(annotated, "VM", "Disk_Shortfall_MiB", "Shortfall_MiB", List.Sum)
        in grouped,

    PartMinFreeMiB = SafeGroup(vPartition, "VM", "Partition_Min_Free_MiB", "Free Space MiB", List.Min),

    // -------------------------------------------------------------------------
    // Disk joins  (Filtered → DiskTotals → PartMinFree → PartShortfall → PartMinFreeMiB)
    // -------------------------------------------------------------------------
    D1 = SafeNestedJoin(Filtered, {"VM"}, DiskTotals, {"VM"}, "_DT", JoinKind.LeftOuter),
    D2 = SafeExpandTableColumn(D1, "_DT", {"Disk_Total_Capacity_MiB"}, {"Disk_Total_Capacity_MiB"}),

    D3 = SafeNestedJoin(D2, {"VM"}, PartMinFree, {"VM"}, "_PMF", JoinKind.LeftOuter),
    D4 = SafeExpandTableColumn(D3, "_PMF", {"Partition_Min_Free_Pct"}, {"Partition_Min_Free_Pct"}),

    D5 = SafeNestedJoin(D4, {"VM"}, PartShortfall, {"VM"}, "_PS", JoinKind.LeftOuter),
    D6 = SafeExpandTableColumn(D5, "_PS", {"Disk_Shortfall_MiB"}, {"Disk_Shortfall_MiB"}),

    D7 = SafeNestedJoin(D6, {"VM"}, PartMinFreeMiB, {"VM"}, "_PMFM", JoinKind.LeftOuter),
    D8 = SafeExpandTableColumn(D7, "_PMFM", {"Partition_Min_Free_MiB"}, {"Partition_Min_Free_MiB"}),

    // -------------------------------------------------------------------------
    // Disk free-space flag & action
    // -------------------------------------------------------------------------
    AddFlag = Table.AddColumn(D8, "Disk_FreeSpace_Flag",
        each let 
            pct = SafeGetField(_, "Partition_Min_Free_Pct"),
            freeMiB = SafeGetField(_, "Partition_Min_Free_MiB"),
            cap = SafeGetField(_, "Disk_Total_Capacity_MiB")
             in if pct = null and freeMiB <> null and cap <> null and cap > 0 then
                    // Calculate percentage if we have MiB values but no percentage
                    let calcPct = (freeMiB / cap) * 100
                    in if calcPct < DISK_MIN_FREE_PCT then "Low free space" else "OK"
                else if pct = null then "No data"
                else if pct < DISK_MIN_FREE_PCT then "Low free space"
                else "OK",
        type text),

    AddAction = Table.AddColumn(AddFlag, "Disk_Action",
        each if _[Disk_FreeSpace_Flag] = "Low free space" then "Expand disk"
             else if _[Disk_FreeSpace_Flag] = "No data"   then "No data"
             else "Keep disk",
        type text),

    // -------------------------------------------------------------------------
    // Filter to actionable recommendations only
    // -------------------------------------------------------------------------
    Actionable = Table.SelectRows(AddAction, each _[Disk_Action] = "Expand disk"),

    // -------------------------------------------------------------------------
    // Output
    // -------------------------------------------------------------------------
    Output = SafeSelectColumns(Actionable, {
        "VM", "Disk_Total_Capacity_MiB",
        "Partition_Min_Free_Pct", "Partition_Min_Free_MiB",
        "Disk_FreeSpace_Flag", "Disk_Shortfall_MiB", "Disk_Action"
    })
in
    Output
