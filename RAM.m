// =============================================================================
// RVTools VM Right-Sizing — Memory Analysis
// =============================================================================
// Author:        Simon Jackson (@sjackson0109)
// Created:       2026-03-07
//
// Description:
//   Standalone Memory right-sizing query. Reads sheet data from the
//   "RVTools_Source" connector query and analyses vInfo and vMemory.
//
//   Outputs only VMs whose Memory allocation should change (Decrease / Increase).
//
//   Includes a 50 % per-step downsize cap — a VM's recommended allocation will
//   never be less than half of its current allocation in a single pass, even if
//   utilisation-based math suggests a larger cut.  This compensates for
//   point-in-time snapshot data.
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

    // Memory sizing thresholds
    MEM_BASIS                 = "ACTIVE",  // "ACTIVE" or "CONSUMED"
    MEM_TARGET_UTIL_PCT       = 60,   // Target memory utilisation %
    MEM_HEADROOM_PCT          = 20,   // Additional headroom % on top of usage
    MEM_MIN_MIB               = 2048, // Minimum RAM in MiB
    MEM_STEP_MIB              = 512,  // Round up to nearest multiple of this
    MEM_MAX_DOWNSIZE_UTIL_PCT = 35,   // Only decrease if util < this %
    MEM_MIN_UPSIZE_UTIL_PCT   = 75,   // Only increase if util >= this %

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

    RoundUpToStep = (value as number, step as number) as number =>
        if step = 0 then value else Number.RoundUp(value / step) * step,

    // -------------------------------------------------------------------------
    // Load & type the sheets we need
    // -------------------------------------------------------------------------
    vInfo = SafeTransformColumnTypes(GetSheet("vInfo"), {
        {"VM", type text}, {"Powerstate", type text}, {"Template", type text},
        {"CPUs", Int64.Type}, {"Memory", Int64.Type}, {"Host", type text},
        {"Folder", type any}, {"Resource Group", type any}
    }),

    vMemory = SafeTransformColumnTypes(GetSheet("vMemory"), {
        {"VM", type text}, {"Size MiB", type number},
        {"Active", type number}, {"Consumed", type number}
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
    // Memory join  (vInfo → vMemory)
    // -------------------------------------------------------------------------
    J1 = SafeNestedJoin(Filtered, {"VM"}, vMemory, {"VM"}, "MEM", JoinKind.LeftOuter),
    J2 = SafeExpandTableColumn(J1, "MEM",
        {"Size MiB", "Active", "Consumed"},
        {"Mem_Size_MiB", "Mem_Active_MiB", "Mem_Consumed_MiB"}),

    // -------------------------------------------------------------------------
    // Memory basis and utilisation
    // -------------------------------------------------------------------------
    AddBasis = Table.AddColumn(J2, "Mem_Basis_MiB",
        each if MEM_BASIS = "CONSUMED" then SafeGetField(_, "Mem_Consumed_MiB")
             else SafeGetField(_, "Mem_Active_MiB"),
        type number),

    AddUtilPct = Table.AddColumn(AddBasis, "Mem_Util_Pct",
        each let s = SafeGetField(_, "Mem_Size_MiB"), b = _[Mem_Basis_MiB]
             in if s = null or s = 0 or b = null then null else (b / s) * 100,
        type number),

    // -------------------------------------------------------------------------
    // Memory recommendation
    // -------------------------------------------------------------------------
    // idealMiB  = RoundUpToStep( basis × headroom / target , MEM_STEP_MIB )
    // Downsize: clamped to max 50 % reduction per step (never cut more than half
    //           the current allocation in one pass).
    // Final recommendation is at least MEM_MIN_MIB.
    // -------------------------------------------------------------------------
    AddRec = Table.AddColumn(AddUtilPct, "Mem_Recommended_MiB_Raw",
        each
            let
                current     = SafeGetField(_, "Mem_Size_MiB"),
                basis       = _[Mem_Basis_MiB],
                target      = MEM_TARGET_UTIL_PCT / 100,
                headroom    = 1 + (MEM_HEADROOM_PCT / 100),
                rawMiB      = if basis = null or target = 0 then null
                              else (basis * headroom) / target,
                idealMiB    = if rawMiB = null then null
                              else List.Max({MEM_MIN_MIB, RoundUpToStep(rawMiB, MEM_STEP_MIB)}),
                // 50 % per-step downsize cap
                minDown     = if current = null then null
                              else List.Max({MEM_MIN_MIB,
                                   RoundUpToStep(current / 2, MEM_STEP_MIB)}),
                raw = if current = null then MEM_MIN_MIB
                      else if idealMiB = null then current
                      else if idealMiB < current then List.Max({idealMiB, minDown})
                      else idealMiB
            in raw,
        type number),

    // -------------------------------------------------------------------------
    // Memory action (threshold-gated)
    // -------------------------------------------------------------------------
    AddAction = Table.AddColumn(AddRec, "Mem_Action",
        each
            let cur  = SafeGetField(_, "Mem_Size_MiB"),
                rec  = _[Mem_Recommended_MiB_Raw],
                util = _[Mem_Util_Pct]
            in  if cur = null or rec = null then "No data"
                else if rec < cur and util <> null and util < MEM_MAX_DOWNSIZE_UTIL_PCT then "Decrease RAM"
                else if rec > cur and util <> null and util >= MEM_MIN_UPSIZE_UTIL_PCT   then "Increase RAM"
                else "Keep RAM",
        type text),

    AddFinal = Table.AddColumn(AddAction, "Mem_Recommended_MiB",
        each let cur = SafeGetField(_, "Mem_Size_MiB"),
                 raw = _[Mem_Recommended_MiB_Raw],
                 act = _[Mem_Action]
             in if act = "Keep RAM" then cur else if act = "No data" then null else raw,
        type number),

    // -------------------------------------------------------------------------
    // Filter to actionable recommendations only
    // -------------------------------------------------------------------------
    Actionable = Table.SelectRows(AddFinal, each
        _[Mem_Action] = "Decrease RAM" or _[Mem_Action] = "Increase RAM"),

    // -------------------------------------------------------------------------
    // Output
    // -------------------------------------------------------------------------
    Output = SafeSelectColumns(Actionable, {
        "VM", "Mem_Size_MiB",
        "Mem_Active_MiB", "Mem_Consumed_MiB", "Mem_Basis_MiB", "Mem_Util_Pct",
        "Mem_Recommended_MiB_Raw", "Mem_Recommended_MiB", "Mem_Action"
    })
in
    Output
