// =============================================================================
// RVTools VM Right-Sizing — Video RAM Analysis
// =============================================================================
// Author:        Simon Jackson (@sjackson0109)
// Created:       2026-03-10
//
// Description:
//   Standalone Video RAM right-sizing query. Reads sheet data from the
//   "RVTools_Source" connector query and analyses vInfo for video memory allocation.
//
//   Identifies VMs with excessive video memory allocation that could be reduced
//   to free up host memory resources. Most server workloads don't require
//   high video memory allocations.
//
//   Outputs only VMs whose Video RAM allocation should change (Decrease mainly).
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

    // Video RAM thresholds
    VIDEO_RAM_DEFAULT_KB     = 4096,   // Standard video RAM allocation in KB
    VIDEO_RAM_MIN_KB         = 4096,   // Minimum video RAM in KB
    VIDEO_RAM_HIGH_UTIL_KB   = 16384,  // Flag VMs with video RAM above this (16MB)

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

    // -------------------------------------------------------------------------
    // Load & type the sheets we need
    // -------------------------------------------------------------------------
    vInfo = SafeTransformColumnTypes(GetSheet("vInfo"), {
        {"VM", type text}, {"Powerstate", type text}, {"Template", type text},
        {"CPUs", Int64.Type}, {"Memory", Int64.Type}, {"Host", type text},
        {"Folder", type any}, {"Resource Group", type any},
        {"Video Ram KB", type number}
    }),

    // -------------------------------------------------------------------------
    // VM base list + filtering
    // -------------------------------------------------------------------------
    Base = SafeSelectColumns(vInfo,
        {"VM", "Powerstate", "Template", "CPUs", "Memory", "Host", "Folder", "Resource Group", "Video Ram KB"}),

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
    // Video RAM analysis
    // -------------------------------------------------------------------------
    AddVideoRAMAnalysis = Table.AddColumn(Filtered, "Video_RAM_Current_KB",
        each SafeGetField(_, "Video Ram KB") ?? VIDEO_RAM_DEFAULT_KB,
        type number),

    AddVideoRAMFlag = Table.AddColumn(AddVideoRAMAnalysis, "Video_RAM_Flag",
        each let current = _[Video_RAM_Current_KB]
             in if current > VIDEO_RAM_HIGH_UTIL_KB then "High video RAM"
                else if current < VIDEO_RAM_MIN_KB then "Low video RAM"
                else "Normal",
        type text),

    AddVideoRAMRecommendation = Table.AddColumn(AddVideoRAMFlag, "Video_RAM_Recommended_KB",
        each let 
            current = _[Video_RAM_Current_KB],
            flag = _[Video_RAM_Flag]
             in if flag = "High video RAM" then VIDEO_RAM_DEFAULT_KB
                else if flag = "Low video RAM" then VIDEO_RAM_MIN_KB
                else current,
        type number),

    AddVideoRAMAction = Table.AddColumn(AddVideoRAMRecommendation, "Video_RAM_Action",
        each let 
            current = _[Video_RAM_Current_KB],
            recommended = _[Video_RAM_Recommended_KB]
             in if recommended < current then "Decrease video RAM"
                else if recommended > current then "Increase video RAM"
                else "Keep video RAM",
        type text),

    AddVideoRAMSavings = Table.AddColumn(AddVideoRAMAction, "Video_RAM_Savings_KB",
        each let 
            current = _[Video_RAM_Current_KB],
            recommended = _[Video_RAM_Recommended_KB]
             in if current > recommended then current - recommended else 0,
        type number),

    // -------------------------------------------------------------------------
    // Filter to actionable recommendations only
    // -------------------------------------------------------------------------
    Actionable = Table.SelectRows(AddVideoRAMSavings, each
        _[Video_RAM_Action] = "Decrease video RAM" or _[Video_RAM_Action] = "Increase video RAM"),

    // -------------------------------------------------------------------------
    // Output
    // -------------------------------------------------------------------------
    Output = SafeSelectColumns(Actionable, {
        "VM", "Video_RAM_Current_KB", "Video_RAM_Flag",
        "Video_RAM_Recommended_KB", "Video_RAM_Savings_KB", "Video_RAM_Action"
    })
in
    Output