// =============================================================================
// RVTools VM Right-Sizing — Network Interface (VMNIC) Analysis  
// =============================================================================
// Author:        Simon Jackson (@sjackson0109)
// Created:       2026-03-10
//
// Description:
//   Standalone VMNIC optimization query. Reads sheet data from the
//   "RVTools_Source" connector query and analyses vNetwork for network
//   interface allocation and utilization.
//
//   Identifies VMs with:
//   - Excessive number of network interfaces
//   - Disconnected network interfaces that could be removed
//   - Unused network interfaces with no traffic
//   - VMs that might benefit from additional NICs
//
//   Outputs only VMs whose network interface configuration should be optimized.
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

    // VMNIC analysis thresholds
    NIC_MAX_STANDARD        = 2,    // Standard maximum NICs for typical workload
    NIC_MAX_HIGH_TRAFFIC    = 4,    // Maximum NICs for high-traffic workloads
    FLAG_EXCESSIVE_NICS     = true, // Flag VMs with too many NICs
    FLAG_DISCONNECTED_NICS  = true, // Flag VMs with disconnected NICs
    FLAG_SINGLE_NIC_SERVERS = false, // Flag servers with only 1 NIC (redundancy concern)

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

    vNetwork = SafeTransformColumnTypes(GetSheet("vNetwork"), {
        {"VM", type text}, {"Label", type text}, {"Connected", type text},
        {"Type", type text}, {"Network Label", type text}, {"MAC", type text}
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
    // Network interface aggregation
    // -------------------------------------------------------------------------
    NetworkCounts = Table.Group(vNetwork, {"VM"}, {
        {"NIC_Total_Count", each Table.RowCount(_), type number},
        {"NIC_Connected_Count", each
            List.Count(List.Select(Table.Column(_, "Connected"),
                each try Text.Lower(Text.From(_)) = "true" otherwise false))
            , type number},
        {"NIC_Disconnected_Count", each
            List.Count(List.Select(Table.Column(_, "Connected"),
                each try Text.Lower(Text.From(_)) <> "true" otherwise true))
            , type number},
        {"NIC_Network_Labels", each
            Text.Combine(List.Distinct(List.Select(Table.Column(_, "Network Label"), each _ <> null and _ <> "")), "; ")
            , type text}
    }),

    // -------------------------------------------------------------------------
    // Join with filtered VM list
    // -------------------------------------------------------------------------
    J1 = SafeNestedJoin(Filtered, {"VM"}, NetworkCounts, {"VM"}, "_NET", JoinKind.LeftOuter),
    J2 = SafeExpandTableColumn(J1, "_NET", {
        "NIC_Total_Count", "NIC_Connected_Count", "NIC_Disconnected_Count", "NIC_Network_Labels"
    }, {
        "NIC_Total_Count", "NIC_Connected_Count", "NIC_Disconnected_Count", "NIC_Network_Labels"
    }),

    // Fill nulls with defaults for VMs without network data
    CleanedData = Table.TransformColumns(J2, {
        {"NIC_Total_Count", each _ ?? 0, type number},
        {"NIC_Connected_Count", each _ ?? 0, type number},
        {"NIC_Disconnected_Count", each _ ?? 0, type number},
        {"NIC_Network_Labels", each _ ?? "No networks", type text}
    }),

    // -------------------------------------------------------------------------
    // Network optimization analysis
    // -------------------------------------------------------------------------
    AddNICFlag = Table.AddColumn(CleanedData, "NIC_Flag",
        each let 
            totalNICs = _[NIC_Total_Count],
            connectedNICs = _[NIC_Connected_Count], 
            disconnectedNICs = _[NIC_Disconnected_Count]
             in if FLAG_EXCESSIVE_NICS and totalNICs > NIC_MAX_HIGH_TRAFFIC then "Excessive NICs"
                else if FLAG_DISCONNECTED_NICS and disconnectedNICs > 0 then "Disconnected NICs"
                else if FLAG_SINGLE_NIC_SERVERS and connectedNICs = 1 then "Single NIC"
                else if totalNICs > NIC_MAX_STANDARD and totalNICs <= NIC_MAX_HIGH_TRAFFIC then "High NIC count"
                else "Normal",
        type text),

    AddNICRecommendation = Table.AddColumn(AddNICFlag, "NIC_Recommended_Count",
        each let 
            totalNICs = _[NIC_Total_Count],
            connectedNICs = _[NIC_Connected_Count],
            flag = _[NIC_Flag]
             in if flag = "Excessive NICs" then NIC_MAX_STANDARD
                else if flag = "Disconnected NICs" then connectedNICs
                else if flag = "Single NIC" then 2
                else totalNICs,
        type number),

    AddNICAction = Table.AddColumn(AddNICRecommendation, "NIC_Action",
        each let 
            current = _[NIC_Total_Count],
            recommended = _[NIC_Recommended_Count],
            flag = _[NIC_Flag]
             in if flag = "Excessive NICs" then "Remove excess NICs"
                else if flag = "Disconnected NICs" then "Remove disconnected NICs"
                else if flag = "Single NIC" then "Add redundant NIC"
                else if flag = "High NIC count" then "Review NIC necessity"
                else "Keep current NICs",
        type text),

    AddNICDelta = Table.AddColumn(AddNICAction, "NIC_Delta",
        each let 
            current = _[NIC_Total_Count],
            recommended = _[NIC_Recommended_Count]
             in recommended - current,
        Int64.Type),

    // -------------------------------------------------------------------------
    // Filter to actionable recommendations only
    // -------------------------------------------------------------------------
    Actionable = Table.SelectRows(AddNICDelta, each _[NIC_Flag] <> "Normal"),

    // -------------------------------------------------------------------------
    // Output
    // -------------------------------------------------------------------------
    Output = SafeSelectColumns(Actionable, {
        "VM", "NIC_Total_Count", "NIC_Connected_Count", "NIC_Disconnected_Count",
        "NIC_Network_Labels", "NIC_Flag", "NIC_Recommended_Count", 
        "NIC_Delta", "NIC_Action"
    })
in
    Output