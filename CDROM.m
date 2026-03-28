// =============================================================================
// RVTools VM Right-Sizing — CDROM Analysis
// =============================================================================
// Author:        Simon Jackson (@sjackson0109)
// Created:       2026-03-10
//
// Description:
//   Standalone CDROM optimization query. Reads sheet data from the
//   "RVTools_Source" connector query and analyses vFloppy for connected
//   CD/DVD drives and floppy devices.
//
//   Identifies VMs with connected but unnecessary CD/DVD drives or floppy
//   devices that could be disconnected to improve performance and reduce
//   attack surface.
//
//   Outputs only VMs with connected removable media that should be reviewed.
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

    // CDROM analysis settings
    FLAG_CONNECTED_CDROMS    = true,    // Flag VMs with connected CD/DVD drives
    FLAG_CONNECTED_FLOPPIES  = true,    // Flag VMs with connected floppy drives
    FLAG_BACKED_CDROMS       = true,    // Flag CD/DVD drives with backing files
    FLAG_BACKED_FLOPPIES     = true,    // Flag floppy drives with backing files

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

    vFloppy = SafeTransformColumnTypes(GetSheet("vFloppy"), {
        {"VM", type text}, {"Label", type text}, {"Connected", type text},
        {"Type", type text}, {"Backing", type text}, {"StartConnected", type text}
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
    // CDROM/Floppy analysis
    // -------------------------------------------------------------------------
    // Add device type categorization
    vFloppyAnalyzed = Table.AddColumn(vFloppy, "Device_Category",
        each let 
            deviceType = try Text.Lower(Text.From(_[Type])) otherwise "",
            label = try Text.Lower(Text.From(_[Label])) otherwise ""
             in if Text.Contains(deviceType, "cdrom") or Text.Contains(deviceType, "cd") or 
                   Text.Contains(deviceType, "dvd") or Text.Contains(label, "cd") or 
                   Text.Contains(label, "dvd") then "CDROM"
                else "Floppy",
        type text),

    // Aggregate findings per VM
    DeviceCounts = Table.Group(vFloppyAnalyzed, {"VM"}, {
        {"CDROM_Count", each
            List.Count(List.Select(Table.Column(_, "Device_Category"), each _ = "CDROM"))
            , type number},
        {"Floppy_Count", each
            List.Count(List.Select(Table.Column(_, "Device_Category"), each _ = "Floppy"))
            , type number},
        {"Connected_CDROM_Count", each
            List.Count(List.Select(Table.ToRecords(_),
                each try ([Device_Category] = "CDROM" and 
                         (Text.Lower([Connected] ?? "") = "true" or Text.Lower([StartConnected] ?? "") = "true")) 
                         otherwise false))
            , type number},
        {"Connected_Floppy_Count", each
            List.Count(List.Select(Table.ToRecords(_),
                each try ([Device_Category] = "Floppy" and 
                         (Text.Lower([Connected] ?? "") = "true" or Text.Lower([StartConnected] ?? "") = "true")) 
                         otherwise false))
            , type number},
        {"Backed_CDROM_Count", each
            List.Count(List.Select(Table.ToRecords(_),
                each try ([Device_Category] = "CDROM" and 
                         ([Backing] <> null and [Backing] <> "")) 
                         otherwise false))
            , type number},
        {"Backed_Floppy_Count", each
            List.Count(List.Select(Table.ToRecords(_),
                each try ([Device_Category] = "Floppy" and 
                         ([Backing] <> null and [Backing] <> "")) 
                         otherwise false))
            , type number}
    }),

    // -------------------------------------------------------------------------
    // Join with filtered VM list
    // -------------------------------------------------------------------------
    J1 = SafeNestedJoin(Filtered, {"VM"}, DeviceCounts, {"VM"}, "_DEV", JoinKind.LeftOuter),
    J2 = SafeExpandTableColumn(J1, "_DEV", {
        "CDROM_Count", "Floppy_Count", "Connected_CDROM_Count", 
        "Connected_Floppy_Count", "Backed_CDROM_Count", "Backed_Floppy_Count"
    }, {
        "CDROM_Count", "Floppy_Count", "Connected_CDROM_Count", 
        "Connected_Floppy_Count", "Backed_CDROM_Count", "Backed_Floppy_Count"
    }),

    // Fill nulls with 0
    CleanedData = Table.TransformColumns(J2, {
        {"CDROM_Count", each _ ?? 0, type number},
        {"Floppy_Count", each _ ?? 0, type number},
        {"Connected_CDROM_Count", each _ ?? 0, type number},
        {"Connected_Floppy_Count", each _ ?? 0, type number},
        {"Backed_CDROM_Count", each _ ?? 0, type number},
        {"Backed_Floppy_Count", each _ ?? 0, type number}
    }),

    // -------------------------------------------------------------------------
    // Flag and action logic
    // -------------------------------------------------------------------------
    AddFlag = Table.AddColumn(CleanedData, "CDROM_Flag",
        each let 
            connectedCDs = _[Connected_CDROM_Count],
            connectedFloppies = _[Connected_Floppy_Count],
            backedCDs = _[Backed_CDROM_Count],
            backedFloppies = _[Backed_Floppy_Count]
             in if (FLAG_CONNECTED_CDROMS and connectedCDs > 0) then "Connected CDROM"
                else if (FLAG_CONNECTED_FLOPPIES and connectedFloppies > 0) then "Connected floppy"
                else if (FLAG_BACKED_CDROMS and backedCDs > 0) then "Backed CDROM"
                else if (FLAG_BACKED_FLOPPIES and backedFloppies > 0) then "Backed floppy"
                else "No issues",
        type text),

    AddAction = Table.AddColumn(AddFlag, "CDROM_Action",
        each let flag = _[CDROM_Flag]
             in if flag = "Connected CDROM" then "Disconnect CDROM"
                else if flag = "Connected floppy" then "Disconnect floppy"
                else if flag = "Backed CDROM" then "Remove CDROM backing"
                else if flag = "Backed floppy" then "Remove floppy backing"
                else "No action",
        type text),

    // -------------------------------------------------------------------------
    // Filter to actionable recommendations only
    // -------------------------------------------------------------------------
    Actionable = Table.SelectRows(AddAction, each _[CDROM_Flag] <> "No issues"),

    // -------------------------------------------------------------------------
    // Output
    // -------------------------------------------------------------------------
    Output = SafeSelectColumns(Actionable, {
        "VM", "CDROM_Count", "Floppy_Count", 
        "Connected_CDROM_Count", "Connected_Floppy_Count",
        "Backed_CDROM_Count", "Backed_Floppy_Count",
        "CDROM_Flag", "CDROM_Action"
    })
in
    Output