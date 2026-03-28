// =============================================================================
// RVTools Source Connector
// =============================================================================
// Author:        Simon Jackson (@sjackson0109)
// Created:       2026-03-07
//
// Description:
//   Shared data source for all RightSizing queries. Opens the RVTools export
//   file once and returns the workbook sheet index.
//
//   Import this as a Power Query named "RVTools_Source" and set it to
//   "Connection Only" (right-click → Load To → Connection Only).
//
//   All RightSizing-*.m queries reference this connector by name.
//
// Configuration:
//   Set RVTOOLS_FILE_PATH to the full path of your RVTools export .xlsx file.
// =============================================================================
let
    // =========================================================================
    // CONFIGURATION
    // =========================================================================

    RVTOOLS_FILE_PATH = "C:\Users\Administrator\RVTools-Analysis\RVTools_export_all_2026-03-03_10.41.56.xlsx",

    // =========================================================================
    // END OF CONFIGURATION
    // =========================================================================

    Source =
        try Excel.Workbook(File.Contents(RVTOOLS_FILE_PATH), null, true)
        otherwise error "Cannot open RVTools file: " & RVTOOLS_FILE_PATH
in
    Source
