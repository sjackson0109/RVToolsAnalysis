' =============================================================================
' RVTools Analysis VBA Code
' =============================================================================
' Instructions:
' 1. Create a new Excel workbook
' 2. Add an "Instructions" sheet
' 3. Insert two ActiveX buttons:
'    - Button1: Caption = "Browse for RVTools File"
'    - Button2: Caption = "Refresh All Analysis"
' 4. Copy this VBA code into the Instructions sheet code module
' 5. Import all .m files as Power Query connections
' =============================================================================

Option Explicit

' Global variable to store the selected file path
Dim SelectedFilePath As String

' =============================================================================
' Button 1: Browse for RVTools XLSX File
' =============================================================================
Private Sub CommandButton1_Click()
    Dim fd As FileDialog
    Dim fileSelected As Boolean
    
    ' Create file dialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    With fd
        .Title = "Select RVTools Export File"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsx;*.xlsm;*.xls"
        .FilterIndex = 1
        .AllowMultiSelect = False
        .InitialFileName = "RVTools_Export"
    End With
    
    ' Show dialog and get selection
    fileSelected = fd.Show
    
    If fileSelected Then
        SelectedFilePath = fd.SelectedItems(1)
        
        ' Update the file path display
        Range("B5").Value = SelectedFilePath  ' Adjust cell reference as needed
        Range("B6").Value = "File selected: " & Format(Now, "yyyy-mm-dd hh:mm:ss")
        
        ' Update the data source connection
        Call UpdateDataSource(SelectedFilePath)
        
        MsgBox "RVTools file selected successfully!" & vbCrLf & _
               "Path: " & SelectedFilePath & vbCrLf & vbCrLf & _
               "Click 'Refresh All Analysis' to update the data.", _
               vbInformation, "File Selected"
    Else
        MsgBox "No file selected.", vbExclamation, "Operation Cancelled"
    End If
    
    Set fd = Nothing
End Sub

' =============================================================================
' Button 2: Refresh All Analysis Queries
' =============================================================================
Private Sub CommandButton2_Click()
    Dim startTime As Double
    Dim conn As WorkbookConnection
    Dim queryCount As Integer
    Dim errorCount As Integer
    Dim errorList As String
    
    ' Check if a file has been selected
    If SelectedFilePath = "" Or Range("B5").Value = "" Then
        MsgBox "Please select an RVTools file first using the 'Browse' button.", _
               vbExclamation, "No File Selected"
        Exit Sub
    End If
    
    ' Disable screen updating for performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    startTime = Timer
    queryCount = 0
    errorCount = 0
    errorList = ""
    
    ' Update status
    Range("B7").Value = "Refreshing queries..."
    Range("B8").Value = "Started: " & Format(Now, "yyyy-mm-dd hh:mm:ss")
    
    ' Refresh all Power Query connections
    For Each conn In ThisWorkbook.Connections
        If conn.Type = xlConnectionTypeOLEDB Or conn.Type = xlConnectionTypeODBC Then
            On Error GoTo ErrorHandler
            
            queryCount = queryCount + 1
            Range("B7").Value = "Refreshing: " & conn.Name
            DoEvents
            
            conn.Refresh
            GoTo NextConnection
            
ErrorHandler:
            errorCount = errorCount + 1
            errorList = errorList & conn.Name & ": " & Err.Description & vbCrLf
            Resume NextConnection
            
NextConnection:
        End If
    Next conn
    
    ' Re-enable screen updating
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    ' Update status with results
    Range("B7").Value = "Refresh completed"
    Range("B8").Value = "Completed: " & Format(Now, "yyyy-mm-dd hh:mm:ss")
    Range("B9").Value = "Duration: " & Format((Timer - startTime) / 60, "0.0") & " minutes"
    Range("B10").Value = "Queries processed: " & queryCount
    
    If errorCount = 0 Then
        MsgBox "All analysis queries refreshed successfully!" & vbCrLf & _
               "Queries processed: " & queryCount & vbCrLf & _
               "Duration: " & Format((Timer - startTime) / 60, "0.0") & " minutes", _
               vbInformation, "Refresh Completed"
    Else
        MsgBox "Refresh completed with " & errorCount & " errors:" & vbCrLf & vbCrLf & _
               errorList & vbCrLf & _
               "Successful queries: " & (queryCount - errorCount) & "/" & queryCount, _
               vbExclamation, "Refresh Completed with Errors"
    End If
End Sub

' =============================================================================
' Helper Function: Update Data Source Connection
' =============================================================================
Private Sub UpdateDataSource(filePath As String)
    Dim conn As WorkbookConnection
    Dim newConnectionString As String
    Dim updated As Boolean
    
    updated = False
    
    ' Create new connection string
    newConnectionString = "OLEDB;Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & filePath & _
                         ";Extended Properties=""Excel 12.0;HDR=YES"";"
    
    ' Find and update the RVTools data source connection
    For Each conn In ThisWorkbook.Connections
        If InStr(1, conn.Name, "RVTools", vbTextCompare) > 0 Or _
           InStr(1, conn.Name, "DATASOURCE", vbTextCompare) > 0 Or _
           InStr(1, conn.Description, "RVTools", vbTextCompare) > 0 Then
            
            On Error GoTo UpdateError
            conn.OLEDBConnection.Connection = newConnectionString
            updated = True
            GoTo NextUpdate
            
UpdateError:
            ' If OLEDB update fails, try ODBCConnection
            On Error GoTo NextUpdate
            If Not conn.ODBCConnection Is Nothing Then
                conn.ODBCConnection.Connection = "DRIVER={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};DBQ=" & filePath
                updated = True
            End If
            
NextUpdate:
            On Error GoTo 0
        End If
    Next conn
    
    If Not updated Then
        MsgBox "Warning: Could not automatically update data source connection. " & _
               "You may need to manually update the connection in Power Query.", _
               vbExclamation, "Connection Update Warning"
    End If
End Sub

' =============================================================================
' Helper Function: Initialize Status Display
' =============================================================================
Private Sub Workbook_Open()
    ' Set up initial status display
    Range("A5").Value = "Selected File:"
    Range("A6").Value = "Status:"
    Range("A7").Value = "Progress:"
    Range("A8").Value = "Last Refresh:"
    Range("A9").Value = "Duration:"
    Range("A10").Value = "Queries:"
    
    Range("B5:B10").ClearContents
    Range("B6").Value = "Ready - Select RVTools file to begin"
End Sub

' =============================================================================
' Alternative Refresh Method (if the above doesn't work)
' =============================================================================
Private Sub RefreshAllQueriesAlternative()
    Dim qt As QueryTable
    Dim ws As Worksheet
    Dim lo As ListObject
    
    ' Method 1: Refresh Query Tables
    For Each ws In ThisWorkbook.Worksheets
        For Each qt In ws.QueryTables
            On Error Resume Next
            qt.Refresh BackgroundQuery:=False
            On Error GoTo 0
        Next qt
    Next ws
    
    ' Method 2: Refresh List Objects (Power Query tables)
    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If lo.SourceType = xlSrcQuery Then
                On Error Resume Next
                lo.QueryTable.Refresh BackgroundQuery:=False
                On Error GoTo 0
            End If
        Next lo
    Next ws
    
    ' Method 3: Refresh all connections
    ThisWorkbook.RefreshAll
End Sub

' =============================================================================
' Form Control Button Handlers (for buttons created by PowerShell script)
' =============================================================================
Sub Button1_Click()
    Call CommandButton1_Click
End Sub

Sub Button2_Click()
    Call CommandButton2_Click
End Sub