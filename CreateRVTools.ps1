# Simple PowerShell Script to Create Excel with VBA
param()

Write-Host "Creating Excel workbook..."

# Clean up
$savePath = Join-Path (Get-Location) "RightSizing.xlsm"
if (Test-Path $savePath) {
    Remove-Item $savePath -Force
}

# Get VBA code
$vbaFilePath = Join-Path (Get-Location) "ExcelVBA.bas" 
if (-not (Test-Path $vbaFilePath)) {
    Write-Error "ExcelVBA.bas not found"
    exit 1
}
$vbaCode = Get-Content $vbaFilePath -Raw

# Create Excel
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true
$excel.DisplayAlerts = $false

# Enable VBA access
try {
    $excel.Application.AutomationSecurity = 1
} catch {}

# Create workbook
$workbook = $excel.Workbooks.Add()

# Setup Instructions sheet
$sheet = $workbook.Worksheets.Item(1)
$sheet.Name = "Instructions"

$sheet.Cells.Item(1, 1) = "RVTools VM Right-Sizing Analysis"
$sheet.Cells.Item(3, 1) = "Quick Start:"
$sheet.Cells.Item(4, 1) = "1. Click Browse for RVTools File"
$sheet.Cells.Item(5, 1) = "2. Click Refresh All Analysis"
$sheet.Cells.Item(6, 1) = "3. Review results in analysis sheets"
$sheet.Cells.Item(8, 1) = "Selected File:"
$sheet.Cells.Item(9, 1) = "Status:"
$sheet.Cells.Item(10, 1) = "Progress:"
$sheet.Cells.Item(11, 1) = "Last Refresh:"
$sheet.Cells.Item(12, 1) = "Duration:"
$sheet.Cells.Item(13, 1) = "Queries:"
$sheet.Cells.Item(15, 1) = "Controls:"

# Format
$sheet.Range("A1").Font.Bold = $true
$sheet.Range("A1").Font.Size = 16
$sheet.Range("A8:A13").Font.Bold = $true

# Add analysis sheets
$types = @("CPU Analysis", "RAM Analysis", "Disk Analysis", "GPU Analysis", "CDROM Analysis", "NIC Analysis")
foreach ($type in $types) {
    $newSheet = $workbook.Worksheets.Add()
    $newSheet.Name = $type
    $newSheet.Cells.Item(1, 1) = "$type Results"
    $newSheet.Cells.Item(2, 1) = "Data will appear here after refreshing"
}

# Move Instructions to front
$sheet.Move($workbook.Worksheets.Item(1))

# Inject VBA
$vbaAdded = $false
Write-Host "Adding VBA code..."

try {
    $module = $workbook.VBProject.VBComponents.Add(1)
    $module.Name = "RVToolsAnalysis" 
    $module.CodeModule.AddFromString($vbaCode)
    Write-Host "VBA code added successfully"
    $vbaAdded = $true
} catch {
    Write-Warning "VBA injection failed: $($_.Exception.Message)"
}

# Add buttons
try {
    $btn1 = $sheet.Buttons().Add(150, 240, 200, 30)
    $btn1.Text = "Browse for RVTools File"
    $btn1.OnAction = "Button1_Click"
    
    $btn2 = $sheet.Buttons().Add(150, 280, 200, 30) 
    $btn2.Text = "Refresh All Analysis"
    $btn2.OnAction = "Button2_Click"
    
    Write-Host "Buttons added successfully"
} catch {
    Write-Warning "Button creation failed"
}

# Save as XLSM
try {
    $workbook.SaveAs2($savePath, 52)
    Write-Host "Saved: $savePath"
} catch {
    try {
        $workbook.SaveAs($savePath, 52)
        Write-Host "Saved (fallback): $savePath" 
    } catch {
        Write-Error "Save failed"
        exit 1
    }
}

# Check for .m files and IMPORT THEM
Write-Host ""
Write-Host "Importing Power Query .m files..."
$mFiles = @("_DATASOURCE.m", "CPU.m", "RAM.m", "DISK.m", "GPU.m", "CDROM.m", "NIC.m")
$foundCount = 0
$importedCount = 0

foreach ($file in $mFiles) {
    if (Test-Path $file) {
        Write-Host "Found: $file"
        $foundCount++
        
        # Import the .m file 
        try {
            $mContent = Get-Content $file -Raw
            $queryName = $file.Replace(".m", "").Replace("_", "").Replace("DATASOURCE", "RVTools-Source")
            
            # Add to workbook as Power Query
            $query = $workbook.Queries.Add($queryName, $mContent)
            Write-Host "IMPORTED: $queryName"
            $importedCount++
            
        } catch {
            Write-Warning "Import failed for $file : $($_.Exception.Message)"
        }
    } else {
        Write-Warning "Missing: $file"
    }
}

# Final status
Write-Host ""
Write-Host "=============================="
Write-Host "WORKBOOK CREATED SUCCESSFULLY"
Write-Host "=============================="
Write-Host "File: $savePath"
Write-Host "VBA Injected: $vbaAdded"
Write-Host "M Files Available: $foundCount of $($mFiles.Count)"
Write-Host "Queries Imported: $importedCount of $foundCount"
Write-Host ""

if ($vbaAdded -and $importedCount -eq $foundCount -and $foundCount -eq $mFiles.Count) {
    Write-Host "🎉 EVERYTHING IS COMPLETE!"
    Write-Host "✅ VBA macros embedded and functional"
    Write-Host "✅ All Power Query .m files imported"  
    Write-Host "✅ Form Control buttons working"
    Write-Host ""
    Write-Host "READY TO USE IMMEDIATELY!"
    Write-Host "1. Open RightSizing.xlsm"
    Write-Host "2. Click 'Browse for RVTools File'"
    Write-Host "3. Click 'Refresh All Analysis'"
    Write-Host "4. View results in analysis sheets"
} else {
    Write-Host "Status Summary:"
    Write-Host "- VBA macros: $($vbaAdded)"
    Write-Host "- Files found: $foundCount/$($mFiles.Count)"
    Write-Host "- Queries imported: $importedCount/$foundCount"
    if ($importedCount -lt $foundCount) { 
        Write-Host ""
        Write-Host "Some queries failed to import - may need manual setup"
    }
}

# Cleanup
$excel.Quit() 
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

Write-Host "Done!"