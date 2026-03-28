# RVTools VM Right-Sizing Analysis - Setup Instructions

## Implementation Options

This guide covers three implementation approaches for RVTools analysis. Choose the method that best fits your environment:

### 1. Excel with Power Query (Basic)
- Simple setup using Excel's built-in Power Query
- Manual file path configuration
- Best for: Small teams, simple analysis

### 2. Excel with VBA Interface (Advanced)  
- User-friendly Excel workbook with buttons and automation
- File browser and one-click refresh
- Best for: Regular use, non-technical users

### 3. Power BI Desktop
- Enterprise reporting with advanced visualizations
- Best for: Large environments, executive dashboards

---

## Prerequisites

**Software Requirements:**
- **RVTools** (v4.x or later) - Download from [robware.net](https://www.robware.net/rvtools/)
- **Excel 2016+** or **Microsoft 365** (for Excel implementations)
- **Power BI Desktop** (for Power BI implementation)
- **VMware vCenter access** with read permissions

**RVTools Permissions Required:**
- Read-only access to vCenter Server
- VM and Host inventory viewing permissions  
- Performance data access (recommended)

---

# Excel Power Query Setup

## Step 1: Export RVTools Data

### 1.1 Connect RVTools to vCenter
1. Launch RVTools
2. Click **File → Connect**
3. Enter your vCenter Server details:
   - **vCenter Server**: `vcenter.yourdomain.com`
   - **Username**: `domain\username` or `username@domain.com`
   - **Password**: Your vCenter password
4. Click **Connect**

### 1.2 Export RVTools Data
1. Wait for data collection to complete (may take 10-30 minutes for large environments)
2. Click **File → Export all to Excel**  
3. Choose save location and filename (e.g., `RVTools_export_all_2026-03-25.xlsx`)
4. Click **Save**
5. **Verify** the exported file contains required worksheets: `vInfo`, `vCPU`, `vMemory`, `vDisk`, `vPartition`, `vHost`

## Step 2: Create Analysis Workbook

1. **Create a new Excel workbook** (or use the RVTools export file itself)
2. **Save the workbook** where you'll store your analysis results

## Step 3: Configure Analysis Queries  

### 3.1 Edit Configuration Constants

Open the `.m` file(s) you want to use in a text editor. Each file has a **CONFIGURATION** section at the top. Update:

**Required Settings:**
- **`RVTOOLS_FILE_PATH`** — Full path to your RVTools export  
  Example: `"C:\Exports\RVTools_export_all_2026-03-25.xlsx"`

**Optional Scope Filters:**
- **`INCLUDE_POWERED_OFF`** — Set to `true` to include powered-off VMs
- **`INCLUDE_TEMPLATES`** — Set to `true` to include VM templates  
- **`EXCLUDE_VM_NAME_CONTAINS`** — Comma-separated strings (e.g., `"backup,test,dev"`)
- **`EXCLUDE_FOLDER_CONTAINS`** — Exclude VMs in matching folders
- **`EXCLUDE_RESOURCE_GROUP_CONTAINS`** — Exclude by resource group

## Step 4: Import Queries into Excel

For each analysis module you want to use:

### 4.1 Create New Query
1. In Excel, go to **Data → Get Data → From Other Sources → Blank Query**
2. In the Power Query editor, click **Advanced Editor**
3. **Delete the default content** and paste the full contents of your `.m` file
4. Click **Done**
5. **Rename the query** (e.g., `RightSizing-CPU`, `RightSizing-Memory`)

### 4.2 Load Results  
1. Click **Close & Load To...**
2. Choose **Table** and select target worksheet
3. Click **OK**

### 4.3 Repeat for Additional Modules
Import each analysis module separately:
- `RightSizing-CPU.m` → CPU over/under allocation analysis
- `RightSizing-Memory.m` → Memory optimization analysis  
- `RightSizing-Disk.m` → Disk space analysis

## Step 5: Run Analysis

1. **Refresh data**: Press **Data → Refresh All**
2. **Processing time**: 1-5 minutes depending on environment size
3. **Review results**: Each query outputs only VMs requiring action

---

# Excel with VBA Interface

For a more user-friendly Excel implementation with automated file browsing and refresh buttons, see the advanced setup below.

## Step 1: Create VBA-Enabled Workbook

1. **Create new Excel workbook**
2. **Save as**: `RVTools-Analysis.xlsm` (macro-enabled format)  
3. **Rename Sheet1** to "Instructions"

## Step 2: Set Up Instructions Sheet

Create the following layout on the Instructions sheet:

| Cell | Content |
|---|---|
| A1 | `RVTools VM Right-Sizing Analysis` |
| A3 | `Quick Start:` |
| A5 | `Selected File:` |
| B5 | (Will display selected file path) |
| A6 | `Status:` |
| B6 | (Status updates) |
| A7 | `Progress:` |
| B7 | (Progress information) |
| A8 | `Last Refresh:` |
| B8 | (Timestamp) |
| A9 | `Duration:` |
| B9 | (Time taken) |
| A10 | `Queries:` |
| B10 | (Query count) |

## Step 3: Add VBA Code

1. Press **Alt+F11** to open VBA Editor
2. In Project Explorer, right-click **"ThisWorkbook"**
3. Select **Insert → Module**
4. Copy and paste the contents of `ExcelVBA.bas` into the new module
5. Close VBA Editor (**Alt+Q**)

## Step 4: Insert ActiveX Buttons

1. Go to **Developer** tab (if not visible: File → Options → Customize Ribbon → Check Developer)
2. Click **Insert → ActiveX Controls → Command Button**

**First Button (File Browser):**
3. Draw button in cell A12 area
4. Right-click button → **Properties**:
   - **Name**: `CommandButton1`
   - **Caption**: `Browse for RVTools File`
   - **Width**: `150`
   - **Height**: `30`

**Second Button (Refresh):**
5. Draw button in cell A14 area  
6. Right-click button → **Properties**:
   - **Name**: `CommandButton2`  
   - **Caption**: `Refresh All Analysis`
   - **Width**: `150`
   - **Height**: `30`

## Step 5: Import Analysis Queries

Follow the same query import process as the basic Excel setup (Step 4 above), importing each `.m` file as a Power Query.

---

# Power BI Desktop Setup

## Step 1: Export RVTools Data

Follow the same RVTools export process as described in the Excel setup above.

## Step 2: Import RVTools Data into Power BI

### 2.1 Create New Power BI File
1. Launch **Power BI Desktop**
2. Click **File → New** (or use existing file)

### 2.2 Import RVTools Excel Data
1. Click **Get Data → Excel Workbook**
2. Navigate to your RVTools export file and click **Open**
3. In the Navigator window:
   - **Check all available worksheets**: `vInfo`, `vCPU`, `vMemory`, `vDisk`, `vPartition`, `vNetwork`, `vHost`
   - Click **Load** (not Transform)

### 2.3 Create Data Source Query
1. Go to **Home → Transform Data** (Power Query Editor)
2. Right-click in the Queries pane → **New Query → Blank Query**
3. **Rename the query** to: `RVTools-Source`
4. Go to **View → Advanced Editor**
5. Copy and paste the entire contents of `_DATASOURCE.m`
6. **Update the file path** in `RVTOOLS_FILE_PATH` to your Excel file location
7. Click **Done**
8. **Important**: Right-click the query → **Enable Load → Uncheck** (Connection Only)

## Step 3: Import Analysis Queries

### 3.1 Available Analysis Modules

| File | Description | Analysis Focus |
|---|---|---|
| `CPU.m` | CPU optimization | vCPU recommendations, utilization analysis |
| `RAM.m` | Memory optimization | RAM recommendations, usage analysis |
| `DISK.m` | Disk space analysis | Free space percentage, expansion needs |
| `GPU.m` | Video memory optimization | Video RAM reduction opportunities |
| `CDROM.m` | Removable media cleanup | Connected drives, backing files |
| `NIC.m` | Network interface optimization | NIC count, connectivity analysis |

### 3.2 Import Process for Each Query
For each analysis module:

1. In Power Query Editor, click **New Query → Blank Query**
2. **Rename the query** (e.g., "CPU Analysis", "RAM Analysis", "Disk Analysis")
3. Go to **View → Advanced Editor**
4. Copy and paste the entire contents of the corresponding `.m` file
5. Click **Done**
6. Repeat for each analysis module you want to use

## Step 4: Configuration

All configuration is done by editing constants at the top of each `.m` file.

### 4.1 Global Scope Filters (Available in all queries)

```
INCLUDE_POWERED_OFF = false          // Include powered-off VMs?
INCLUDE_TEMPLATES = false            // Include VM templates?
EXCLUDE_VM_NAME_CONTAINS = ""        // e.g., "backup,test,dev"  
EXCLUDE_FOLDER_CONTAINS = ""         // e.g., "archive,templates"
EXCLUDE_RESOURCE_GROUP_CONTAINS = "" // e.g., "test,development"
```

### 4.2 CPU Analysis Settings

```
CPU_TARGET_UTIL_PCT = 70            // Target CPU utilisation percentage
CPU_HEADROOM_PCT = 20                // Additional headroom percentage  
CPU_MIN_VCPU = 2                     // Minimum vCPU count
CPU_MAX_DOWNSIZE_UTIL_PCT = 50       // Only decrease if util < this %
CPU_MIN_UPSIZE_UTIL_PCT = 85         // Only increase if util >= this %
CPU_MAX_INCREASE_PCT = 25            // Maximum % increase per pass
```

### 4.3 Memory Analysis Settings

```
MEM_BASIS = "ACTIVE"                 // "ACTIVE" or "CONSUMED"
MEM_TARGET_UTIL_PCT = 60             // Target memory utilisation percentage
MEM_HEADROOM_PCT = 20                // Additional headroom percentage
MEM_MIN_MIB = 2048                   // Minimum RAM in MiB
MEM_STEP_MIB = 512                   // Round to nearest multiple of this value
MEM_MAX_DOWNSIZE_UTIL_PCT = 35       // Only decrease if util < this %
MEM_MIN_UPSIZE_UTIL_PCT = 75         // Only increase if util >= this %
```

### 4.4 Disk Analysis Settings

```
DISK_MIN_FREE_PCT = 20               // Flag VMs with partitions below this % free
```

## Step 5: Running Analysis

### 5.1 Refresh Data
1. **Close Power Query Editor**: Click **Close & Apply**
2. **Refresh analysis**: Click **Home → Refresh**
3. **Processing time**: Typically 1-5 minutes depending on environment size

### 5.2 Review Results

Each analysis outputs a table containing only VMs that need attention:

**CPU Analysis Output:**
- `VM` — Virtual machine name
- `CPUs` — Current vCPU allocation  
- `CPU_Util_Pct` — Current utilization percentage
- `CPU_Recommended_vCPU` — Recommended vCPU count
- `CPU_Action` — "Decrease vCPU" or "Increase vCPU"

**Memory Analysis Output:**
- `VM` — Virtual machine name
- `Mem_Size_MiB` — Current memory allocation
- `Mem_Util_Pct` — Memory utilization percentage  
- `Mem_Recommended_MiB` — Recommended memory allocation
- `Mem_Action` — "Decrease RAM" or "Increase RAM"

**Disk Analysis Output:**
- `VM` — Virtual machine name
- `Partition_Min_Free_Pct` — Lowest free space percentage
- `Disk_Shortfall_MiB` — Additional space needed
- `Disk_Action` — "Expand disk"

---

# Configuration Reference

All analyses use the same core configuration pattern. Settings are defined as constants at the top of each `.m` file.

## Required Settings (All Queries)

| Constant | Example | Description |
|---|---|---|
| `RVTOOLS_FILE_PATH` | `"C:\Exports\RVTools_export.xlsx"` | Full path to RVTools export file |

## Scope Filters (All Queries)

| Constant | Default | Description |
|---|---|---|
| `INCLUDE_POWERED_OFF` | `false` | Include powered-off VMs in analysis |
| `INCLUDE_TEMPLATES` | `false` | Include VM templates in analysis |
| `EXCLUDE_VM_NAME_CONTAINS` | `""` | Comma-separated substrings to exclude VMs by name |
| `EXCLUDE_FOLDER_CONTAINS` | `""` | Comma-separated substrings to exclude VMs by folder |
| `EXCLUDE_RESOURCE_GROUP_CONTAINS` | `""` | Comma-separated substrings to exclude by resource group |

## CPU Analysis Configuration

| Constant | Default | Description |  
|---|---|---|
| `CPU_TARGET_UTIL_PCT` | `70` | Target CPU utilisation % for recommended vCPU count |
| `CPU_HEADROOM_PCT` | `20` | Additional headroom % added to observed demand |
| `CPU_MIN_VCPU` | `2` | Minimum vCPU count (recommendations never go below) |
| `CPU_MAX_DOWNSIZE_UTIL_PCT` | `50` | Only recommend vCPU decrease if current util < this % |
| `CPU_MIN_UPSIZE_UTIL_PCT` | `85` | Only recommend vCPU increase if current util >= this % |
| `CPU_MAX_INCREASE_PCT` | `25` | Maximum % increase in vCPU count per sizing pass |

## Memory Analysis Configuration

| Constant | Default | Description |
|---|---|---|
| `MEM_BASIS` | `"ACTIVE"` | Memory metric to use: `"ACTIVE"` or `"CONSUMED"` |
| `MEM_TARGET_UTIL_PCT` | `60` | Target memory utilisation % for recommended RAM |
| `MEM_HEADROOM_PCT` | `20` | Additional headroom % added to observed usage |
| `MEM_MIN_MIB` | `2048` | Minimum RAM in MiB (recommendations never go below) |
| `MEM_STEP_MIB` | `512` | Round recommended RAM to nearest multiple of this |
| `MEM_MAX_DOWNSIZE_UTIL_PCT` | `35` | Only recommend RAM decrease if current util < this % |
| `MEM_MIN_UPSIZE_UTIL_PCT` | `75` | Only recommend RAM increase if current util >= this % |

## Disk Analysis Configuration

| Constant | Default | Description |
|---|---|---|
| `DISK_MIN_FREE_PCT` | `20` | Flag VM for disk expansion if any partition < this % free |

---

# Implementation Guidelines

## Step 6: Interpreting Results

### 6.1 Action Priorities

1. **High Impact**: CPU and RAM reductions (immediate resource savings)
2. **Medium Impact**: Disk expansions (prevent outages)  
3. **Low Impact**: Video RAM, CDROM, NIC optimizations (efficiency gains)

### 6.2 Implementation Best Practices

**CPU Downsizing:**
- Start with VMs showing < 50% utilization
- Test thoroughly before production changes
- Monitor performance metrics post-implementation

**Memory Downsizing:**  
- Use built-in 50% reduction cap for safety
- Choose ACTIVE vs CONSUMED basis appropriately  
- Maintain configured buffer above actual usage

**Disk Expansion:**
- Prioritize VMs below 20% free space
- Plan expansions based on shortfall calculations
- Implement regular monitoring schedule

### 6.3 Pre-Implementation Validation

Before making changes:
1. **Cross-reference** with application teams
2. **Check maintenance windows** for planned changes
3. **Verify backup procedures** and rollback plans
4. **Test in development** environments first

## Step 7: Maintenance Schedule  

### 7.1 Regular Refresh Frequency
- **Weekly**: Dynamic/high-change environments
- **Monthly**: Stable environments  
- **Before major changes**: Capacity planning exercises

### 7.2 Updating Analysis
1. Re-export fresh RVTools data
2. Update file path in data source query
3. Refresh all analysis queries
4. Review new recommendations

### 7.3 Seasonal Configuration
- **Adjust thresholds** for known usage patterns
- **Modify headroom** for growth planning
- **Update exclusion filters** for policy compliance

---

# Troubleshooting

## Common Issues & Solutions

### "Column not found" Errors
- **Cause**: RVTools version differences or incomplete export
- **Solution**: Verify all worksheets exported; check for column name changes

### "File not found" Errors  
- **Cause**: Incorrect file path in data source query
- **Solution**: Update file path constant; ensure file accessibility

### Empty/No Results
- **Cause**: Overly restrictive filtering or no optimization opportunities
- **Solution**: Review filter settings; verify source data completeness

### Performance Issues
- **Cause**: Large dataset or complex filtering
- **Solution**: Increase timeout settings; consider environment filtering

## Support Resources

- **RVTools Documentation**: [robware.net](https://www.robware.net/rvtools/)
- **Power BI Community**: [community.powerbi.com](https://community.powerbi.com)
- **Excel Power Query Help**: [Microsoft Support](https://support.microsoft.com/excel)

---

# File Structure Reference

```
RVTools-Analysis/
├── README.md                    # Project overview and quick start
├── Instructions.md              # This file - detailed setup instructions
├── _DATASOURCE.m               # Data source connector (all implementations)
├── CPU.m                       # CPU optimization analysis query
├── RAM.m                       # Memory optimization analysis query  
├── DISK.m                      # Disk space analysis query
├── GPU.m                       # Video memory optimization query
├── CDROM.m                     # Removable media analysis query
├── NIC.m                       # Network interface analysis query
├── ExcelInstructions.txt       # Quick reference for Excel setup
├── ExcelSetup.txt              # Detailed Excel VBA workbook guide
├── ExcelVBA.bas                # VBA code for advanced Excel interface
└── RVTools_Export_[DATE].xlsx  # Your RVTools export (you need to save these here)
```

---

**Created**: March 25, 2026  
**Author**: Simon Jackson ([@sjackson0109](https://github.com/sjackson0109))  
**Version**: 2.0  
**Compatibility**: RVTools 4.x+, Excel 2016+, Power BI Desktop                                                                                                   