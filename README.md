# RVTools VM Right-Sizing Recommendations

> **Author:** Simon Jackson ([@sjackson0109](https://github.com/sjackson0109))  
> **Created:** 2026-03-03  
> **Last Modified:** 2026-03-07

---

## What is RVTools Analysis?

RVTools Analysis is an automated VM right-sizing solution that processes [RVTools](https://www.robware.net/rvtools/) exports to generate actionable infrastructure optimization recommendations. The solution identifies over-allocated and under-allocated virtual machines across CPU, memory, disk, and other resources.

## How It Works

The analysis uses **Power Query (M) language** to process RVTools data through independent, self-contained queries:

| Analysis Module | Target Resources | Identifies |
|---|---|---|
| **CPU Analysis** | vCPU allocation | Over/under-provisioned virtual CPUs |
| **Memory Analysis** | RAM allocation | Memory waste and shortages |
| **Disk Analysis** | Storage utilization | Low disk space requiring expansion |
| **GPU Analysis** | Video memory | Unused video RAM allocation |
| **Network Analysis** | Virtual NICs | Excess or disconnected network interfaces |
| **Media Analysis** | Virtual media | Connected CD/DVD drives and ISO files |

## What It Does

The solution automatically:

- **Analyzes** current VM resource utilization from RVTools snapshots
- **Calculates** optimal resource allocation based on actual usage patterns
- **Recommends** specific increases, decreases, or maintenance actions
- **Prioritizes** recommendations by impact and resource savings
- **Excludes** VMs that are already optimally sized ("Keep" recommendations)

## Key Features

- **Safety-First Approach**: Built-in 50% downsize caps prevent over-aggressive reductions
- **Flexible Implementation**: Works with Excel Power Query or Power BI Desktop
- **Threshold-Based Gating**: Only recommends changes when utilization crosses defined thresholds
- **Point-in-Time Compensation**: Headroom calculations account for snapshot data limitations
- **Environment Filtering**: Exclude templates, powered-off VMs, or specific naming patterns

> **Technical Note**: CPU analysis uses RVTools "Overall" MHz data (actual demand) rather than "Max" (entitlement), providing more accurate utilization calculations. All recommendations include configurable headroom percentages to ensure performance stability.

---

## Prerequisites

| Requirement | Details |
|---|---|
| **Excel** | Microsoft Excel 2016+ or Microsoft 365 (Power Query built-in) |
| **RVTools** | v4.x — export using *File → Export All to xlsx* |
| **RVTools sheets required** | `vInfo`, `vCPU`, `vMemory`, `vDisk`, `vPartition`, `vHost` |

---

## Implementation Options

Choose the implementation that best fits your environment:

| Implementation | Best For | See Instructions |
|---|---|---|
| **Excel with Power Query** | Simple analysis, Excel-familiar teams | [Basic Setup Instructions](Instructions.md#excel-power-query-setup) |
| **Excel with VBA Interface** | User-friendly interface, automated workflows | [Advanced Excel Setup](Instructions.md#excel-with-vba-interface) |
| **Power BI Desktop** | Advanced visualization, enterprise reporting | [Power BI Setup Instructions](Instructions.md#power-bi-desktop-setup) |

---

## Quick Start

1. **Export RVTools Data**: Use "File → Export all to xlsx" from RVTools
2. **Choose Implementation**: Select Excel or Power BI based on your needs
3. **Follow Setup**: See [detailed setup instructions](Instructions.md)
4. **Configure Analysis**: Set file paths and thresholds in the queries
5. **Run Analysis**: Refresh to generate recommendations

📋 **[View Complete Setup Instructions →](Instructions.md)**

---

## Configuration Reference

All settings are plain constants at the top of each `.m` file. No external table needed.

### Required (all queries)

| Constant | Default | Description |
|---|---|---|
| `RVTOOLS_FILE_PATH` | `"C:\Path\To\RVTools_export.xlsx"` | Full file path to the RVTools export |

### Scope Filters (all queries)

| Constant | Default | Description |
|---|---|---|
| `INCLUDE_POWERED_OFF` | `false` | Include powered-off VMs in analysis |
| `INCLUDE_TEMPLATES` | `false` | Include VM templates in analysis |
| `EXCLUDE_VM_NAME_CONTAINS` | `""` | Comma-separated substrings — VMs whose name contains any of these are excluded |
| `EXCLUDE_FOLDER_CONTAINS` | `""` | Comma-separated substrings — VMs in folders matching any of these are excluded |
| `EXCLUDE_RESOURCE_GROUP_CONTAINS` | `""` | Comma-separated substrings — exclusion by resource group |

### CPU Sizing (RightSizing-CPU.m)

| Constant | Default | Description |
|---|---|---|
| `CPU_TARGET_UTIL_PCT` | `70` | Target CPU utilisation % the recommended vCPU count should achieve |
| `CPU_HEADROOM_PCT` | `20` | Additional headroom % added on top of observed demand before sizing |
| `CPU_MIN_VCPU` | `2` | Minimum vCPU count — recommendations will never go below this |
| `CPU_MAX_DOWNSIZE_UTIL_PCT` | `50` | Only recommend a vCPU decrease if current utilisation is below this % |
| `CPU_MIN_UPSIZE_UTIL_PCT` | `85` | Only recommend a vCPU increase if current utilisation is at or above this % |
| `CPU_MAX_INCREASE_PCT` | `25` | Maximum % increase in vCPU count per sizing pass |

### Memory Sizing (RightSizing-Memory.m)

| Constant | Default | Description |
|---|---|---|
| `MEM_BASIS` | `"ACTIVE"` | Which memory metric to use: `"ACTIVE"` or `"CONSUMED"` |
| `MEM_TARGET_UTIL_PCT` | `60` | Target memory utilisation % the recommended RAM should achieve |
| `MEM_HEADROOM_PCT` | `20` | Additional headroom % added on top of observed usage before sizing |
| `MEM_MIN_MIB` | `2048` | Minimum RAM in MiB — recommendations will never go below this |
| `MEM_STEP_MIB` | `512` | Round recommended RAM up to the nearest multiple of this value |
| `MEM_MAX_DOWNSIZE_UTIL_PCT` | `35` | Only recommend a RAM decrease if current utilisation is below this % |
| `MEM_MIN_UPSIZE_UTIL_PCT` | `75` | Only recommend a RAM increase if current utilisation is at or above this % |

### Disk (RightSizing-Disk.m)

| Constant | Default | Description |
|---|---|---|
| `DISK_MIN_FREE_PCT` | `20` | Flag a VM for disk expansion if any partition has less than this % free |

---

## Output Columns

### RightSizing-CPU.m

Only VMs requiring a CPU change are shown.

| Column | Description |
|---|---|
| `VM` | Virtual machine name |
| `CPUs` | Current vCPU count |
| `CPU_Demand_MHz` | Current CPU demand (Overall MHz from vCPU) |
| `Host_Speed_MHz` | Host CPU clock speed in MHz |
| `CPU_Demand_Cores` | Effective cores in use (Demand MHz ÷ Host Speed) |
| `CPU_Util_Pct` | Current utilisation % (Demand Cores ÷ vCPUs × 100) |
| `CPU_Recommended_vCPU_Raw` | Calculated recommendation (before action gating) |
| `CPU_Recommended_vCPU` | Final recommendation |
| `CPU_Action` | `Decrease vCPU` or `Increase vCPU` |

### RightSizing-Memory.m

Only VMs requiring a RAM change are shown.

| Column | Description |
|---|---|
| `VM` | Virtual machine name |
| `Mem_Size_MiB` | Current provisioned RAM in MiB |
| `Mem_Active_MiB` | Active guest memory in MiB |
| `Mem_Consumed_MiB` | Consumed memory in MiB |
| `Mem_Basis_MiB` | The memory metric used for sizing (Active or Consumed) |
| `Mem_Util_Pct` | Basis as % of provisioned RAM |
| `Mem_Recommended_MiB_Raw` | Calculated recommendation (before action gating) |
| `Mem_Recommended_MiB` | Final recommendation |
| `Mem_Action` | `Decrease RAM` or `Increase RAM` |

### RightSizing-Disk.m

Only VMs with low free space are shown.

| Column | Description |
|---|---|
| `VM` | Virtual machine name |
| `Disk_Total_Capacity_MiB` | Total provisioned disk capacity across all virtual disks |
| `Partition_Min_Free_Pct` | Lowest free % across all partitions |
| `Partition_Min_Free_MiB` | Lowest free MiB across all partitions |
| `Disk_FreeSpace_Flag` | `Low free space` |
| `Disk_Shortfall_MiB` | Additional MiB needed to bring all partitions up to threshold |
| `Disk_Action` | `Expand disk` |

---

## Safety Caps

Both CPU and Memory include a **50 % per-step downsize cap** — the recommended allocation will never be less than half the current allocation in a single pass. This prevents aggressive cuts from point-in-time snapshot data. Run the analysis again after applying changes to continue right-sizing iteratively.

---

## How Sizing is Calculated

### CPU

$$
\text{RecommendedvCPU} = \left\lceil \frac{\text{DemandCores} \times (1 + \text{HeadroomPct}/100)}{\text{TargetUtilPct}/100} \right\rceil
$$

Clamped to a minimum of `CPU_MIN_VCPU` and a maximum increase of `CPU_MAX_INCREASE_PCT`. A decrease is only recommended when utilisation falls below `CPU_MAX_DOWNSIZE_UTIL_PCT`; an increase only when at or above `CPU_MIN_UPSIZE_UTIL_PCT`.

### Memory

$$
\text{RawMiB} = \frac{\text{BasisMiB} \times (1 + \text{HeadroomPct}/100)}{\text{TargetUtilPct}/100}
$$

Rounded up to the nearest `MEM_STEP_MIB` and clamped to a minimum of `MEM_MIN_MIB`. Both downsizes and upsizes are threshold-gated. A 50 % per-step floor prevents cutting more than half the current allocation.

### Disk

$$
\text{ShortfallMiB} = \sum_{\text{partitions}} \max\left(0, \text{CapacityMiB} \times \frac{\text{DISK\_MIN\_FREE\_PCT}}{100} - \text{FreeSpaceMiB}\right)
$$

VMs are flagged for expansion when any partition's free % drops below `DISK_MIN_FREE_PCT`.

---

## Contributing

Pull requests and issues are welcome. Please open an issue first to discuss significant changes.

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/my-improvement`)
3. Commit your changes (`git commit -m 'Add my improvement'`)
4. Push to the branch (`git push origin feature/my-improvement`)
5. Open a Pull Request
