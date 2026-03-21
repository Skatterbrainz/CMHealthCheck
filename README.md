# CMHealthCheck
# CMHealthCheck

ConfigMgr health-check reporting PowerShell module

![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue)
![Platform](https://img.shields.io/badge/platform-Windows-blue)
![License](https://img.shields.io/badge/license-MIT-green)

CMHealthCheck helps you collect Configuration Manager health data and generate report-ready output for operational reviews, troubleshooting, and periodic health checks.

## Overview

The module is built around a two-phase workflow:

1. Data collection from ConfigMgr, SQL, WMI/CIM, registry, and service layers
2. Report generation (Word/HTML) with summary and detailed sections

This approach lets you collect once and render reports multiple ways.

## Features

- Configuration Manager site health data collection
- Summary and detailed report generation workflows
- Word and HTML report export support
- XML-based report definitions and message content
- Helper functions for SQL, services, network, boundaries, collections, and package insights

## Requirements

- PowerShell 5.1 or later
- Access to a Configuration Manager environment
- Required permissions to query SQL/WMI/CIM and remote systems (when applicable)
- Microsoft Word 2013 or later for Word report generation

## Installation

### From PowerShell Gallery

```powershell
Install-Module CMHealthCheck
```

### From GitHub

1. Clone the repository:

```bash
git clone https://github.com/ds0934/CMHealthCheck.git
cd CMHealthCheck
```

2. Import the module:

```powershell
Import-Module ./CMHealthCheck.psd1
```

## Usage

Import the module and inspect available commands:

```powershell
Import-Module CMHealthCheck
Get-Command -Module CMHealthCheck
Get-Help Invoke-CMHealthCheck -Detailed
```

Example end-to-end flow:

```powershell
# Run collection and generate outputs from a config file
Invoke-CMHealthCheck -ConfigFile .\Config.xml -ReportFile .\CMHealthReport.docx

# Collect health-check data
Get-CMHealthCheck -ConfigFile .\Config.xml -OutputPath .\Output

# Summarize prior output
Get-CMHealthCheckSummary -Path .\Output

# Export report content
Export-CMHealthReport -Path .\Output -ReportFile .\CMHealthReport.docx
```

## Function Reference

| Function | Description | Documentation |
|---|---|---|
| `Invoke-CMHealthCheck` | End-to-end workflow for collection and reporting | [Docs/Invoke-CMHealthCheck.md](Docs/Invoke-CMHealthCheck.md) |
| `Get-CMHealthCheck` | Collect health-check data from configured targets | [Docs/Get-CMHealthCheck.md](Docs/Get-CMHealthCheck.md) |
| `Get-CMHealthCheckSummary` | Produce summarized health-check output | [Docs/Get-CMHealthCheckSummary.md](Docs/Get-CMHealthCheckSummary.md) |
| `Export-CMHealthReport` | Export collected data into report output | [Docs/Export-CMHealthReport.md](Docs/Export-CMHealthReport.md) |

## Compatibility Notes

Tested environments include:

- ConfigMgr 2012 R2 through current branch/technical preview releases
- Windows Server 2012, 2012 R2, 2016, 2019, 2022
- SQL Server 2012, 2014, 2016, 2017, 2019, 2022
- Windows 10 and Windows 11
- Office 2013, 2016, 2019, Microsoft 365

## Samples and Docs

- Usage examples are available under [samples](samples)
- Command documentation is available under [Docs](Docs)

## Contributing

Contributions are welcome. You can help by:

- Reporting bugs
- Suggesting improvements
- Submitting pull requests
- Improving documentation

Please open an issue or pull request in this repository.

## Removal

To remove the module:

```powershell
Remove-Module CMHealthCheck -ErrorAction SilentlyContinue
Uninstall-Module CMHealthCheck
```

## License

This project is licensed under the MIT License. See [LICENSE](LICENSE) for details.
