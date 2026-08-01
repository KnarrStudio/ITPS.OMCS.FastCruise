# FastCruise PowerShell

A quick operational test ("Fast Cruise") of workstations. Captures the date,
time, and user who performed the check, records hardware/software facts
automatically, and walks the technician through a short set of prompts for
physical location, phone number, notes, and facility issues.

This is a refactor of the original single-file `Start-FastCruise.ps1` script
into a proper PowerShell module, a shared logging module, and one editable
configuration file. It replaces both `Scripts/` and the earlier, incomplete
`FastCruise-PowerShell/` folder in this repository.

## Project layout

```
FastCruise-PowerShell/
├── Run-FastCruise.ps1              <- what you actually run
├── Config/
│   ├── FastCruise.config.psd1      <- EDIT THIS for your site (paths, software list, desk labels, etc.)
│   ├── ComputerLocation.json       <- Department > Building > Room reference data used by the location picker
│   └── ComputerLocation.hashtable.psd1   <- easier-to-edit source for ComputerLocation.json (see ConvertTo-LocationJson)
├── Modules/
│   ├── FCLogging/                  <- common logging module, reusable by any script in this repo
│   │   ├── FCLogging.psd1
│   │   └── FCLogging.psm1
│   └── FastCruise/                 <- the FastCruise module itself
│       ├── FastCruise.psd1
│       ├── FastCruise.psm1
│       ├── Public/                 <- exported functions (Start-FastCruise, Export-ComputerDescription, ConvertTo-LocationJson, Show-FastCruiseMenu)
│       └── Private/                <- internal helper functions
└── Reports/                        <- default local landing spot for generated reports
```

## Requirements

- Windows, with Windows PowerShell 5.1 **or** PowerShell 7+ (`pwsh`).
- The `NetAdapter` module (ships with Windows) for MAC address lookups.
- Domain-joined workstation for the network report path; falls back to a
  local temp path automatically if the domain isn't reachable.

## Usage

1. Edit `Config\FastCruise.config.psd1` for your site: report paths, the
   software versions to check, desk labels, application tests, and logging
   options. **You should not need to edit any `.ps1` or `.psm1` file to
   reconfigure FastCruise for a new location.**
2. Edit `Config\ComputerLocation.json` (or edit
   `ComputerLocation.hashtable.psd1` and run `ConvertTo-LocationJson`) to
   describe your Department/Building/Room hierarchy.
3. Run it:

   ```powershell
   .\Run-FastCruise.ps1
   ```

   Useful variations:

   ```powershell
   # Record an entry by hand instead of via auto-detection
   .\Run-FastCruise.ps1 -ManualInput

   # Run one check and exit, without the follow-up menu (e.g. from a scheduled task)
   .\Run-FastCruise.ps1 -NoMenu

   # Use a config file somewhere else
   .\Run-FastCruise.ps1 -ConfigPath 'D:\FastCruiseConfigs\Warehouse.config.psd1'
   ```

4. After a batch of Fast Cruises, build an Active Directory description list:

   ```powershell
   Import-Module .\Modules\FastCruise\FastCruise.psd1
   Export-ComputerDescription -InputReportFile 'S:\FastCruise\FastCruise_2026-July.csv' -OutputListFile 'S:\FastCruise\ComputerDescriptions.csv'
   ```

## What changed from the original scripts

- **Modules instead of one big script.** Each piece of functionality
  (installed-software lookup, workstation info, MAC address, location
  picker, facility issues, application test, file I/O) is now its own
  documented function in `Modules\FastCruise\Private`, imported by the
  `FastCruise` module. `Start-FastCruise` orchestrates them instead of
  defining them inline on every run.
- **A common logging module (`FCLogging`).** Every function calls
  `Write-FCLog` instead of a mix of `Write-Verbose`/`Write-Warning`/
  `Write-Output`. Log lines are timestamped, leveled, color-coded on
  screen, and appended to a monthly rolling log file.
- **Comment-based help everywhere.** Every function in every module has a
  full `.SYNOPSIS` / `.DESCRIPTION` / `.PARAMETER` / `.EXAMPLE` block —
  run `Get-Help <FunctionName> -Full` for any of them.
- **One config file, not values buried in code.** Software checks, desk
  labels, all file paths, the phone-number pattern, and application-test
  definitions moved from hardcoded script variables into
  `Config\FastCruise.config.psd1`.
- **Hashtable conversion removed.** `Convert-JSONToHash` is gone.
  `ConvertFrom-Json` already returns a navigable object; the location
  picker reads it directly.
- **Uniform file formats.** All CSV reads/writes go through
  `Export-FastCruiseRecord` / `Import-FastCruiseRecord`, and all JSON/text
  writes go through `Set-FCContent` / `Add-FCContent`, so every file this
  project produces is UTF-8 without a byte-order-mark, regardless of
  whether it was written by Windows PowerShell 5.1 or PowerShell 7. (The
  original `Configfiles\computerlocation.json` was UTF-16 while the CSV
  reports were not - that inconsistency is what prompted this.)
- **One consistent prompt UI.** `Show-InputDialog`, `Show-ConfirmDialog`,
  and `Show-SelectionDialog` (all Windows Forms) replace the mix of
  `Microsoft.VisualBasic` InputBox/MsgBox and `Out-GridView` pickers used
  before, and work the same way on PowerShell 7+ as on 5.1.
- **`Get-WmiObject` replaced with `Get-CimInstance`** for PowerShell 7+
  compatibility.
- **Bug fix:** `Start-ApplicationTest` now actually launches `TestProgram`
  with `TestFile` as its argument (the original opened `TestFile` directly
  and never used `TestProgram`).
- **Duplicate logic removed:** the "last four characters of a MAC address"
  formatting existed almost identically in both the main script and
  `Export-ComputerDescription.ps1`; it is now one shared helper
  (`ConvertTo-MacAddressSuffix`).

## Retired

- `Scripts\Log-Inventory.ps1` and `Scripts\Log-Inventory.ps1.bak` - an
  earlier, parallel version of the same script. Superseded by this project;
  kept in git history but not carried forward.
- `Configfiles\computerlocation.json` and `Scripts\computerlocation.json` -
  byte-identical duplicates of the same file. Superseded by the single
  `Config\ComputerLocation.json`.
- The original `FastCruise-PowerShell\` module attempt - incomplete (missing
  WSUS/domain detection, the menu, several fields) and contained a
  hardcoded personal file path. This project replaces it.
