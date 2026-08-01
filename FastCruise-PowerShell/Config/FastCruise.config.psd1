@{
    # =====================================================================
    # FastCruise.config.psd1
    #
    # Edit THIS file to change how FastCruise behaves at your site.
    # You should not need to open any .ps1/.psm1 file to reconfigure
    # FastCruise for a new location, share, or set of software checks.
    #
    # This is a native PowerShell data file (.psd1) - only literal values
    # are allowed (strings, numbers, arrays, booleans, and hashtables).
    # No code execution happens when this file is read.
    # =====================================================================

    # --- Network share mapping -----------------------------------------
    # FastCruise expects a mapped drive for the shared report location.
    # If NetworkDriveLetter is not already mapped, it is mapped from
    # NetworkSharePath the first time the script runs.
    NetworkDriveLetter = 'S:'
    NetworkSharePath   = '\\localhost\Folder-1'

    # --- Report locations -------------------------------------------------
    # ReportRoot / ReportFileName: the shared, network report every
    # workstation appends to. The month and year are inserted into the
    # file name automatically (e.g. FastCruise_2026-July.csv).
    # LocalReportPath: a per-workstation copy kept locally in case the
    # network path is unavailable.
    ReportRoot         = 'S:\FastCruise'
    ReportFileName     = 'FastCruise.csv'
    LocalReportPath    = 'C:\temp\FastCruise\FastCruiseFile.csv'

    # --- Location reference data -----------------------------------------
    # JSON file describing the Department > Building > Room hierarchy used
    # to populate the location picker. See Config\ComputerLocation.json.
    LocationConfigFile = 'S:\ComputerLocation.json'

    # Desk labels offered in the desk picker.
    DeskLabels = @(
        'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I',
        'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q'
    )

    # --- Software version checks -----------------------------------------
    # Each entry is matched (as a substring) against installed software
    # display names, and the version found is recorded on the report.
    SoftwareChecks = @(
        'Axway',
        'Mozilla Firefox',
        'McAfee Agent',
        'Java'
    )

    # --- Facility issue reporting -----------------------------------------
    # Rooms whose Desk label is in FacilityIssueDesks are prompted for
    # facility issues (broken lights, A/C, etc.) after each Fast Cruise.
    FacilityIssueFile  = 'S:\FC-Facility_Issue\Facility_Issue_Report.txt'
    FacilityIssueDesks = @('A', 'B')

    # --- Phone number validation -------------------------------------------
    # Regular expression a manually-entered phone number must match.
    PhoneNumberPattern = '^\d{3}-\d{3}-\d{4}'

    # --- Optional application launch tests ---------------------------------
    # Set RunApplicationTests to $true to prompt for these on every run.
    RunApplicationTests = $false
    ApplicationTests    = @{
        Adobe = @{
            TestFile    = 'S:\Information-Systems\Scripts\FastCruise\FastCruiseTestFile.pdf'
            TestProgram = 'C:\Program Files (x86)\Adobe\Acrobat 2015\Acrobat\Acrobat.exe'
            ProcessName = 'Acrobat'
        }
        PowerPoint = @{
            TestFile    = 'S:\Information-Systems\Scripts\FastCruise\FastCruiseTestFile.pptx'
            TestProgram = 'C:\Program Files (x86)\Microsoft Office\Office16\POWERPNT.EXE'
            ProcessName = 'POWERPNT'
        }
    }

    # --- Logging -------------------------------------------------------------
    Logging = @{
        LogDirectory = 'S:\FastCruise\Logs'
        MinimumLevel = 'Info'   # Debug | Verbose | Info | Warning | Error
    }
}
