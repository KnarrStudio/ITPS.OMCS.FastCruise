# filepath: /FastCruise-PowerShell/FastCruise-PowerShell/scripts/Run-FastCruise.ps1

# This script serves as a wrapper to execute the main FastCruise functionality by calling Start-FastCruise.ps1 with the appropriate parameters.

# Import the main FastCruise script
. ../src/Start-FastCruise.ps1

# Define parameters for the FastCruise execution
$FastCruiseParams = @{
    FastCruiseReportPath = 'S:\FastCruise'
    FastCruiseFile       = 'FastCruise.csv'
    Verbose              = $true
}

# Execute the FastCruise functionality
Start-FastCruise @FastCruiseParams