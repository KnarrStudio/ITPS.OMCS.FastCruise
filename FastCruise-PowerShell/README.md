# FastCruise PowerShell Project

## Overview
The FastCruise PowerShell project is designed to streamline the process of gathering and reporting workstation information, installed software, and facility issues. This modular approach allows for better organization and maintainability of the code.

## Project Structure
```
FastCruise-PowerShell
├── src
│   ├── Start-FastCruise.ps1
│   ├── Get-InstalledSoftware.ps1
│   ├── Get-WorkstationInfo.ps1
│   ├── Get-MacAddress.ps1
│   ├── Get-ComputerLocation.ps1
│   ├── Get-LastComputerStatus.ps1
│   ├── Get-FacilityIssues.ps1
│   ├── Start-ApplicationTest.ps1
│   ├── Show-VbForm.ps1
│   ├── Show-AsciiMenu.ps1
│   └── utils
│       └── Convert-JSONToHash.ps1
├── scripts
│   └── Run-FastCruise.ps1
├── data
│   └── ComputerLocation.json
├── README.md
```

## Installation
1. Clone the repository to your local machine.
2. Open PowerShell and navigate to the project directory.
3. Ensure that you have the necessary permissions to run scripts. You may need to set the execution policy:
   ```powershell
   Set-ExecutionPolicy RemoteSigned
   ```

## Usage
To run the FastCruise application, execute the `Run-FastCruise.ps1` script located in the `scripts` directory. This script serves as a wrapper to initiate the main functionality of the application.

```powershell
.\scripts\Run-FastCruise.ps1
```

## Script Descriptions
- **Start-FastCruise.ps1**: Main entry point for the FastCruise application. Initializes the script and calls necessary functions.
- **Get-InstalledSoftware.ps1**: Retrieves a list of installed software on the workstation.
- **Get-WorkstationInfo.ps1**: Retrieves various information about the workstation, such as manufacturer and model.
- **Get-MacAddress.ps1**: Retrieves the MAC address of the workstation's network adapter.
- **Get-ComputerLocation.ps1**: Retrieves computer location information from a JSON file or prompts for user input.
- **Get-LastComputerStatus.ps1**: Imports the last recorded status of the workstation from a CSV file.
- **Get-FacilityIssues.ps1**: Allows users to report issues related to the facility and writes them to a text file.
- **Start-ApplicationTest.ps1**: Tests whether specified applications can be launched successfully.
- **Show-VbForm.ps1**: Creates and displays a Visual Basic form for user input.
- **Show-AsciiMenu.ps1**: Creates a simple ASCII menu for user interaction.
- **Convert-JSONToHash.ps1**: Converts a JSON object into a PowerShell hash table for easier manipulation.

## Contributing
Contributions are welcome! Please submit a pull request or open an issue for any enhancements or bug fixes.

## License
This project is licensed under the MIT License. See the LICENSE file for more details.