# Powershell_Get_AD_Group_Members_GUI
This script retrieves the members of an Active Directory Group and displays it in a GUI, additionally it allows for saving to a CSV file.

## Usage
To use it, there are two options. Either create a shortcut that will be used to launch the script, or alternatively, open a command line (Powershell or CMD). 

## (Option #1) Shortcut
For the first option, a shortcut is created and the target is set to:

```C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -ExecutionPolicy Bypass -File GroupMembersSummary.ps1```

Additionally I recommend opening the properties of the shortcut, click the tab "Shortcut", click "Advanced", and then check "Run as administrator"

## (Option #2) Run script from Command Line
Open either Powershell or CMD, and run the following command from the location of the script:

```powershell.exe -ExecutionPolicy Bypass -File GroupMembersSummary.ps1```

## Dependencies
For this script to work, you must have the `Powershell Active Directory Module` installed.

 A good guide on how to install the Powershell Active Directory Module can be found [HERE](https://4sysops.com/wiki/how-to-install-the-powershell-active-directory-module/)
