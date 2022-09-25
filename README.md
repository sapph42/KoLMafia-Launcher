# KoLMafia-Launcher
A Windows launcher and auto-updater for the .jar version of KoLMafia

Usage:
Edit Line 2 to point to your local version of KoLMafia - specifically the directory that the KoLMafia .jar file is located in
Then run the ps1 file

To build your localized script as an executable, download the Mafia icon file from:
 https://raw.githubusercontent.com/kolmafia/kolmafia/r26795/src/images/app_icon64.png
Convert the png to an ico file using any available converter
Run the following code in PowerShell, in the folder where you downloaded Launch-KoL.ps1:
  Install-Module -Name ps2exe -Scope CurrentUser
  Invoke-ps2exe -inputFile .\Launch-KoL.ps1 -outputFile .\Launch-KoL.exe -iconFile .\app_icon64.ico
