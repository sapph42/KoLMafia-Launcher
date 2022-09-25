Install-Module -Name ps2exe -Scope CurrentUser
If (Test-Path .\app_icon64.ico) {
    Invoke-ps2exe -inputFile .\Launch-KoL.ps1 -outputFile .\Launch-KoL.exe -iconFile .\app_icon64.ico
} else {
    Invoke-ps2exe -inputFile .\Launch-KoL.ps1 -outputFile .\Launch-KoL.exe
}