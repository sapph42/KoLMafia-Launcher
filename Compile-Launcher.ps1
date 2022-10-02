$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
if (Get-Module | Where-Object {$_.Name -eq 'ps2exe'}) {
} else {
    if (Get-Module -ListAvailable | Where-Object {$_.Name -eq 'ps2exe'}) {
        Import-Module 'ps2exe' -Verbose
    } else {
        if (Find-Module -Name 'ps2exe' | Where-Object {$_.Name -eq 'ps2exe'}) {
            Install-Module -Name 'ps2exe' -Force -Verbose -Scope CurrentUser
            Import-Module 'ps2exe' -Verbose
        } else {
            EXIT 1
        }
    }
}
Set-Location $scriptPath
If (Test-Path .\app_icon64.ico) {
    Invoke-ps2exe -inputFile .\Launch-KoL.ps1 -outputFile .\Launch-KoL.exe -iconFile .\app_icon64.ico
} else {
    Invoke-ps2exe -inputFile .\Launch-KoL.ps1 -outputFile .\Launch-KoL.exe
}