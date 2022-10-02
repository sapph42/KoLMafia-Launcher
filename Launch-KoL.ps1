$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
#Don't Edit This Script
#Edit Launch-KoL.pref instead
if (-not (Test-Path $scriptPath\Launch-Kol.pref)) {
    $default = @"
<?xml version="1.0"?>
<Location>
    $scriptPath
</Location>
"@
    $default | Out-File -FilePath $scriptPath\Launch-Kol.pref -NoClobber
}
$installLocation = (Select-Xml -Path $scriptPath\Launch-KoL.pref -XPath '/Location').Node.InnerText.Trim()
if ($installLocation -match 'InsertYourKolPath') {
    $installLocation = $scriptPath
}

Function Get-ShellOpenFromExtention {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Extension
    )
    begin {
        if ($Extension -notmatch '\.?(?<Extension>[a-z0-9]{1,3}$)') {
            return $null
        }
        $Extension = -join('.', $Matches.Extension)
    }
    process {
        try {
            $RegisteredApplication = (Get-ItemProperty "Registry::HKEY_CLASSES_ROOT\$Extension" -ErrorAction SilentlyContinue).'(default)'
            if ($null -eq $RegisteredApplication) {return $null}
            $ShellOpen = (Get-ItemProperty "Registry::HKEY_CLASSES_ROOT\$RegisteredApplication\shell\open\command" -ErrorAction SilentlyContinue).'(default)'
        } catch {
            return $null
        }
        if ($null -eq $ShellOpen) {return $null}
        $ShellOpen = $ShellOpen.Replace(' %*','')
        if ($shellOpen -match '(?(^")(?<path>"[^"]*")|(?<path>[^ ]*)) (?<params>.*)') {
            return @{
                appPath = $Matches.path
                arguments = $Matches.params
            }
        }
    }
    end {
        return $null
    }
}

#Get Registered Application for jar files
$OpenCommand = Get-ShellOpenFromExtention -Extension '.jar'
if ($null -eq $OpenCommand) {
    #Could not pull jar association from registry.  Assume defaults and hope for the best
    $javaPath = "javaw.exe"
    $params = "-jar ""%1"""
} else {
    $javaPath = $OpenCommand.appPath
    $params = $OpenCommand.arguments
}
#Check kolmafia for name of latest build
$base = "https://builds.kolmafia.us/job/Kolmafia/lastSuccessfulBuild/artifact/dist/"
Set-Location $installLocation
$current = Get-ChildItem .\*.jar
if ($current.count -gt 1) {
    $current | Sort-Object -Property Name | Select-Object -First 1 | Remove-Item
    $current = $current | Sort-Object -Property Name -Descending | Select-Object -First 1
} elseif ($current.Count -eq 0) {
    $message = "No jar file found in the provided folder. Download latest mafia to $($installLocation)?"
    $answer = [System.Windows.Forms.MessageBox]::Show($message,"Mafia Not Found!",[System.Windows.Forms.MessageBoxButtons]::YesNo,[System.Windows.Forms.MessageBoxIcon]::Question,[System.Windows.Forms.MessageBoxDefaultButton]::Button1,0)
    if ($answer -eq [System.Windows.Forms.DialogResult]::No) {
        exit
    } else {
        $current = @{Name = '1'}
    }
}
try {
    $response = Invoke-WebRequest $base
    $latest = ($response.Links | Select href | ? {$_.href -like "*.jar"}).href
    #Get Current local build
    if ($current.Name -ne $latest) {
        #If the current local build does not match the latest build, remove local build and download latest
        Invoke-WebRequest $(-join($base, $latest)) -OutFile ".\$latest"
        if ($current.Name -ne '1') {
            Remove-Item $current.FullName
        }
    }
} catch {
    $StatusCode = $_.Exception.Response.StatusCode.value__
    $StatusDescription = $_.Exception.Response.StatusDescription
    [System.Windows.Forms.MessageBox]::Show("Received a $StatusCode ($StatusDescription) error when attempting to retreive the latest version.")
    $latest = $current
} finally {
    #Note: In theory Invoke-Item .\$latest should suffice.  However, doing this seems to load KoL without settings data.  So we do it the hard way
    Start-Process -FilePath $javaPath -ArgumentList $params.Replace("%1", ".\$latest")
}