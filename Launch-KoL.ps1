Add-Type -AssemblyName 'System.Windows.Forms, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089'
$scriptPath = $PSScriptRoot
if (-not (Test-Path $scriptPath\Launch-Kol.pref)) {
    $default = @"
<?xml version="1.0"?>
    <Location>$scriptPath</Location>
    <MaxAttempts>3</MaxAttempts>
"@
    $default | Out-File -FilePath $scriptPath\Launch-Kol.pref -NoClobber
}
$installLocation = (Select-Xml -Path $scriptPath\Launch-KoL.pref -XPath '/Location').Node.InnerText.Trim()
try {
    $maxAttempts = (Select-Xml -Path $scriptPath\Launch-KoL.pref -XPath '/MaxAttempts').Node.InnerText.Trim()
} catch {
    $maxAttempts = 3
}
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
    $exists = $true
    $localFingerprint = (Get-FileHash $current -Algorithm MD5).Hash.ToLower()
} elseif ($current.Count -eq 0) {
    $message = "No jar file found in the provided folder. Download latest mafia to $($installLocation)?"
    $answer = [System.Windows.Forms.MessageBox]::Show($message,"Mafia Not Found!",[System.Windows.Forms.MessageBoxButtons]::YesNo,[System.Windows.Forms.MessageBoxIcon]::Question,[System.Windows.Forms.MessageBoxDefaultButton]::Button1,0)
    if ($answer -eq [System.Windows.Forms.DialogResult]::No) {
        exit
    } else {
        $exists = $false
    }
} else {
    $exists = $true
    $localFingerprint = (Get-FileHash $current -Algorithm MD5).Hash.ToLower()
}
try {
    $response = Invoke-WebRequest $base
    $latest = ($response.Links | Select-Object href | Where-Object {$_.href -like "*.jar"}).href
    $jarURI = -join($base, $latest)
    $fingerprintURI = -join($jarURI, '/*fingerprint*/')
    $response = Invoke-WebRequest $fingerprintURI
    $cannonicalFingerprint = ($response.links | Select-Object href, innerText | Where-Object {$_.href -eq '/' -and $_.innerText -match '[a-z0-9]{32}'}).innerText
    #Get Current local build
    if ((-not $exists) -or ($current.Name -ne $latest) -or ($localFingerprint -ne $cannonicalFingerprint)) {
        #If the current local build does not match the latest build, remove local build and download latest
        Invoke-WebRequest $jarURI -OutFile ".\$latest"
        $attempts = 1
        $localFingerprint = (Get-FileHash ".\$latest" -Algorithm MD5).Hash.ToLower()
        while (($cannonicalFingerprint -ne $localFingerprint) -and ($attempts -le $maxAttempts)) {
            Remove-Item ".\$latest"
            Invoke-WebRequest $jarURI -OutFile ".\$latest"
            $attempts++
        }
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
