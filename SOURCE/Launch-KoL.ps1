param(
    [switch]$noLaunch = $false,
    [switch]$killOnUpdate = $false,
    [switch]$Silent = $false
)

Add-Type -AssemblyName 'System.Windows.Forms, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089'
Function Initialize-Installtion {
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $scriptPath
    $OpenFileDialog.filter = “JAR files (*.jar)| *.jar”
    $OpenFileDialog.Title = "Select Location of KolMafia JAR"
    $OpenFileDialog.ShowDialog() | Out-Null
    $installLocation = $OpenFileDialog.filename | Split-Path -Parent
    $default = @"
<?xml version="1.0"?>
<preferences>
    <Location>$installLocation</Location>
    <MaxAttempts>3</MaxAttempts>
</preferences>
"@
    $default | Out-File -FilePath $scriptPath\Launch-Kol.pref -NoClobber
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
Function Get-WebFile {
    [CmdletBinding()]
    [OutputType([boolean])]
    param (
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$URI,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Destination,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [ValidateSet('Foreground','High','Medium','Low')]
        [string]$Priority,

        [Parameter(Mandatory, ParameterSetName='Fingerprint')]
        [ValidateNotNullOrEmpty()]
        [string]$Fingerprint,

        [Parameter(Mandatory, ParameterSetName='Fingerprint')]
        [ValidateNotNullOrEmpty()]
        [ValidateSet('SHA1','SHA256','SHA384','SHA512','MD5')]
        [string]$Algorithm
    )
    Start-BitsTransfer -Source $URI -Destination $Destination -Priority $Priority
    If ($PSBoundParameters.ContainsKey('Fingerprint')) {
        $TargetHash = (Get-FileHashLocal -Path $Destination -Algorithm $Algorithm).ToLower()
        if ($TargetHash -eq $Fingerprint) {
            return $true
        } else {
            return $false
        }
    } else {
        return $true
    }
}

Function Get-FileHashLocal {
    [OutputType([string])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Path,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [ValidateSet('SHA1','SHA256','SHA384','SHA512','MD5')]
        [string]$Algorithm        
    )

    if (Get-Command Get-FileHash -ErrorAction SilentlyContinue) {
        return (Get-FileHash -Path $Path -Algorithm $Algorithm).Hash
    } else {
        $FileStream = New-Object System.IO.FileStream($Path, [System.IO.FileMode]::Open)
        $StringBuilder = New-Object System.Text.StringBuilder
        [System.Security.Cryptography.HashAlgorithm]::Create($Algorithm).ComputeHash($FileStream) | ForEach-Object {[Void]$StringBuilder.Append($_.ToString("x2"))}
        $FileStream.Close()
        return $StringBuilder.ToString()
    }
}

$scriptPath = $PSScriptRoot
if (-not (Test-Path $scriptPath\Launch-Kol.pref)) {
    if (-not $Silent) {
        Initialize-Installtion
    } else {
        EXIT 1
    }
}
$preferences = [xml](Get-Content $scriptPath\Launch-Kol.pref)
$installLocation = $preferences.preferences.Location.Trim()
try {
    $maxAttempts = $preferences.preferences.MaxAttempts.Trim()
} catch {
    $maxAttempts = 3
    $preferences.preferences.AppendChild($preferences.CreateElement("MaxAttempts"))
    $preferences.Save("$scriptPath\Launch-Kol.pref")
}
if ($installLocation -match 'InsertYourKolPath') {
    $installLocation = $scriptPath
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
if ($killOnUpdate) {
    Get-Process javaw | Stop-Process -Force
}
if (Get-Process javaw) {
    if ($killOnUpdate) {
        Get-Process javaw | Stop-Process -Force
    } elseif ($Silent) {
        EXIT 1
    } else {
        $message = "A javaw.exe process has been detected. For safety, update cannot continue without killing this process."
        $answer = [System.Windows.Forms.MessageBox]::Show($message,"Java Interpreter Already Running",[System.Windows.Forms.MessageBoxButtons]::OKCancel,[System.Windows.Forms.MessageBoxIcon]::Warning,[System.Windows.Forms.MessageBoxDefaultButton]::Button2,0)
        if ($answer -eq [System.Windows.Forms.DialogResult]::Cancel) {
            EXIT 1
        } else {
            Get-Process javaw | Stop-Process -Force
        }
    }
}

#Check kolmafia for name of latest build
$base = "https://builds.kolmafia.us/job/Kolmafia/lastSuccessfulBuild/artifact/dist/"
Set-Location $installLocation
$current = Get-ChildItem .\*.jar
if ($current.count -gt 1) {
    $current | Sort-Object -Property Name | Select-Object -First 1 | Remove-Item
    $current = $current | Sort-Object -Property Name -Descending | Select-Object -First 1
    $exists = $true
    $localFingerprint = (Get-FileHashLocal -Path $current -Algorithm MD5).ToLower()
} elseif ($current.Count -eq 0) {
    if (-not $Silent) {
        $message = "No jar file found in the provided folder. Download latest mafia to $($installLocation)?"
        $answer = [System.Windows.Forms.MessageBox]::Show($message,"Mafia Not Found!",[System.Windows.Forms.MessageBoxButtons]::YesNo,[System.Windows.Forms.MessageBoxIcon]::Question,[System.Windows.Forms.MessageBoxDefaultButton]::Button1,0)
        if ($answer -eq [System.Windows.Forms.DialogResult]::No) {
            exit
        } else {
            $exists = $false
        }
    } else {
        $exists = $false
    } 
} else {
    $exists = $true
    $localFingerprint = (Get-FileHashLocal -Path $current -Algorithm MD5).ToLower()
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
        $attempts = 1
        $downloadSuccess = Get-WebFile -URI $jarURI -Destination ".\$latest" -Priority Foreground -Fingerprint $cannonicalFingerprint -Algorithm MD5
        while ((-not $downloadSuccess) -and ($attempts -le $maxAttempts)) {
            Remove-Item ".\$latest"
            $downloadSuccess = Get-WebFile -URI $jarURI -Destination ".\$latest" -Priority Foreground -Fingerprint $cannonicalFingerprint -Algorithm MD5
            $attempts++
        }
        if ($exists -and $downloadSuccess) {
            Remove-Item $current.FullName
        }
    }
} catch {
    if (-not $Silent) {
        [System.Windows.Forms.MessageBox]::Show("Received a error when attempting to retreive and/or save the latest version. Your previous version has been kept, and will now be run.")
    }
    $latest = $current
} finally {
    #Note: In theory Invoke-Item .\$latest should suffice.  However, doing this seems to load KoL without settings data.  So we do it the hard way
    if (-not $noLaunch) {
        Start-Process -FilePath $javaPath -ArgumentList $params.Replace("%1", ".\$latest")
    }
}