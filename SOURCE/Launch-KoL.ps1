param(
    [switch]$noLaunch = $false,
    [switch]$killOnUpdate = $false,
    [switch]$Silent = $false
)

Add-Type -AssemblyName 'System.Windows.Forms, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089'
class Preferences {
    hidden [string]$_PrefPath
    hidden [string]$_installLocation
    hidden [int]$_maxAttempts
    hidden [boolean]$_silent

    Preferences(
        [string]$p
    ){
        $this._PrefPath = $p
        If (-not (Test-Path $p\Launch-Kol.pref)) {
            [preferences]::InitPref($p)
        }
        $this.LoadVals()
    }

    Preferences (
        [string]$p,
        [boolean]$s
    ){
        If (-not (Test-Path $p\Launch-Kol.pref) -and $s) {
            throw "No pref file found, silence precludes requesting install location"
        } elseif (-not (Test-Path $p\Launch-Kol.pref) -and -not $s) {
            [preferences]::InitPref($p)
        }
        $this._silent = $s
        $this._PrefPath = $p
        $this.LoadVals()
    }

    hidden static [void]InitPref([string]$p){
        $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $OpenFileDialog.initialDirectory = $p
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
        $default | Out-File -FilePath $p\Launch-Kol.pref -NoClobber
    }

    hidden [void]LoadVals(){
        $prefFile = "$($this._PrefPath)\Launch-Kol.pref"
        If (-not (Test-Path $prefFile)) {
            [preferences]::InitPref($this._PrefPath)
        }
        $prefs = [xml](Get-Content $prefFile)
        $this._installLocation = $prefs.preferences.Location.Trim()
        try {
            $this._maxAttempts = $prefs.preferences.MaxAttempts.Trim()
        } catch {
            $this.SetMaxAttempts(3)
        }
    }

    [string]GetLocation(){
        return $this._installLocation
    }

    [void]SetLocation([string]$p){
        $prefs = [xml](Get-Content "$($this._PrefPath)\Launch-Kol.pref")
        $prefs.preferences.Location = $p
        $prefs.Save("$($this._PrefPath)\Launch-Kol.pref")
    }

    [int]GetMaxAttempts(){
        return $this._maxAttempts
    }

    [void]SetMaxAttempts([int]$m){
        $prefs = [xml](Get-Content "$($this._PrefPath)\Launch-Kol.pref")
        try {
            $prefs.preferences.MaxAttempts = [string]$m
        } catch {
            $prefs.preferences.AppendChild($prefs.CreateElement('MaxAttempts'))
            $prefs.preferences.MaxAttempts = [string]$m
        }
        $prefs.Save("$($this._PrefPath)\Launch-Kol.pref")
    }
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
$preferences = [preferences]::new($scriptPath,$Silent) 
$installLocation = $preferences.GetLocation()
$maxAttempts = $preferences.GetMaxAttempts()
if ($installLocation -match 'InsertYourKolPath') {
    $installLocation = $scriptPath
    $preferences.SetLocation($scriptPath)
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

$msg = 'A javaw.exe process has been detected. For safety, update cannot continue without killing this process.'
$title = 'Java Interpreter Already Running'
$buttons = [System.Windows.Forms.MessageBoxButtons]::OKCancel
$cancelbutton = [System.Windows.Forms.DialogResult]::Cancel
$icon = [System.Windows.Forms.MessageBoxIcon]::Warning
$default = [System.Windows.Forms.MessageBoxDefaultButton]::Button2
$options = 0
if (Get-Process javaw -ErrorAction SilentlyContinue) {
    if ($killOnUpdate) {
        Get-Process javaw | Stop-Process -Force
    } elseif ($Silent) {
        EXIT 1
    } else {
        $answer = [System.Windows.Forms.MessageBox]::Show($msg,$title,$buttons,$icon,$default,$options)
        if ($answer -eq $cancelbutton) {
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
$msg = "No jar file found in the provided folder. Download latest mafia to $($installLocation)?"
$title = 'Mafia Not Found!'
$buttons = [System.Windows.Forms.MessageBoxButtons]::YesNo
$nobutton = [System.Windows.Forms.DialogResult]::No
$icon = [System.Windows.Forms.MessageBoxIcon]::Question
$default = [System.Windows.Forms.MessageBoxDefaultButton]::Button1
$options = 0
if ($current.count -gt 1) {
    $current | Sort-Object -Property Name | Select-Object -First 1 | Remove-Item
    $current = $current | Sort-Object -Property Name -Descending | Select-Object -First 1
    $exists = $true
    $localFingerprint = (Get-FileHashLocal -Path $current -Algorithm MD5).ToLower()
} elseif ($current.Count -eq 0) {
    if (-not $Silent) {
        $answer = [System.Windows.Forms.MessageBox]::Show($msg,$title,$buttons,$icon,$default,$options)
        if ($answer -eq $nobutton) {
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