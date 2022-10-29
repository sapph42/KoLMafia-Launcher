param(
    [switch]$noLaunch = $false,
    [switch]$killOnUpdate = $false,
    [switch]$Silent = $false,
    [switch]$Verbose
)

Add-Type -AssemblyName 'System.Windows.Forms, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089'
class Preferences {
    hidden [string]$_PrefPath
    hidden [string]$_installLocation
    hidden [int]$_maxAttempts
    hidden [boolean]$_silent
    hidden [string]$_skippedVersion

    Preferences(
        [string]$p
    ){
        $this._PrefPath = $p
        If (-not ([preferences]::PrefsExistAndNotNull($p))) {
            [preferences]::InitPref($p)
        }
        $this.LoadVals()
    }

    Preferences (
        [string]$p,
        [boolean]$s
    ){
        If (-not ([preferences]::PrefsExistAndNotNull($p)) -and $s) {
            throw "No registry config found, silence precludes requesting install location"
        } elseif (-not ([preferences]::PrefsExistAndNotNull($p)) -and -not $s) {
            [preferences]::InitPref($p)
        }
        $this._silent = $s
        $this._PrefPath = $p
        $this.LoadVals()
    }

    hidden static [boolean]PrefsExistAndNotNull([string]$p){
        if (-not (Test-Path $p)) {
            return $false
        }
        if (-not (Get-ItemPropertyValue -Path $p -Name PathToKoL -ErrorAction SilentlyContinue)) {
            return $false
        }
        if ( [string]::IsNullOrEmpty((Get-ItemPropertyValue -Path $p -Name PathToKoL)) ) {
            return $false
        }
        return $true
    }

    hidden static [void]InitPref([string]$p){
        if (-not
                (Test-Path $p) -and -not
                (Test-Path ($p | Split-Path))
        ) {
            New-Item ($p | Split-Path)
            Write-Verbose "Creating key at $($p | Split-Path)"
            
        }
        if (-not (Test-Path $p)) {
            New-Item $p
            Write-Verbose "Creating key at $p"
        }
        if (-not
            (Get-ItemPropertyValue -Path $p -Name PathToKoL -ErrorAction SilentlyContinue) -and
            [string]::IsNullOrEmpty((Get-ItemPropertyValue -Path $p -Name PathToKoL -ErrorAction SilentlyContinue))
        ) {
            $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
            $OpenFileDialog.initialDirectory = $env:USERPROFILE
            $OpenFileDialog.filter = “JAR files (*.jar)| *.jar”
            $OpenFileDialog.Title = "Select Location of KolMafia JAR"
            $OpenFileDialog.ShowDialog() | Out-Null
            $installLocation = $OpenFileDialog.filename | Split-Path -Parent
            New-ItemProperty -Path $p -Name 'PathToKoL' -Value $installLocation -PropertyType String
            Write-Verbose "Creating value PathToKoL at $p"
        }
        try {
            Get-ItemPropertyValue -Path $p -Name MaxDownloadAttempts | Out-Null
        } catch {
            New-ItemProperty -Path $p -Name 'MaxDownloadAttempts' -Value 3 -PropertyType Dword
            Write-Verbose "Creating value MaxDownloadAttempts at $p"
        }
        try {
            Get-ItemPropertyValue -Path $p -Name SkippedVersion | Out-Null
        } catch {
            New-ItemProperty -Path $p -Name 'SkippedVersion' -Value '' -PropertyType String
            Write-Verbose "Creating value SkippedVersion at $p"
        }
    }

    hidden [void]LoadVals(){
        If (-not ([preferences]::PrefsExistAndNotNull($this._PrefPath))) {
            [preferences]::InitPref($this._PrefPath)
        }
        $this._installLocation = Get-ItemPropertyValue -Path $this._PrefPath -Name PathToKoL
        $this._maxAttempts = Get-ItemPropertyValue -Path $this._PrefPath -Name MaxDownloadAttempts
        try {
            Get-ItemPropertyValue -Path $this._PrefPath -Name SkippedVersion | Out-Null
        } catch {
            New-ItemProperty -Path $this._PrefPath -Name 'SkippedVersion' -Value '' -PropertyType String
            Write-Verbose "Creating value SkippedVersion at $($this._PrefPath)"
        } finally {
            $this._skippedVersion = Get-ItemPropertyValue -Path $this._PrefPath -Name SkippedVersion
        }
    }

    [string]GetLocation(){
        return $this._installLocation
    }

    [void]SetLocation([string]$p){
        Set-ItemProperty -Path $this._PrefPath -Name PathToKoL -Value $p
        $this._installLocation = $p
    }

    [int]GetMaxAttempts(){
        return $this._maxAttempts
    }

    [void]SetMaxAttempts([int]$m){
        Set-ItemProperty -Path $this._PrefPath -Name MaxDownloadAttempts -Value $m
        $this._maxAttempts = $m
    }

    [string]GetSkippedVersion(){
        return $this._skippedVersion
    }

    [void]SetSkippedVersion([string]$v){
        Set-ItemProperty -Path $this._PrefPath -Name SkippedVersion -Value $v
        $this._skippedVersion = $v
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
        [string]$Algorithm,

        [Parameter(Mandatory, ParameterSetName='NoFingerprint')]
        [switch]$NoFingerPrint
    )
    Start-BitsTransfer -Source $URI -Destination $Destination -Priority $Priority
    Write-Verbose "Fetching file $URI"
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

    Write-Verbose "Generating local hash"
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

Function CheckForUpdate() {
    param(
        [string]$SkippedVersion
    )
    try {
        $Parameters = @{
            Path = 'HKLM:\SOFTWARE\WOW6432Node\Sapph Tools\KoLMafia Launcher'
            Name = 'Version'
        }
        $localVer = Get-ItemPropertyValue @Parameters
    } catch {
        $caller = (Get-Process -Id $pid).Path
        $localVer = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($caller).FileVersion    
    }
    try {
        Write-Verbose "Fetching published release version"
        $releaseVer = (Invoke-WebRequest `
            https://raw.githubusercontent.com/sapph42/KoLMafia-Launcher/main/version.txt
        ).ToString().Trim()
    } catch {
        return $null
    }
    if ($releaseVer -eq $SkippedVersion) {
        Write-Verbose "Local version up-to-date"
        return $null
    }
    if ($localVer -ne $releaseVer) {
        Write-Verbose "Version mismatch"
        $msg = "An updated version ($($releaseVer)) of KoLMafia Launcher is available! Update now? (Or skip this release?)"
        $title = 'Update available!'
        $buttons = [System.Windows.Forms.MessageBoxButtons]::YesNo
        $nobutton = [System.Windows.Forms.DialogResult]::No
        $icon = [System.Windows.Forms.MessageBoxIcon]::Question
        $default = [System.Windows.Forms.MessageBoxDefaultButton]::Button1
        $options = 0
        $answer = [System.Windows.Forms.MessageBox]::Show($msg,$title,$buttons,$icon,$default,$options)
        if ($answer -eq $nobutton) {
            return $releaseVer
        }
        $sourceURI = "https://github.com/sapph42/KoLMafia-Launcher/raw/main/KoLMafia-Launcher_$($releaseVer).exe"
        $destination = "$($env:TEMP)\$($targetInstallerName)"
        try {
            Get-WebFile -URI $sourceURI -Destination $destination -Priority Foreground -NoFingerPrint | Out-Null
            Start-Process $destination -Verb RunAs
            EXIT 0
        } catch {
        }
        return $null
    } else {
        return $null
    }
}

if ($Verbose) {
    $VerbosePreference = "Continue"
}
$preferences = [preferences]::new('HKCU:\Software\Sapph Tools\KoLMafia Launcher\',$Silent)
$installLocation = $preferences.GetLocation()
$maxAttempts = $preferences.GetMaxAttempts()
$skippedVersion = $preferences.GetSkippedVersion()

if (-not $Silent) {
    $retVal = CheckForUpdate -SkippedVersion $skippedVersion
    if ($null -ne $retVal) {
        Write-Version "New release skipped"
        $preferences.SetSkippedVersion($retVal)
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
        Write-Verbose "Checking for active Java process"
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
    if ($null -eq $cannonicalFingerprint) {
        $cannonicalFingerprint = ($response.AllElements | Where-Object {$_.tagName -eq 'LI' -and $_.class -eq 'jenkins-breadcrumbs__list-item' -and $_.innerText -match '[a-z0-9]{32}'}).innerText
    }
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
        Write-Verbose "Launching JAR"
        Start-Process -FilePath $javaPath -ArgumentList $params.Replace("%1", ".\$latest")
    }
}