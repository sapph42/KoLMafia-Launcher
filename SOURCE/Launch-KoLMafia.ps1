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
        try {
            $needtoask = [string]::IsNullOrEmpty((Get-ItemPropertyValue -Path $p -Name PathToKoL -ErrorAction SilentlyContinue))
        } catch {
            $needtoask = $true
        }
        if ($needtoask) {
            $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
            $OpenFileDialog.initialDirectory = $env:USERPROFILE
            $OpenFileDialog.filter = “JAR files (*.jar)| *.jar”
            $OpenFileDialog.Title = "Select Location of KolMafia JAR"
            $OpenFileDialog.ShowDialog() | Out-Null
            $askfordir = [string]::IsNullOrEmpty($OpenFileDialog.FileName)
            if (-not $askfordir) {
                $installLocation = $OpenFileDialog.filename | Split-Path -Parent
                New-ItemProperty -Path $p -Name 'PathToKoL' -Value $installLocation -PropertyType String
                Write-Verbose "Creating value PathToKoL at $p"
            }
        } else {
            $askfordir = $false
        }
        if ($askfordir) {
            $msg = "No file was selected. Should Launcher assume you don't have KoLMafia installed and ask where you want it? (Selecting No will exit immediately)"
            $title = 'No file selected'
            $buttons = [System.Windows.Forms.MessageBoxButtons]::YesNo
            $nobutton = [System.Windows.Forms.DialogResult]::No
            $icon = [System.Windows.Forms.MessageBoxIcon]::Question
            $default = [System.Windows.Forms.MessageBoxDefaultButton]::Button1
            $options = 0
            $answer = [System.Windows.Forms.MessageBox]::Show($msg,$title,$buttons,$icon,$default,$options)
            if ($answer -eq $nobutton) {
                EXIT 0
            }  else {
                $SelectFolderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
                $SelectFolderDialog.SelectedPath = $env:USERPROFILE
                $SelectFolderDialog.Description = "Select where you want to install KoLMafia"
                $SelectFolderDialog.ShowNewFolderButton = $true
                $SelectFolderDialog.ShowDialog() | Out-Null
                if ([string]::IsNullOrEmpty($SelectFolderDialog.SelectedPath)) {
                    EXIT 0
                } else {
                    $installLocation = $SelectFolderDialog.SelectedPath
                    New-ItemProperty -Path $p -Name 'PathToKoL' -Value $installLocation -PropertyType String
                    Write-Verbose "Creating value PathToKoL at $p"
                }
            }
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
        $targetInstallerName = "KoLMafia-Launcher_$($releaseVer).exe"
        $sourceURI = "https://github.com/sapph42/KoLMafia-Launcher/raw/main/$($targetInstallerName)"
        $destination = "$($env:TEMP)\$($targetInstallerName)"
        try {
            Get-WebFile -URI $sourceURI -Destination $destination -Priority Foreground -NoFingerPrint | Out-Null
            Start-Process $destination -Verb RunAs
            EXIT 0
        } catch {
            Write-Error "Failed to download."
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
        Write-Error "New release skipped"
        $preferences.SetSkippedVersion($retVal)
    }
}


#Get Registered Application for jar files
$OpenCommand = Get-ShellOpenFromExtention -Extension '.jar'
if ($null -eq $OpenCommand[0]) {
    #Could not pull jar association from registry.  Assume defaults and hope for the best
    $javaPath = "javaw.exe"
    $params = "-jar ""%1"""
} else {
    $javaPath = $OpenCommand.appPath
    $params = $OpenCommand.arguments
}
if ($killOnUpdate) {
    Get-Process javaw -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
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
# SIG # Begin signature block
# MIIVkAYJKoZIhvcNAQcCoIIVgTCCFX0CAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUXp/x2bakklfimZelkIZxQbDj
# DwGgghHwMIIFbzCCBFegAwIBAgIQSPyTtGBVlI02p8mKidaUFjANBgkqhkiG9w0B
# AQwFADB7MQswCQYDVQQGEwJHQjEbMBkGA1UECAwSR3JlYXRlciBNYW5jaGVzdGVy
# MRAwDgYDVQQHDAdTYWxmb3JkMRowGAYDVQQKDBFDb21vZG8gQ0EgTGltaXRlZDEh
# MB8GA1UEAwwYQUFBIENlcnRpZmljYXRlIFNlcnZpY2VzMB4XDTIxMDUyNTAwMDAw
# MFoXDTI4MTIzMTIzNTk1OVowVjELMAkGA1UEBhMCR0IxGDAWBgNVBAoTD1NlY3Rp
# Z28gTGltaXRlZDEtMCsGA1UEAxMkU2VjdGlnbyBQdWJsaWMgQ29kZSBTaWduaW5n
# IFJvb3QgUjQ2MIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAjeeUEiIE
# JHQu/xYjApKKtq42haxH1CORKz7cfeIxoFFvrISR41KKteKW3tCHYySJiv/vEpM7
# fbu2ir29BX8nm2tl06UMabG8STma8W1uquSggyfamg0rUOlLW7O4ZDakfko9qXGr
# YbNzszwLDO/bM1flvjQ345cbXf0fEj2CA3bm+z9m0pQxafptszSswXp43JJQ8mTH
# qi0Eq8Nq6uAvp6fcbtfo/9ohq0C/ue4NnsbZnpnvxt4fqQx2sycgoda6/YDnAdLv
# 64IplXCN/7sVz/7RDzaiLk8ykHRGa0c1E3cFM09jLrgt4b9lpwRrGNhx+swI8m2J
# mRCxrds+LOSqGLDGBwF1Z95t6WNjHjZ/aYm+qkU+blpfj6Fby50whjDoA7NAxg0P
# OM1nqFOI+rgwZfpvx+cdsYN0aT6sxGg7seZnM5q2COCABUhA7vaCZEao9XOwBpXy
# bGWfv1VbHJxXGsd4RnxwqpQbghesh+m2yQ6BHEDWFhcp/FycGCvqRfXvvdVnTyhe
# Be6QTHrnxvTQ/PrNPjJGEyA2igTqt6oHRpwNkzoJZplYXCmjuQymMDg80EY2NXyc
# uu7D1fkKdvp+BRtAypI16dV60bV/AK6pkKrFfwGcELEW/MxuGNxvYv6mUKe4e7id
# FT/+IAx1yCJaE5UZkADpGtXChvHjjuxf9OUCAwEAAaOCARIwggEOMB8GA1UdIwQY
# MBaAFKARCiM+lvEH7OKvKe+CpX/QMKS0MB0GA1UdDgQWBBQy65Ka/zWWSC8oQEJw
# IDaRXBeF5jAOBgNVHQ8BAf8EBAMCAYYwDwYDVR0TAQH/BAUwAwEB/zATBgNVHSUE
# DDAKBggrBgEFBQcDAzAbBgNVHSAEFDASMAYGBFUdIAAwCAYGZ4EMAQQBMEMGA1Ud
# HwQ8MDowOKA2oDSGMmh0dHA6Ly9jcmwuY29tb2RvY2EuY29tL0FBQUNlcnRpZmlj
# YXRlU2VydmljZXMuY3JsMDQGCCsGAQUFBwEBBCgwJjAkBggrBgEFBQcwAYYYaHR0
# cDovL29jc3AuY29tb2RvY2EuY29tMA0GCSqGSIb3DQEBDAUAA4IBAQASv6Hvi3Sa
# mES4aUa1qyQKDKSKZ7g6gb9Fin1SB6iNH04hhTmja14tIIa/ELiueTtTzbT72ES+
# BtlcY2fUQBaHRIZyKtYyFfUSg8L54V0RQGf2QidyxSPiAjgaTCDi2wH3zUZPJqJ8
# ZsBRNraJAlTH/Fj7bADu/pimLpWhDFMpH2/YGaZPnvesCepdgsaLr4CnvYFIUoQx
# 2jLsFeSmTD1sOXPUC4U5IOCFGmjhp0g4qdE2JXfBjRkWxYhMZn0vY86Y6GnfrDyo
# XZ3JHFuu2PMvdM+4fvbXg50RlmKarkUT2n/cR/vfw1Kf5gZV6Z2M8jpiUbzsJA8p
# 1FiAhORFe1rYMIIGGjCCBAKgAwIBAgIQYh1tDFIBnjuQeRUgiSEcCjANBgkqhkiG
# 9w0BAQwFADBWMQswCQYDVQQGEwJHQjEYMBYGA1UEChMPU2VjdGlnbyBMaW1pdGVk
# MS0wKwYDVQQDEyRTZWN0aWdvIFB1YmxpYyBDb2RlIFNpZ25pbmcgUm9vdCBSNDYw
# HhcNMjEwMzIyMDAwMDAwWhcNMzYwMzIxMjM1OTU5WjBUMQswCQYDVQQGEwJHQjEY
# MBYGA1UEChMPU2VjdGlnbyBMaW1pdGVkMSswKQYDVQQDEyJTZWN0aWdvIFB1Ymxp
# YyBDb2RlIFNpZ25pbmcgQ0EgUjM2MIIBojANBgkqhkiG9w0BAQEFAAOCAY8AMIIB
# igKCAYEAmyudU/o1P45gBkNqwM/1f/bIU1MYyM7TbH78WAeVF3llMwsRHgBGRmxD
# eEDIArCS2VCoVk4Y/8j6stIkmYV5Gej4NgNjVQ4BYoDjGMwdjioXan1hlaGFt4Wk
# 9vT0k2oWJMJjL9G//N523hAm4jF4UjrW2pvv9+hdPX8tbbAfI3v0VdJiJPFy/7Xw
# iunD7mBxNtecM6ytIdUlh08T2z7mJEXZD9OWcJkZk5wDuf2q52PN43jc4T9OkoXZ
# 0arWZVeffvMr/iiIROSCzKoDmWABDRzV/UiQ5vqsaeFaqQdzFf4ed8peNWh1OaZX
# nYvZQgWx/SXiJDRSAolRzZEZquE6cbcH747FHncs/Kzcn0Ccv2jrOW+LPmnOyB+t
# AfiWu01TPhCr9VrkxsHC5qFNxaThTG5j4/Kc+ODD2dX/fmBECELcvzUHf9shoFvr
# n35XGf2RPaNTO2uSZ6n9otv7jElspkfK9qEATHZcodp+R4q2OIypxR//YEb3fkDn
# 3UayWW9bAgMBAAGjggFkMIIBYDAfBgNVHSMEGDAWgBQy65Ka/zWWSC8oQEJwIDaR
# XBeF5jAdBgNVHQ4EFgQUDyrLIIcouOxvSK4rVKYpqhekzQwwDgYDVR0PAQH/BAQD
# AgGGMBIGA1UdEwEB/wQIMAYBAf8CAQAwEwYDVR0lBAwwCgYIKwYBBQUHAwMwGwYD
# VR0gBBQwEjAGBgRVHSAAMAgGBmeBDAEEATBLBgNVHR8ERDBCMECgPqA8hjpodHRw
# Oi8vY3JsLnNlY3RpZ28uY29tL1NlY3RpZ29QdWJsaWNDb2RlU2lnbmluZ1Jvb3RS
# NDYuY3JsMHsGCCsGAQUFBwEBBG8wbTBGBggrBgEFBQcwAoY6aHR0cDovL2NydC5z
# ZWN0aWdvLmNvbS9TZWN0aWdvUHVibGljQ29kZVNpZ25pbmdSb290UjQ2LnA3YzAj
# BggrBgEFBQcwAYYXaHR0cDovL29jc3Auc2VjdGlnby5jb20wDQYJKoZIhvcNAQEM
# BQADggIBAAb/guF3YzZue6EVIJsT/wT+mHVEYcNWlXHRkT+FoetAQLHI1uBy/YXK
# ZDk8+Y1LoNqHrp22AKMGxQtgCivnDHFyAQ9GXTmlk7MjcgQbDCx6mn7yIawsppWk
# vfPkKaAQsiqaT9DnMWBHVNIabGqgQSGTrQWo43MOfsPynhbz2Hyxf5XWKZpRvr3d
# MapandPfYgoZ8iDL2OR3sYztgJrbG6VZ9DoTXFm1g0Rf97Aaen1l4c+w3DC+IkwF
# kvjFV3jS49ZSc4lShKK6BrPTJYs4NG1DGzmpToTnwoqZ8fAmi2XlZnuchC4NPSZa
# PATHvNIzt+z1PHo35D/f7j2pO1S8BCysQDHCbM5Mnomnq5aYcKCsdbh0czchOm8b
# kinLrYrKpii+Tk7pwL7TjRKLXkomm5D1Umds++pip8wH2cQpf93at3VDcOK4N7Ew
# oIJB0kak6pSzEu4I64U6gZs7tS/dGNSljf2OSSnRr7KWzq03zl8l75jy+hOds9TW
# SenLbjBQUGR96cFr6lEUfAIEHVC1L68Y1GGxx4/eRI82ut83axHMViw1+sVpbPxg
# 51Tbnio1lB93079WPFnYaOvfGAA0e0zcfF/M9gXr+korwQTh2Prqooq2bYNMvUoU
# KD85gnJ+t0smrWrb8dee2CvYZXD5laGtaAxOfy/VKNmwuWuAh9kcMIIGWzCCBMOg
# AwIBAgIRAMk34xEfUn56n/sz5LWjiPQwDQYJKoZIhvcNAQEMBQAwVDELMAkGA1UE
# BhMCR0IxGDAWBgNVBAoTD1NlY3RpZ28gTGltaXRlZDErMCkGA1UEAxMiU2VjdGln
# byBQdWJsaWMgQ29kZSBTaWduaW5nIENBIFIzNjAeFw0yMjEwMDQwMDAwMDBaFw0y
# MzEwMDQyMzU5NTlaMFMxCzAJBgNVBAYTAlVTMRAwDgYDVQQIDAdHZW9yZ2lhMRgw
# FgYDVQQKDA9OaWNob2xhcyBHaWJzb24xGDAWBgNVBAMMD05pY2hvbGFzIEdpYnNv
# bjCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBAM0H3/vArf2jWjqAUzbK
# mjl7zqi/YEsPrXiBLxM8DBxqYaDKoDsFqgepPTPlHjvvnSJRc5AxMbwQfaig2e25
# VhcrjPFBB+/V8kWl4UL1v+dc7ho48YNpcatxSHtl8iCs3JiTy5pFSYQcQif4Shxi
# iVA1AP8I7RcoY55JAnj4XRtJIFl8URENn+1RAng+6x9sBGw9AIN/6v4gD2uS4gB8
# mIYD6jz29Oc5+iDW+nzI2NyIGDo3FcF+fjFDDsltiIt7FGbN0t/HrJSv9stEkTBd
# /on7jtCyuZwJZUcBf0ipS4hOAbKCPVG8Xokf63CleKU3nXZ2D27iSNTaGeH+pFLT
# LrB5oXdHMwK4x2ceQddo4hYoZ5GmCKHRlgslqFWCxKKW72J0GMqoE3jvWCH9H57q
# ps559mZnRylxIN+PDec3NorhRe4LlR2AG9tlVUqf8SXOl4/q1drTMQ2hMR5PC3cS
# D+HEy5Yer500d58a3oFIo8P/RaMqUppAXIO7WjQRZ6fgHOVocfBogQWXwhppt57c
# DOg6e6FL1yJrPsWmlD6/NgdroltxkiDxVaHg41UiCxzcpTyVK7Oigr+88jg2IjoA
# k9amwy+6wRFfZ5SynntQqCEl5WPq6Al6OcW0RubaOMeJ85y0Ak6V1eu8FkE1Ritk
# kT5ENFlU25pXAcZd6k/ooJr9AgMBAAGjggGnMIIBozAfBgNVHSMEGDAWgBQPKssg
# hyi47G9IritUpimqF6TNDDAdBgNVHQ4EFgQUwUfhr0xfojulOc9js0512tWYg6Mw
# DgYDVR0PAQH/BAQDAgeAMAwGA1UdEwEB/wQCMAAwEwYDVR0lBAwwCgYIKwYBBQUH
# AwMwSgYDVR0gBEMwQTA1BgwrBgEEAbIxAQIBAwIwJTAjBggrBgEFBQcCARYXaHR0
# cHM6Ly9zZWN0aWdvLmNvbS9DUFMwCAYGZ4EMAQQBMEkGA1UdHwRCMEAwPqA8oDqG
# OGh0dHA6Ly9jcmwuc2VjdGlnby5jb20vU2VjdGlnb1B1YmxpY0NvZGVTaWduaW5n
# Q0FSMzYuY3JsMHkGCCsGAQUFBwEBBG0wazBEBggrBgEFBQcwAoY4aHR0cDovL2Ny
# dC5zZWN0aWdvLmNvbS9TZWN0aWdvUHVibGljQ29kZVNpZ25pbmdDQVIzNi5jcnQw
# IwYIKwYBBQUHMAGGF2h0dHA6Ly9vY3NwLnNlY3RpZ28uY29tMBwGA1UdEQQVMBOB
# EXNhcHBoQHNhcHBoLnRvb2xzMA0GCSqGSIb3DQEBDAUAA4IBgQATNVNyEXjn00Cw
# 6fgqhamVMwm/Mi6MtaPqQPaQVj9jLjmfxZDF6E4c1SP+Ezj6m+3Ho82ma8lwOWcB
# Uan7ujg/curKLsQALgs54vxsAhwNlA0zHMnW9R507MkuEQiyrLNcvtgo47NFY9nh
# aFowS2ZkAqaR1z0cbpAQjSlxo448yUTOcQZ6MRKBUGq2d2WvxqGFmwSRLzkTKlsV
# DuuMgqdxN0E7uVVl0ZBCFLMtYz4QsRCuB1J4eXFa8yebSYXSVvsX45uFKBHRewRD
# XcfkXljsh6aKDu15fLSPHyXIpJaTx1f/Q53fDR9WWAkOKvGXworW67otjeKg88MJ
# ByVXMxmekqmWugfaCjKi3B4mtm17l/itCezXr24rkeuiPMWSB3k4OM15cr06HaeS
# ti9leQHRs2pBAnxFQ+uruRqaG51F+CrXObXNXi+hrAlpzmLwV5ZKlvhDoN1XiINE
# hdJL/jDlZ/TjDT8iizeJQVHVDH8IVRcTA7uRpIMdTghm2PQ/5GgxggMKMIIDBgIB
# ATBpMFQxCzAJBgNVBAYTAkdCMRgwFgYDVQQKEw9TZWN0aWdvIExpbWl0ZWQxKzAp
# BgNVBAMTIlNlY3RpZ28gUHVibGljIENvZGUgU2lnbmluZyBDQSBSMzYCEQDJN+MR
# H1J+ep/7M+S1o4j0MAkGBSsOAwIaBQCgeDAYBgorBgEEAYI3AgEMMQowCKACgACh
# AoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsxDjAM
# BgorBgEEAYI3AgEVMCMGCSqGSIb3DQEJBDEWBBQIEYQyDxaMAX9vbGsVofyvu6Fz
# MTANBgkqhkiG9w0BAQEFAASCAgB3IrBY/12/LI+iEZ5ZV0PPMm/p1VHdvD/1J9WX
# 4Vlt+xLpf1VIsMocCOctMFo8pa1aT28twvdaNcePS5V6fMkJ/bQwI2/gayasbgn8
# 9rg897H1G8Sn6iNp/uzxfnnptVACnx0NsVvSPcDIXq5nKIFzTe8yFoWu19n/9spd
# 7w7MdcvE0Z6y0snkgakgbo5fDArFeZOIrvfIv/Q1kwKJqLNxlfcVs/oFfSk4S/uO
# NiYlOPsGGQkDtTkKTe/lRor+3rw7iOEGhDKr4rqB2tt4g/fGH335HQLX2YNqyr0i
# /wlau6FRpPqUnBajhZ+KqDcm55yTzdCn+FtLLsLzCMVMaWype5y/MsNlaBthHkT5
# +cNFlG8ev3Hw+G7JGFrR2qRFBC8tOBI+MydUnSkTKXqBS5H3wJyfYyCofoROjE8b
# OgT0YuRhUswZgS75o6w9sTKgWGv+deUBprx+oFAWVbWgB1PaAGQF4Lom5DeomkUJ
# KJxKG9uc+24QMAqfU5osUTKel/aStr4w00F29RS/Ek/exlG4RQPa3+Yw4i1XaMEx
# LYs8/BloW81TYc/ri2c72pRh+1BpaK7q6CG+4Gx3TSPwThNsSlfoiEvCxc7ENr8I
# Yv3M0GICP83BVSPXxhq361LsBrx3zJOdS05AP/PKnidY7FAOL0w/P8RqGhnuxzdB
# Dra+4g==
# SIG # End signature block
