<#
************ MANAGEENGINE ENDPOINT CENTRAL AGENT INSTALLATION (Download then Install) ************
 
Usage:
    .\ECAgentInstall.ps1    # runs with embedded URL
 
Notes:
    - Requires admin privileges.
    - Logs: C:\ProgramData\EndpointCentral\AgentInstall.log
#>
 
[CmdletBinding()]
param(
    # Default URL embedded
    [string]$AgentUrl = "https://publicresourcesalrawaf.blob.core.windows.net/public/Scripts/LocalOffice_Agent.exe",
 
    [string]$ExeFileName = "LocalOffice_Agent.exe",
    [string]$InstallSource = "GPO",
    [string]$Proxy = $null,
    [int]$DownloadRetries = 3
)
 
$ErrorActionPreference = "Stop"
 
# --- Logging setup ---
$LogDir  = Join-Path $env:ProgramData "EndpointCentral"
$LogFile = Join-Path $LogDir "AgentInstall.log"
New-Item -ItemType Directory -Path $LogDir -Force | Out-Null
function Write-Log([string]$Message, [string]$Level = "INFO") {
    $ts = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $line = "[$ts][$Level] $Message"
    Write-Output $line
    Add-Content -Path $LogFile -Value $line
}
 
Write-Log "==== Starting Endpoint Central Agent installation ===="
Write-Log "Parameters: Url='$AgentUrl' ExeFileName='$ExeFileName' InstallSource='$InstallSource' Proxy='$Proxy'"
 
# --- TLS 1.2 ---
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
 
# --- Architecture & registry detection ---
if ([System.Environment]::Is64BitOperatingSystem) {
    $regKey = 'HKLM:\SOFTWARE\Wow6432Node\AdventNet\DesktopCentral\DCAgent'
    Write-Log "64-bit OS detected"
} else {
    $regKey = 'HKLM:\SOFTWARE\AdventNet\DesktopCentral\DCAgent'
    Write-Log "32-bit OS detected"
}
 
$agentVersion = $null
if (Test-Path $regKey) {
    try {
        $agentVersion = (Get-ItemProperty $regKey).DCAgentVersion
        if ($agentVersion) {
            Write-Log "Existing agent detected. Version: $agentVersion"
        } else {
            Write-Log "Agent registry key found but version value is missing."
        }
    } catch {
        Write-Log "Failed to read agent version from registry: $($_.Exception.Message)" "WARN"
    }
} else {
    Write-Log "Agent registry key not found."
}
 
# --- Skip if installed ---
if ($agentVersion) {
    Write-Log "Skipping installation because agent appears to be installed. Exiting."
    return
}
 
# --- Prepare download location ---
$TempRoot   = $env:SystemRoot
if (-not $TempRoot) { $TempRoot = $env:TEMP }
$WorkDir    = Join-Path $TempRoot "Temp"
$LocalExe   = Join-Path $WorkDir $ExeFileName
 
New-Item -ItemType Directory -Path $WorkDir -Force | Out-Null
 
# --- Download function with retry ---
function Download-Agent {
    param(
        [string]$Url,
        [string]$Destination,
        [int]$Retries = 3,
        [string]$ProxyUrl = $null
    )
    $attempt = 0
    while ($attempt -lt $Retries) {
        $attempt++
        try {
            Write-Log "Downloading agent (attempt $attempt/$Retries) from $Url to $Destination"
            $commonParams = @{
                Uri             = $Url
                OutFile         = $Destination
                UseBasicParsing = $true
            }
            if ($ProxyUrl) {
                Write-Log "Using proxy: $ProxyUrl"
                $commonParams['Proxy'] = $ProxyUrl
                $commonParams['ProxyUseDefaultCredentials'] = $true
            }
            Invoke-WebRequest @commonParams
            if (Test-Path $Destination -PathType Leaf) {
                $size = (Get-Item $Destination).Length
                if ($size -gt 0) {
                    Write-Log "Download completed. Size: $size bytes"
                    return $true
                } else {
                    Write-Log "Downloaded file size is 0 bytes." "WARN"
                }
            } else {
                Write-Log "Downloaded file not found at $Destination" "WARN"
            }
        } catch {
            Write-Log "Download failed: $($_.Exception.Message)" "WARN"
            Start-Sleep -Seconds ([Math]::Min(10, $attempt * 3))
        }
    }
    return $false
}
 
# --- Perform download ---
$downloadOk = Download-Agent -Url $AgentUrl -Destination $LocalExe -Retries $DownloadRetries -ProxyUrl $Proxy
if (-not $downloadOk) {
    Write-Log "All download attempts failed. Aborting." "ERROR"
    throw "Failed to download agent from $AgentUrl"
}
 
# --- Install with vendor-supported switch ---
$arguments = @("/silent", "INSTALLSOURCE=$InstallSource")
Write-Log "Starting silent install via Start-Process..."
$proc = Start-Process -FilePath $LocalExe -ArgumentList $arguments -Wait -PassThru -NoNewWindow
 
Write-Log "Installer exit code: $($proc.ExitCode)"
if ($proc.ExitCode -ne 0) {
    throw "Installer returned non-zero exit code: $($proc.ExitCode)"
}
 
# --- Post-install verification ---
Start-Sleep -Seconds 5
if (Test-Path $regKey) {
    $newVersion = (Get-ItemProperty $regKey).DCAgentVersion
    if ($newVersion) {
        Write-Log "Agent installed successfully. Version: $newVersion"
    } else {
        Write-Log "Agent registry present but version missing after install." "WARN"
    }
} else {
    Write-Log "Agent registry key not found after install." "WARN"
}
 
# --- Optional cleanup ---
try {
    if (Test-Path $LocalExe) {
        Remove-Item $LocalExe -Force -ErrorAction SilentlyContinue
        Write-Log "Cleaned up installer at $LocalExe"
    }
} catch {
    Write-Log "Cleanup warning: $($_.Exception.Message)" "WARN"
}
 
Write-Log "==== Endpoint Central Agent installation script completed ===="
