#requires -Version 5.1
<#
.SYNOPSIS
  Download a ZIP (or per-user ZIPs) from Google Drive, locate the correct JPG/PNG for the current user
  WITHOUT fully extracting huge archives, and deploy it as the default Outlook signature (classic Outlook).

.CHANGELOG
  2025-09-23: PS 5.1-only networking (Invoke-WebRequest + cookies) for Google Drive.
  2025-09-23: Added SAM→FirstnameLastname fuzzy match (e.g., kblevins → ^K[a-z]+Blevins\.(jpg|jpeg|png)$).
  2025-09-23: Added inner per-user ZIP support (e.g., KeshaBlevins.zip).
  2025-09-23: Fast path — stream the target image directly from the ZIP (no massive temp extraction).
  2025-09-24: Manual overrides: 'pibrodie' → 'PierreBrodie'; 'pbrodie' → 'PatriciaBrodie' (outer images and inner zips).

.NOTES
  - Classic Outlook only. New Outlook (Monarch) ignores local signatures.
  - Run as the signed-in user (HKCU + %APPDATA%).

.PARAMETER GoogleDriveFileId
  Drive file ID of the ZIP to download.

.PARAMETER SignatureName
  Name of the signature to create/update in Outlook (default: Company-Standard).

.PARAMETER ImagePattern
  Which naming pattern to try first: UPN | SAM | DisplayName (default: SAM).

.PARAMETER ForceKillOutlook
  If set, closes Outlook to pick up the new signature immediately.
#>
param(
  [Parameter(Mandatory = $true)]
  [string]$GoogleDriveFileId,

  [Parameter(Mandatory = $false)]
  [string]$SignatureName = "Company-Standard",

  [Parameter(Mandatory = $false)]
  [ValidateSet("UPN","SAM","DisplayName")]
  [string]$ImagePattern = "SAM",

  [switch]$ForceKillOutlook
)

#region Helpers
function Write-Log {
  param([string]$Message, [string]$Level = "INFO")
  $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
  Write-Host "[$Level ] $Message"
}

function Ensure-Tls12 {
  try {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 -bor [Net.SecurityProtocolType]::Tls13
  } catch {
    try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}
  }
}

function To-TitleCase([string]$s) {
  if ([string]::IsNullOrWhiteSpace($s)) { return $null }
  $s = $s.ToLower()
  return ([System.Globalization.CultureInfo]::InvariantCulture).TextInfo.ToTitleCase($s)
}

function Get-UserContext {
  $upn = $null
  try {
    $upn = (whoami /upn) 2>$null
    if ([string]::IsNullOrWhiteSpace($upn)) { $upn = $null }
  } catch {}
  if (-not $upn -and $env:USERDNSDOMAIN) { $upn = "$($env:USERNAME)@$($env:USERDNSDOMAIN)" }

  $sam = $env:USERNAME

  $displayName = $null
  try {
    $regPath = "HKCU:\Software\Microsoft\Office\16.0\Common\Identity"
    $displayName = (Get-ItemProperty -Path $regPath -Name "FriendlyName" -ErrorAction Stop).FriendlyName
  } catch {
    try {
      $dn = (Get-CimInstance Win32_ComputerSystem).UserName
      if ($dn -and $dn.Contains("\")) { $dn = $dn.Split("\")[-1] }
      $displayName = $dn
    } catch {}
  }
  if (-not $displayName) {
    if ($upn -and $upn.Contains("@")) { $displayName = $upn.Split("@")[0] } else { $displayName = $sam }
  }

  # Derive compact variants from DisplayName (e.g., "Kesha Blevins" -> "KeshaBlevins")
  $firstLastCompact = $null
  $fullCompact = $null
  try {
    $tokens = [regex]::Split($displayName, "[^A-Za-z]+") | Where-Object { $_ -ne "" }
    if ($tokens.Count -ge 2) {
      $first = To-TitleCase $tokens[0]
      $last  = To-TitleCase $tokens[-1]
      $firstLastCompact = "$first$last"
      $fullCompact = ($tokens | ForEach-Object { To-TitleCase $_ }) -join ''
    } elseif ($tokens.Count -eq 1) {
      $fullCompact = To-TitleCase $tokens[0]
    }
  } catch {}

  [pscustomobject]@{
    UPN               = $upn
    SAM               = $sam
    DisplayName       = $displayName
    FirstLastCompact  = $firstLastCompact
    FullCompact       = $fullCompact
  }
}

function Invoke-GDriveDownload {
  param(
    [Parameter(Mandatory=$true)][string]$FileId,
    [Parameter(Mandatory=$true)][string]$DestinationPath
  )

  Ensure-Tls12
  $base = "https://drive.google.com/uc?export=download&id=$FileId"

  # Stage 1: initial request (capture cookies and maybe token)
  $resp1 = Invoke-WebRequest -Uri $base -UseBasicParsing -SessionVariable gdsess -Headers @{
    "User-Agent"="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"
  } -ErrorAction Stop

  # If Content-Disposition exists, download directly
  if ($resp1.Headers["Content-Disposition"]) {
    Invoke-WebRequest -Uri $base -OutFile $DestinationPath -UseBasicParsing -WebSession $gdsess -Headers @{
      "User-Agent"="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"
    } -ErrorAction Stop | Out-Null
    return $DestinationPath
  }

  # Try to parse confirm token from HTML body
  $html = $resp1.Content
  $confirmToken = $null

  $m = [regex]::Match($html, 'confirm=([^&"]+).*?id=' + [regex]::Escape($FileId))
  if ($m.Success) { $confirmToken = $m.Groups[1].Value }

  if (-not $confirmToken) {
    $m2 = [regex]::Match($html, 'name="confirm"\s+value="([^"]+)"')
    if ($m2.Success) { $confirmToken = $m2.Groups[1].Value }
  }

  $url2 = $base
  if ($confirmToken) {
    $url2 = "https://drive.google.com/uc?export=download&confirm=$confirmToken&id=$
