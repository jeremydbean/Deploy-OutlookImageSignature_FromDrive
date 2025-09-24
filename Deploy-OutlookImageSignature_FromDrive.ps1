#requires -Version 5.1
<#
.SYNOPSIS
  Download a file from Google Drive by FileId or full sharing URL.
  Handles the "scan warning / confirm token" page and keeps cookies.
#>
param(
  [string]$FileId,
  [string]$Url,
  [string]$OutFile,
  [switch]$Overwrite
)

function Write-Log([string]$m,[string]$lvl="INFO"){
  $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
  Write-Host "[$lvl ] $m"
}

function Ensure-Tls12 {
  try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 -bor [Net.SecurityProtocolType]::Tls13 } catch {
    try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}
  }
}

function Get-DriveFileIdFromUrl([string]$u){
  if ([string]::IsNullOrWhiteSpace($u)) { return $null }
  $m = [regex]::Match($u, 'file/d/([^/]+)/')
  if ($m.Success) { return $m.Groups[1].Value }
  $m = [regex]::Match($u, '(?:\?|&)id=([^&]+)')
  if ($m.Success) { return $m.Groups[1].Value }
  return $null
}

function Get-FilenameFromContentDisposition([string]$cd){
  if ([string]::IsNullOrWhiteSpace($cd)) { return $null }
  $m = [regex]::Match($cd, 'filename\*=UTF-8''[^;]*?([^;]+)')
  if ($m.Success) {
    try { return [System.Uri]::UnescapeDataString($m.Groups[1].Value) } catch {}
  }
  $m2 = [regex]::Match($cd, 'filename="?([^";]+)"?')
  if ($m2.Success) { return $m2.Groups[1].Value }
  return $null
}

function Try-GetConfirmToken([string]$html,[string]$id){
  if ([string]::IsNullOrWhiteSpace($html)) { return $null }
  $m = [regex]::Match($html, 'confirm=([^&"]+).*?id=' + [regex]::Escape($id))
  if ($m.Success) { return $m.Groups[1].Value }
  $m = [regex]::Match($html, 'name="confirm"\s+value="([^"]+)"')
  if ($m.Success) { return $m.Groups[1].Value }
  return $null
}

function Test-IsZipFile([string]$Path){
  try {
    $fs = [IO.File]::OpenRead($Path)
    try {
      $buf = New-Object byte[] 4
      [void]$fs.Read($buf,0,4)
      $sig = [System.Text.Encoding]::ASCII.GetString($buf)
      return ($sig -eq ("PK" + [char]3 + [char]4)) -or ($sig -eq ("PK" + [char]5 + [char]6)) -or ($sig -eq ("PK" + [char]7 + [char]8))
    } finally { $fs.Dispose() }
  } catch { return $false }
}

try {
  Ensure-Tls12

  if (-not $FileId -and $Url) {
    $FileId = Get-DriveFileIdFromUrl $Url
    if (-not $FileId) { throw "Could not extract FileId from -Url. Provide a valid Drive URL or use -FileId directly." }
  }
  if (-not $FileId) { throw "Provide -FileId or -Url." }

  $base = "https://drive.google.com/uc?export=download&id=$FileId"
  Write-Log "Requesting: $base"

  $resp1 = Invoke-WebRequest -UseBasicParsing -Uri $base -SessionVariable sess -Headers @{
    "User-Agent"="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"
  } -ErrorAction Stop

  $proposedName = $null
  if ($resp1.Headers["Content-Disposition"]) {
    $proposedName = Get-FilenameFromContentDisposition $resp1.Headers["Content-Disposition"]
  }
  if (-not $OutFile) {
    $OutFile = if ($proposedName) { Join-Path (Get-Location) $proposedName } else { Join-Path (Get-Location) ($FileId + ".zip") }
  }

  if ((Test-Path $OutFile) -and (-not $Overwrite)) { throw "OutFile already exists: $OutFile (use -Overwrite to replace)" }

  if ($resp1.Headers["Content-Disposition"]) {
    Write-Log "Direct download available; saving to: $OutFile"
    Invoke-WebRequest -UseBasicParsing -Uri $base -OutFile $OutFile -WebSession $sess -Headers @{
      "User-Agent"="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"
    } -ErrorAction Stop | Out-Null
  } else {
    $token = Try-GetConfirmToken -html $resp1.Content -id $FileId
    $url2 = if ($token) { "https://drive.google.com/uc?export=download&confirm=$token&id=$FileId" } else { $base }
    if ($token) { Write-Log "Confirm token acquired; proceeding with final download." } else { Write-Log "No confirm token found; attempting final download anyway (may fail)." "WARN" }
    Invoke-WebRequest -UseBasicParsing -Uri $url2 -OutFile $OutFile -WebSession $sess -Headers @{
      "User-Agent"="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"
    } -ErrorAction Stop | Out-Null
  }

  if (-not (Test-Path $OutFile)) { throw "Download failed (no file at $OutFile)." }

  $bytes = (Get-Item $OutFile).Length
  Write-Log ("Downloaded to: {0} ({1:N0} bytes)" -f $OutFile, $bytes)

  if ($OutFile.ToLower().EndsWith(".zip") -and -not (Test-IsZipFile $OutFile)) {
    Write-Log "File does not look like a ZIP. Share link may not be public or Drive returned HTML." "WARN"
  }

  exit 0
}
catch {
  Write-Log ("ERROR: " + $_.Exception.Message) "ERROR"
  if ($FileId -and $OutFile -and -not (Test-Path $OutFile)) {
    try {
      Write-Log "Fallback: BITS transfer…" "WARN"
      Start-BitsTransfer -Source ("https://drive.google.com/uc?export=download&id=" + $FileId) -Destination $OutFile -ErrorAction Stop
      if (Test-Path $OutFile) {
        $bytes = (Get-Item $OutFile).Length
        Write-Log ("Downloaded via BITS to: {0} ({1:N0} bytes)" -f $OutFile, $bytes) "WARN"
        exit 0
      }
    } catch {
      Write-Log ("BITS fallback failed: " + $_.Exception.Message) "ERROR"
    }
  }
  exit 1
}
