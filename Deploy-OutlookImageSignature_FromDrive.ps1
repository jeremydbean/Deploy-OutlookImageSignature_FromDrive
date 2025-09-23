#requires -Version 5.1
<#
.SYNOPSIS
  Download a ZIP (or per-user ZIPs) from Google Drive, extract the proper JPG/PNG for the current user,
  and deploy it as the default Outlook signature (classic Outlook).

.CHANGELOG
  2025-09-23: Replaced System.Net.Http client with Invoke-WebRequest + cookie session (PS 5.1).
  2025-09-23: Added support to match inner ZIP file names by SAM (domain format: first initial + last name).
  2025-09-23: Added matching for FirstnameLastname.* based on DisplayName (e.g., 'Kesha Blevins' -> 'KeshaBlevins.jpg').

.DESCRIPTION
  Handles two packaging styles:
   A) Single ZIP that contains all users' images (flat or in subfolders).
   B) Single ZIP that contains many per-user ZIPs (e.g., jbean.zip, KeshaBlevins.zip), each with that user's images.
  The script will pick the correct per-user asset by matching:
     Preferred: -ImagePattern (UPN | SAM | DisplayName)
     Fallbacks: the other keys and derived variants, including FirstnameLastname collapsed.
  File types supported: .jpg, .jpeg, .png

.PARAMETER GoogleDriveFileId
  The file ID from Google Drive (from a URL like https://drive.google.com/file/d/<ID>/view).

.PARAMETER SignatureName
  The logical name of the signature to create.

.PARAMETER ImagePattern
  Which naming pattern to try first: UPN, SAM, DisplayName. The script will then automatically fall back to the others.
  Default is SAM (first initial + lastname).

.PARAMETER ForceKillOutlook
  Close Outlook if running so the signature is picked up immediately.

.EXAMPLE
  .\Deploy-OutlookImageSignature_FromDrive.ps1 -GoogleDriveFileId "1yg6cqUoLf1Zw5LdjNf59NCeyEJZgOmXo" -SignatureName "Company-Standard" -ImagePattern SAM -ForceKillOutlook

.NOTES
  - Works for classic Outlook (Office 16.0). The new Outlook (Monarch) does not use local signatures.
  - Ensure devices can reach Google Drive download endpoints.
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
    try {
      [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    } catch {}
  }
}

function To-TitleCase([string]$s) {
  if ([string]::IsNullOrWhiteSpace($s)) { return $null }
  $s = $s.ToLower()
  return ([CultureInfo]::InvariantCulture).TextInfo.ToTitleCase($s)
}

function Get-UserContext {
  # Build a rich context: UPN, SAM, DisplayName, plus name-derived variants like FirstLastCollapsed.
  $upn = $null
  try {
    $upn = (whoami /upn) 2>$null
    if ([string]::IsNullOrWhiteSpace($upn)) { $upn = $null }
  } catch {}
  if (-not $upn -and $env:USERDNSDOMAIN) {
    $upn = "$($env:USERNAME)@$($env:USERDNSDOMAIN)"
  }

  $sam = $env:USERNAME

  $displayName = $null
  try {
    $regPath = "HKCU:\Software\Microsoft\Office\16.0\Common\Identity"
    $displayName = (Get-ItemProperty -Path $regPath -Name "FriendlyName" -ErrorAction Stop).FriendlyName
  } catch {
    try {
      $displayName = (Get-CimInstance Win32_ComputerSystem).UserName
      if ($displayName -and $displayName.Contains("\")) { $displayName = $displayName.Split("\")[-1] }
    } catch {}
  }
  if (-not $displayName) {
    if ($upn -and $upn.Contains("@")) { $displayName = $upn.Split("@")[0] } else { $displayName = $sam }
  }

  # Derive compact name variants from DisplayName (e.g., "Kesha Blevins" -> "KeshaBlevins")
  $firstLastCompact = $null
  $fullCompact = $null
  try {
    $tokens = [regex]::Split($displayName, "[^A-Za-z]+") | Where-Object { $_ -ne "" }
    if ($tokens.Count -ge 2) {
      $first = To-TitleCase $tokens[0]
      $last  = To-TitleCase $tokens[-1]
      $firstLastCompact = "$first$last"
      # Also a compact of all tokens, e.g., "KeshaMarieBlevins"
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
    $url2 = "https://drive.google.com/uc?export=download&confirm=$confirmToken&id=$FileId"
  }

  # Stage 2: final download
  Invoke-WebRequest -Uri $url2 -OutFile $DestinationPath -UseBasicParsing -WebSession $gdsess -Headers @{
    "User-Agent"="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"
  } -ErrorAction Stop | Out-Null

  if (-not (Test-Path $DestinationPath)) {
    throw "Google Drive download failed; destination file not found."
  }

  return $DestinationPath
}

function Expand-ZipToTemp {
  param([string]$ZipPath)
  $tempRoot = Join-Path $env:TEMP ("SigZip_" + [Guid]::NewGuid().ToString("N"))
  New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null
  Add-Type -AssemblyName System.IO.Compression.FileSystem
  [System.IO.Compression.ZipFile]::ExtractToDirectory($ZipPath, $tempRoot)
  return $tempRoot
}

function Expand-ZipFile {
  param(
    [string]$ZipPath,
    [string]$TargetFolder
  )
  Add-Type -AssemblyName System.IO.Compression.FileSystem
  [System.IO.Compression.ZipFile]::ExtractToDirectory($ZipPath, $TargetFolder)
}

function Get-FallbackKeyOrder {
  param([string]$Preferred)
  $all = @("UPN","SAM","DisplayName","FirstLastCompact","FullCompact")
  @($Preferred) + ($all | Where-Object { $_ -ne $Preferred }) | Select-Object -Unique
}

function Try-FindInnerUserZip {
  param(
    [string]$Folder,
    [pscustomobject]$Ctx,
    [string]$PreferredPattern
  )

  $order = Get-FallbackKeyOrder -Preferred $PreferredPattern
  $zips = Get-ChildItem -LiteralPath $Folder -Recurse -File -Include *.zip -ErrorAction SilentlyContinue

  foreach ($key in $order) {
    $name = $Ctx.$key
    if (-not $name) { continue }
    $cands = @(
      ($name + ".zip"),
      ($name.ToLower() + ".zip")
    )
    foreach ($z in $zips) {
      if ($cands -contains $z.Name) { return $z.FullName }
    }
  }
  return $null
}

function Find-UserImageInFolder {
  param(
    [string]$Folder,
    [pscustomobject]$Ctx,
    [string]$PreferredPattern # UPN/SAM/DisplayName/FirstLastCompact/FullCompact
  )
  $extensions = @(".jpg",".jpeg",".png")
  $order = Get-FallbackKeyOrder -Preferred $PreferredPattern

  foreach ($key in $order) {
    foreach ($ext in $extensions) {
      $name = $Ctx.$key
      if ($name) {
        $candidates = @(
          (Join-Path $Folder ($name + $ext)),
          (Join-Path $Folder ($name.ToLower() + $ext))
        )
        foreach ($p in $candidates) {
          if (Test-Path -LiteralPath $p) { return $p }
        }
        # recursive search (in case of subfolders)
        $found = Get-ChildItem -LiteralPath $Folder -Recurse -File -ErrorAction SilentlyContinue | Where-Object {
          $_.Name -ieq ($name + $ext)
        } | Select-Object -First 1
        if ($found) { return $found.FullName }
      }
    }
  }
  return $null
}

function New-HtmlSignature {
  param(
    [string]$SigFolder,
    [string]$SigBaseName,
    [string]$ImageFile
  )
  $htmlPath = Join-Path $SigFolder ($SigBaseName + ".htm")
  $rtfPath  = Join-Path $SigFolder ($SigBaseName + ".rtf")
  $txtPath  = Join-Path $SigFolder ($SigBaseName + ".txt")

  $imgFileName = [IO.Path]::GetFileName($ImageFile)
  $html = @"
<!DOCTYPE html>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<style type="text/css">
  body { margin:0; padding:0; }
  img  { border:0; display:block; }
</style>
</head>
<body>
  <img src="$imgFileName" alt="Signature" />
</body>
</html>
"@
  $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
  [IO.File]::WriteAllText($htmlPath, $html, $utf8NoBom)
  [IO.File]::WriteAllText($rtfPath, "{\rtf1\ansi\deff0 {\fonttbl {\f0 Arial;}}\f0\fs20 [Signature image not shown in RTF mode]}", [Text.Encoding]::ASCII)
  [IO.File]::WriteAllText($txtPath, "[Signature image not shown in Plain Text mode]", [Text.Encoding]::UTF8)
  return $htmlPath
}

function Set-OutlookDefaultSignature {
  param([string]$SigName)
  $mailSettings = "HKCU:\Software\Microsoft\Office\16.0\Common\MailSettings"
  if (-not (Test-Path $mailSettings)) { New-Item -Path $mailSettings -Force | Out-Null }
  New-ItemProperty -Path $mailSettings -Name "NewSignature" -Value $SigName -PropertyType String -Force | Out-Null
  New-ItemProperty -Path $mailSettings -Name "ReplySignature" -Value $SigName -PropertyType String -Force | Out-Null
  Write-Log "Set default signatures (New/Reply) to '$SigName'."
}
#endregion Helpers

#region Main
try {
  Ensure-Tls12

  $ctx = Get-UserContext
  Write-Log "User context -> UPN='$($ctx.UPN)', SAM='$($ctx.SAM)', DisplayName='$($ctx.DisplayName)', FirstLastCompact='$($ctx.FirstLastCompact)', FullCompact='$($ctx.FullCompact)'"

  # 1) Download Drive file to temp
  $tempZip = Join-Path $env:TEMP ("sigpkg_" + [Guid]::NewGuid().ToString("N") + ".zip")
  Write-Log "Downloading signature package from Google Drive..."
  Invoke-GDriveDownload -FileId $GoogleDriveFileId -DestinationPath $tempZip | Out-Null
  if (-not (Test-Path $tempZip)) { throw "Download failed: $tempZip not found." }
  Write-Log "Downloaded: $tempZip"

  # 2) Extract main ZIP
  $extractFolder = Expand-ZipToTemp -ZipPath $tempZip
  Write-Log "Extracted to: $extractFolder"

  # 3) If the main ZIP contains per-user ZIPs, pick the right one and extract it to a nested temp folder
  $userZip = Try-FindInnerUserZip -Folder $extractFolder -Ctx $ctx -PreferredPattern $ImagePattern
  if ($userZip) {
    Write-Log "Found per-user ZIP: $userZip"
    $inner = Join-Path $env:TEMP ("SigUser_" + [Guid]::NewGuid().ToString("N"))
    New-Item -ItemType Directory -Path $inner -Force | Out-Null
    Expand-ZipFile -ZipPath $userZip -TargetFolder $inner
    $searchRoot = $inner
  } else {
    $searchRoot = $extractFolder
  }

  # 4) Find image for current user within the chosen folder
  $srcImage = Find-UserImageInFolder -Folder $searchRoot -Ctx $ctx -PreferredPattern $ImagePattern
  if (-not $srcImage) {
    throw "Could not locate an image for user using UPN/SAM/DisplayName/FirstLastCompact/FullCompact (.jpg/.jpeg/.png)."
  }
  Write-Log "Matched user image: $srcImage"

  # 5) Ensure Outlook signature directory exists
  $sigFolder = Join-Path $env:APPDATA "Microsoft\Signatures"
  if (-not (Test-Path $sigFolder)) {
    New-Item -Path $sigFolder -ItemType Directory -Force | Out-Null
  }

  # 6) Copy the image into Signatures and create signature HTML
  $targetImage = Join-Path $sigFolder ([IO.Path]::GetFileName($srcImage))
  Copy-Item -LiteralPath $srcImage -Destination $targetImage -Force
  Write-Log "Copied image to: $targetImage"

  $htmlPath = New-HtmlSignature -SigFolder $sigFolder -SigBaseName $SignatureName -ImageFile $targetImage
  Write-Log "Created HTML signature: $htmlPath"

  # 7) Set defaults
  Set-OutlookDefaultSignature -SigName $SignatureName

  # 8) Optionally bounce Outlook
  if ($ForceKillOutlook) {
    $outlook = Get-Process OUTLOOK -ErrorAction SilentlyContinue
    if ($outlook) {
      Write-Log "Closing Outlook to reload signature..."
      $outlook | Stop-Process -Force
      Start-Sleep -Seconds 2
    }
  }

  Write-Log "DONE. Signature '$SignatureName' deployed."
  exit 0
}
catch {
  Write-Log ("ERROR: " + $_.Exception.Message) "ERROR"
  exit 1
}
finally {
  # Clean up temp files/folders
  try {
    if (Test-Path $tempZip) { Remove-Item -LiteralPath $tempZip -Force -ErrorAction SilentlyContinue }
    if ($extractFolder -and (Test-Path $extractFolder)) { Remove-Item -LiteralPath $extractFolder -Recurse -Force -ErrorAction SilentlyContinue }
    if ($inner -and (Test-Path $inner)) { Remove-Item -LiteralPath $inner -Recurse -Force -ErrorAction SilentlyContinue }
  } catch {}
}
#endregion Main
