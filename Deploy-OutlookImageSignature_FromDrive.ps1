#requires -Version 5.1
<#
.SYNOPSIS
  Download a ZIP (or per-user ZIPs) from Google Drive, locate the correct JPG/PNG for the current user
  WITHOUT fully extracting huge archives, and deploy it as the default Outlook signature (classic Outlook).

.CHANGELOG
  2025-09-23: PS 5.1-only networking (Invoke-WebRequest + cookies) for Google Drive.
  2025-09-23: Added SAM→FirstnameLastname fuzzy match.
  2025-09-23: Added inner per-user ZIP support.
  2025-09-23: Fast path — stream the target image directly from the ZIP (no massive temp extraction).
  2025-09-24: Manual overrides for 'pibrodie' → 'PierreBrodie' and 'pbrodie' → 'PatriciaBrodie'.
  2025-09-24: PowerShell 5.1 compatibility (removed ternary operator).
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

  $resp1 = Invoke-WebRequest -Uri $base -UseBasicParsing -SessionVariable gdsess -Headers @{
    "User-Agent"="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"
  } -ErrorAction Stop

  if ($resp1.Headers["Content-Disposition"]) {
    Invoke-WebRequest -Uri $base -OutFile $DestinationPath -UseBasicParsing -WebSession $gdsess -Headers @{
      "User-Agent"="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"
    } -ErrorAction Stop | Out-Null
    return $DestinationPath
  }

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

  Invoke-WebRequest -Uri $url2 -OutFile $DestinationPath -UseBasicParsing -WebSession $gdsess -Headers @{
    "User-Agent"="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"
  } -ErrorAction Stop | Out-Null

  if (-not (Test-Path $DestinationPath)) {
    throw "Google Drive download failed; destination file not found."
  }

  return $DestinationPath
}

function Test-IsZipFile {
  param([string]$Path)
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

function Add-ZipAssemblies {
  try { Add-Type -AssemblyName System.IO.Compression | Out-Null } catch {}
  try { Add-Type -AssemblyName System.IO.Compression.FileSystem | Out-Null } catch {}
}

function Get-FallbackKeyOrder {
  param([string]$Preferred)
  $all = @("UPN","SAM","DisplayName","FirstLastCompact","FullCompact")
  @($Preferred) + ($all | Where-Object { $_ -ne $Preferred }) | Select-Object -Unique
}

function Get-InitialLastFromSam {
  param([string]$Sam)
  if ([string]::IsNullOrWhiteSpace($Sam) -or $Sam.Length -lt 2) { return $null }
  $s = ($Sam -replace "[^A-Za-z]","")
  if ($s.Length -lt 2) { return $null }
  $initial = $s.Substring(0,1).ToUpper()
  $lastLower = $s.Substring(1).ToLower()
  [pscustomobject]@{ Initial = $initial; LastLower = $lastLower }
}

function New-RegexFromInitialLast {
  param([string]$Initial,[string]$LastLower,[switch]$ZipMode)
  $tail = $null
  if ($ZipMode) { $tail = '\.zip$' } else { $tail = '\.(jpg|jpeg|png)$' }
  return "^(?i)" + [regex]::Escape($Initial) + "[a-z]+" + [regex]::Escape($LastLower) + $tail
}

function Get-ManualOverrideNames {
  param([string]$Sam)
  $names = @()
  if ($Sam -ieq "pibrodie") { $names += "PierreBrodie" }
  if ($Sam -ieq "pbrodie")  { $names += "PatriciaBrodie" }
  return $names
}

function Extract-EntryToTemp {
  param([System.IO.Compression.ZipArchiveEntry]$Entry,[string]$Ext)
  $target = Join-Path $env:TEMP ("sigimg_" + [Guid]::NewGuid().ToString("N") + $Ext)
  $in = $Entry.Open()
  try {
    $out = [IO.File]::Open($target,[IO.FileMode]::Create,[IO.FileAccess]::Write,[IO.FileShare]::None)
    try { $in.CopyTo($out) } finally { $out.Dispose() }
  } finally { $in.Dispose() }
  return $target
}

function Find-ImageFromZip {
  param(
    [string]$ZipPath,
    [pscustomobject]$Ctx,
    [string]$PreferredPattern
  )
  Add-ZipAssemblies

  $exts = @(".jpg",".jpeg",".png")
  $order = Get-FallbackKeyOrder -Preferred $PreferredPattern
  $manual = Get-ManualOverrideNames -Sam $Ctx.SAM

  $zip = [System.IO.Compression.ZipFile]::OpenRead($ZipPath)
  try {
    if ($manual.Count -gt 0) {
      foreach ($e in $zip.Entries) {
        if ($e.Length -eq 0) { continue }
        $base = [IO.Path]::GetFileNameWithoutExtension($e.Name)
        $ext  = [IO.Path]::GetExtension($e.Name)
        if ($exts -icontains $ext -and ($manual -icontains $base)) {
          Write-Log "Manual override matched '$($e.FullName)'."
          return Extract-EntryToTemp -Entry $e -Ext $ext
        }
      }
    }

    foreach ($key in $order) {
      $name = $Ctx.$key
      if ([string]::IsNullOrWhiteSpace($name)) { continue }
      foreach ($e in $zip.Entries) {
        if ($e.Length -eq 0) { continue }
        $base = [IO.Path]::GetFileNameWithoutExtension($e.Name)
        $ext  = [IO.Path]::GetExtension($e.Name)
        if ($exts -icontains $ext -and ($base -ieq $name -or $base -ieq $name.ToLower())) {
          Write-Log "Matched '$($e.FullName)' by key '$key'."
          return Extract-EntryToTemp -Entry $e -Ext $ext
        }
      }
    }

    $parts = Get-InitialLastFromSam -Sam $Ctx.SAM
    if ($parts) {
      $rx = New-RegexFromInitialLast -Initial $parts.Initial -LastLower $parts.LastLower
      foreach ($e in $zip.Entries) {
        if ($e.Length -eq 0) { continue }
        $base = [IO.Path]::GetFileName($e.Name)
        if ($base -match $rx) {
          $ext  = [IO.Path]::GetExtension($base)
          Write-Log "Fuzzy matched '$($e.FullName)' via '$rx'."
          return Extract-EntryToTemp -Entry $e -Ext $ext
        }
      }
    }

    if ($manual.Count -gt 0) {
      foreach ($e in $zip.Entries) {
        if ($e.Length -eq 0) { continue }
        if ([IO.Path]::GetExtension($e.Name) -ieq ".zip") {
          $base = [IO.Path]::GetFileNameWithoutExtension($e.Name)
          if ($manual -icontains $base) {
            $candidate = Find-ImageInsideInnerZip -OuterEntry $e -Ctx $Ctx -Order $order -Manual $manual
            if ($candidate) { return $candidate }
          }
        }
      }
    }

    foreach ($key in $order) {
      $name = $Ctx.$key
      if ([string]::IsNullOrWhiteSpace($name)) { continue }
      foreach ($e in $zip.Entries) {
        if ($e.Length -eq 0) { continue }
        if ([IO.Path]::GetExtension($e.Name) -ieq ".zip") {
          $base = [IO.Path]::GetFileNameWithoutExtension($e.Name)
          if ($base -ieq $name -or $base -ieq $name.ToLower()) {
            $candidate = Find-ImageInsideInnerZip -OuterEntry $e -Ctx $Ctx -Order $order -Manual $manual
            if ($candidate) { return $candidate }
          }
        }
      }
    }

    if ($parts) {
      $zrx = New-RegexFromInitialLast -Initial $parts.Initial -LastLower $parts.LastLower -ZipMode
      foreach ($e in $zip.Entries) {
        if ($e.Length -eq 0) { continue }
        if ([IO.Path]::GetExtension($e.Name) -ieq ".zip") {
          $name = [IO.Path]::GetFileName($e.Name)
          if ($name -match $zrx) {
            $candidate = Find-ImageInsideInnerZip -OuterEntry $e -Ctx $Ctx -Order $order -Manual $manual
            if ($candidate) { return $candidate }
          }
        }
      }
    }
  }
  finally { $zip.Dispose() }

  return $null
}

function Find-ImageInsideInnerZip {
  param(
    [System.IO.Compression.ZipArchiveEntry]$OuterEntry,
    [pscustomobject]$Ctx,
    [string[]]$Order,
    [string[]]$Manual
  )
  Add-ZipAssemblies
  $exts = @(".jpg",".jpeg",".png")

  $ms = New-Object IO.MemoryStream
  $s = $OuterEntry.Open()
  try { $s.CopyTo($ms) } finally { $s.Dispose() }
  $ms.Position = 0

  $inner = New-Object System.IO.Compression.ZipArchive($ms,[System.IO.Compression.ZipArchiveMode]::Read,$false)
  try {
    if ($Manual -and $Manual.Count -gt 0) {
      foreach ($e in $inner.Entries) {
        if ($e.Length -eq 0) { continue }
        $base = [IO.Path]::GetFileNameWithoutExtension($e.Name)
        $ext  = [IO.Path]::GetExtension($e.Name)
        if ($exts -icontains $ext -and ($Manual -icontains $base)) {
          Write-Log "Manual override matched inner '$($e.FullName)'."
          return Extract-EntryToTemp -Entry $e -Ext $ext
        }
      }
    }

    foreach ($key in $Order) {
      $name = $Ctx.$key
      if ([string]::IsNullOrWhiteSpace($name)) { continue }
      foreach ($e in $inner.Entries) {
        if ($e.Length -eq 0) { continue }
        $base = [IO.Path]::GetFileNameWithoutExtension($e.Name)
        $ext  = [IO.Path]::GetExtension($e.Name)
        if ($exts -icontains $ext -and ($base -ieq $name -or $base -ieq $name.ToLower())) {
          Write-Log "Matched inner '$($e.FullName)' by key '$key'."
          return Extract-EntryToTemp -Entry $e -Ext $ext
        }
      }
    }

    $parts = Get-InitialLastFromSam -Sam $Ctx.SAM
    if ($parts) {
      $rx = New-RegexFromInitialLast -Initial $parts.Initial -LastLower $parts.LastLower
      foreach ($e in $inner.Entries) {
        if ($e.Length -eq 0) { continue }
        $base = [IO.Path]::GetFileName($e.Name)
        if ($base -match $rx) {
          $ext  = [IO.Path]::GetExtension($base)
          Write-Log "Fuzzy matched inner '$($e.FullName)' via '$rx'."
          return Extract-EntryToTemp -Entry $e -Ext $ext
        }
      }
    }
  }
  finally { $inner.Dispose(); $ms.Dispose() }
  return $null
}

function New-HtmlSignature {
  param([string]$SigFolder,[string]$SigBaseName,[string]$ImageFile)
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

  $tempZip = Join-Path $env:TEMP ("sigpkg_" + [Guid]::NewGuid().ToString("N") + ".zip")
  Write-Log "Downloading signature package from Google Drive..."
  Invoke-GDriveDownload -FileId $GoogleDriveFileId -DestinationPath $tempZip | Out-Null
  if (-not (Test-Path $tempZip)) { throw "Download failed: $tempZip not found." }
  $size = (Get-Item $tempZip).Length
  Write-Log ("Downloaded: {0} ({1:N0} bytes)" -f $tempZip, $size)

  if (-not (Test-IsZipFile -Path $tempZip)) {
    throw "Downloaded file is not a ZIP (bad token). The Drive link may require auth or a new confirm token."
  }

  $imgTemp = Find-ImageFromZip -ZipPath $tempZip -Ctx $ctx -PreferredPattern $ImagePattern
  if (-not $imgTemp) {
    throw "Could not locate an image for user using UPN/SAM/DisplayName/FirstLastCompact/FullCompact (.jpg/.jpeg/.png)."
  }
  Write-Log "Staged image: $imgTemp"

  $sigFolder = Join-Path $env:APPDATA "Microsoft\Signatures"
  if (-not (Test-Path $sigFolder)) {
    New-Item -Path $sigFolder -ItemType Directory -Force | Out-Null
  }

  $targetImage = Join-Path $sigFolder ([IO.Path]::GetFileName($imgTemp))
  Copy-Item -LiteralPath $imgTemp -Destination $targetImage -Force
  Write-Log "Copied image to: $targetImage"

  $htmlPath = New-HtmlSignature -SigFolder $sigFolder -SigBaseName $SignatureName -ImageFile $targetImage
  Write-Log "Created HTML signature: $htmlPath"

  Set-OutlookDefaultSignature -SigName $SignatureName

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
  try {
    if (Test-Path $tempZip) { Remove-Item -LiteralPath $tempZip -Force -ErrorAction SilentlyContinue }
    if ($imgTemp -and (Test-Path $imgTemp)) { Remove-Item -LiteralPath $imgTemp -Force -ErrorAction SilentlyContinue }
  } catch {}
}
#endregion Main
