# Write a fully fixed PowerShell script (PS 5.1), clean ASCII, correct #requires newline,
# proper string quoting (so '&' is inside quotes), + manual overrides + interactive menu fallback.
script = r"""#requires -Version 5.1
<#
.SYNOPSIS
  Download a ZIP (or per-user ZIPs) from Google Drive, locate the correct JPG/PNG for the current user,
  and deploy it as the default Outlook signature (classic Outlook). If no auto-match is found, the script
  will prompt the user with a numbered list of all images found to choose from.

.CHANGELOG
  2025-09-23: PS 5.1-only networking (Invoke-WebRequest + cookies) for Google Drive.
  2025-09-23: Added SAM→FirstnameLastname fuzzy match.
  2025-09-23: Added inner per-user ZIP support.
  2025-09-23: Stream the target image directly from the ZIP (no massive temp extraction).
  2025-09-24: Manual overrides for 'pibrodie' → 'PierreBrodie' and 'pbrodie' → 'PatriciaBrodie'.
  2025-09-24: PS 5.1 compatibility (no ternary operator).
  2025-09-24: Interactive fallback menu (1., 2., 3., …) to select image when auto-match fails.
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
  }
  finally { $zip.Dispose() }

  return $null
}

function Prompt-SelectImageFromZip {
  param([string]$ZipPath)
  Add-ZipAssemblies
  $exts = @(".jpg",".jpeg",".png")

  $zip = [System.IO.Compression.ZipFile]::OpenRead($ZipPath)
  try {
    $candidates = New-Object System.Collections.Generic.List[object]
    $index = 1

    foreach ($e in $zip.Entries) {
      if ($e.Length -eq 0) { continue }
      $ext = [IO.Path]::GetExtension($e.Name)
      if ($exts -icontains $ext) {
        $display = $e.FullName
        Write-Host ("  {0}. {1}" -f $index, $display)
        $candidates.Add([pscustomobject]@{ Type="Outer"; Entry=$e; Ext=$ext })
        $index++
      }
    }

    foreach ($outer in $zip.Entries) {
      if ($outer.Length -eq 0) { continue }
      if ([IO.Path]::GetExtension($outer.Name) -ieq ".zip") {
        $ms = New-Object IO.MemoryStream
        $s = $outer.Open()
        try { $s.CopyTo($ms) } finally { $s.Dispose() }
        $ms.Position = 0

        $inner = New-Object System.IO.Compression.ZipArchive($ms,[System.IO.Compression.ZipArchiveMode]::Read,$false)
        try {
          foreach ($ie in $inner.Entries) {
            if ($ie.Length -eq 0) { continue }
            $ext = [IO.Path]::GetExtension($ie.Name)
            if ($exts -icontains $ext) {
              $display = ("{0} -> {1}" -f $outer.Name, $ie.FullName)
              Write-Host ("  {0}. {1}" -f $index, $display)
              $candidates.Add([pscustomobject]@{ Type="Inner"; Outer=$outer; InnerName=$ie.FullName; Ext=$ext })
              $index++
            }
          }
        }
        finally { $inner.Dispose(); $ms.Dispose() }
      }
    }

    if ($candidates.Count -eq 0) {
      Write-Log "No images (.jpg/.jpeg/.png) found anywhere in the ZIP." "ERROR"
      return $null
    }

    $sel = $null
    while ($true) {
      $ans = Read-Host "No auto-match found. Enter the number of the correct image (1-$($candidates.Count))"
      if ([int]::TryParse($ans, [ref]$sel)) {
        if ($sel -ge 1 -and $sel -le $candidates.Count) { break }
      }
      Write-Host "Please enter a valid number between 1 and $($candidates.Count)."
    }

    $choice = $candidates[$sel-1]
    if ($choice.Type -eq "Outer") {
      return Extract-EntryToTemp -Entry $choice.Entry -Ext $choice.Ext
    } else {
      $ms2 = New-Object IO.MemoryStream
      $s2 = $choice.Outer.Open()
      try { $s2.CopyTo($ms2) } finally { $s2.Dispose() }
      $ms2.Position = 0
      $inner2 = New-Object System.IO.Compression.ZipArchive($ms2,[System.IO.Compression.ZipArchiveMode]::Read,$false)
      try {
        $target = $inner2.GetEntry($choice.InnerName)
        if (-not $target) { throw "Chosen inner entry not found on re-open." }
        return Extract-EntryToTemp -Entry $target -Ext $choice.Ext
      }
      finally { $inner2.Dispose(); $ms2.Dispose() }
    }
  }
  finally { $zip.Dispose() }
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
    Write-Log "No automatic match found. Listing available images for manual selection..." "WARN"
    $imgTemp = Prompt-SelectImageFromZip -ZipPath $tempZip
  }
  if (-not $imgTemp) {
    throw "No image selected; aborting."
  }
  Write-Log "Chosen image staged: $imgTemp"

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
"""
path = "/mnt/data/Deploy-OutlookImageSignature_FromDrive_FIXED.ps1"
with open(path, "w", encoding="utf-8") as f:
    f.write(script)
print(path)
