param(
  [Parameter(Mandatory = $true)]
  [ValidateSet("patch", "minor", "major", "set")]
  [string]$Mode,
  [string]$Value = ""
)

$ErrorActionPreference = "Stop"

$versionFile = Join-Path $PSScriptRoot "version.py"
if (-not (Test-Path -LiteralPath $versionFile)) {
  throw "version.py file not found at $versionFile"
}

$raw = (Get-Content -LiteralPath $versionFile -Raw)
if ($raw -match 'APP_VERSION\s*=\s*"([^"]+)"') {
  $current = $Matches[1].Trim()
} elseif ($raw -match "APP_VERSION\s*=\s*'([^']+)'") {
  $current = $Matches[1].Trim()
} else {
  throw "APP_VERSION not found in $versionFile"
}

if ($current -notmatch '^(\d+)\.(\d+)\.(\d+)(?:[\-+][A-Za-z0-9\.\-]+)?$') {
  throw "Current APP_VERSION is invalid: $current"
}

if ($Mode -eq "set") {
  if ([string]::IsNullOrWhiteSpace($Value)) {
    throw "For mode 'set', provide -Value like 2.2.0"
  }
  if ($Value -notmatch '^\d+\.\d+\.\d+(?:[\-+][A-Za-z0-9\.\-]+)?$') {
    throw "Invalid version value: $Value"
  }
  $next = $Value
} else {
  $major = [int]$Matches[1]
  $minor = [int]$Matches[2]
  $patch = [int]$Matches[3]
  switch ($Mode) {
    "patch" { $patch += 1 }
    "minor" { $minor += 1; $patch = 0 }
    "major" { $major += 1; $minor = 0; $patch = 0 }
  }
  $next = "$major.$minor.$patch"
}

$updated = [regex]::Replace($raw, 'APP_VERSION\s*=\s*"([^"]+)"', "APP_VERSION = `"$next`"")
if ($updated -eq $raw) {
  $updated = [regex]::Replace($raw, "APP_VERSION\s*=\s*'([^']+)'", "APP_VERSION = `"$next`"")
}
if ($updated -eq $raw) {
  throw "Could not update APP_VERSION assignment in $versionFile"
}

Set-Content -LiteralPath $versionFile -Value $updated -Encoding UTF8
Write-Output "APP_VERSION updated: $current -> $next"
