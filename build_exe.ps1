param(
  [string]$VersionFile = "version.py"
)

$ErrorActionPreference = "Stop"

function Get-AppVersion {
  param([string]$Path)
  if (-not (Test-Path -LiteralPath $Path)) {
    throw "Version file '$Path' was not found."
  }
  $raw = (Get-Content -LiteralPath $Path -Raw)
  if ($raw -match 'APP_VERSION\s*=\s*"([^"]+)"') {
    $version = $Matches[1].Trim()
  } elseif ($raw -match "APP_VERSION\s*=\s*'([^']+)'") {
    $version = $Matches[1].Trim()
  } else {
    throw "Could not find APP_VERSION in $Path"
  }
  if ($version -notmatch '^\d+\.\d+\.\d+([\-+][A-Za-z0-9\.\-]+)?$') {
    throw "Invalid APP_VERSION '$version' in $Path. Expected semantic version like 2.1.0"
  }
  return $version
}

function Get-FileVersionTuple {
  param([string]$SemVer)
  $core = $SemVer.Split('-', 2)[0].Split('+', 2)[0]
  $parts = $core.Split('.')
  $major = [int]$parts[0]
  $minor = [int]$parts[1]
  $patch = [int]$parts[2]
  return @($major, $minor, $patch, 0)
}

$appVersion = Get-AppVersion -Path $VersionFile
$versionParts = Get-FileVersionTuple -SemVer $appVersion
$versionPyParts = ($versionParts -join ', ')

$tempVersionInfo = Join-Path $PSScriptRoot 'pyinstaller_version_info.txt'

@"
# UTF-8
VSVersionInfo(
  ffi=FixedFileInfo(
    filevers=($versionPyParts),
    prodvers=($versionPyParts),
    mask=0x3f,
    flags=0x0,
    OS=0x40004,
    fileType=0x1,
    subtype=0x0,
    date=(0, 0)
    ),
  kids=[
    StringFileInfo(
      [
      StringTable(
        '040904B0',
        [StringStruct('CompanyName', '120 DEGREEZ MEP ENGINEERING'),
        StringStruct('FileDescription', 'Mechanical and Plumbing Calculator'),
        StringStruct('FileVersion', '$appVersion'),
        StringStruct('InternalName', 'Mechanical and Plumbing Calculator'),
        StringStruct('LegalCopyright', 'Copyright (C) 2026'),
        StringStruct('OriginalFilename', 'Mechanical and Plumbing Calculator.exe'),
        StringStruct('ProductName', 'Mechanical and Plumbing Calculator'),
        StringStruct('ProductVersion', '$appVersion')])
      ]),
    VarFileInfo([VarStruct('Translation', [1033, 1200])])
  ]
)
"@ | Set-Content -LiteralPath $tempVersionInfo -Encoding UTF8

Write-Output "Building Mechanical and Plumbing Calculator v$appVersion"

python -m PyInstaller `
  --noconfirm `
  --clean `
  --onefile `
  --windowed `
  --name "Mechanical and Plumbing Calculator" `
  --icon "runoff-icon.ico" `
  --version-file "$tempVersionInfo" `
  --add-data "app;app" `
  main.py

Write-Output "Build finished. EXE version: $appVersion"
