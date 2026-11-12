<#
    run-pmrs.ps1 - launch a browser with the "PMRS Report -> Excel" extension
    already loaded, and open the rejected-claims report.

    Usage:
        .\run-pmrs.ps1                 # Edge (recommended) with its own profile
        .\run-pmrs.ps1 -Browser Chrome # Chrome (opens the site; extension must be
                                        #  installed once via chrome://extensions)
        .\run-pmrs.ps1 -Url "https://pmrs.digitalsp.gov.kh/claims/approved/list"
#>
param(
    [ValidateSet('Auto', 'Edge', 'Chrome')]
    [string]$Browser = 'Auto',

    [string]$Url = 'https://pmrs.digitalsp.gov.kh/claims/rejected/list'
)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$ext = Join-Path $root 'extension'

if (-not (Test-Path (Join-Path $ext 'manifest.json'))) {
    throw "Extension folder not found: $ext"
}

function Find-Browser([string]$name) {
    $candidates = switch ($name) {
        'Edge' {
            @(
                "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe",
                "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe",
                "$env:LOCALAPPDATA\Microsoft\Edge\Application\msedge.exe"
            )
        }
        'Chrome' {
            @(
                "$env:ProgramFiles\Google\Chrome\Application\chrome.exe",
                "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe",
                "$env:LOCALAPPDATA\Google\Chrome\Application\chrome.exe"
            )
        }
    }
    foreach ($c in $candidates) { if (Test-Path $c) { return $c } }
    return $null
}

$edge = Find-Browser 'Edge'
$chrome = Find-Browser 'Chrome'

if ($Browser -eq 'Auto') {
    $Browser = if ($edge) { 'Edge' } elseif ($chrome) { 'Chrome' } else { throw 'No Edge or Chrome found.' }
}

if ($Browser -eq 'Edge') {
    if (-not $edge) { throw 'Microsoft Edge not found.' }

    $profile = Join-Path $root '.browser-profile'
    New-Item -ItemType Directory -Force -Path $profile | Out-Null

    $args = @(
        "--user-data-dir=`"$profile`"",
        "--load-extension=`"$ext`"",
        "--disable-extensions-except=`"$ext`"",
        "--no-first-run",
        "--no-default-browser-check",
        "--new-window",
        $Url
    )

    Write-Host 'Starting Microsoft Edge with the PMRS Excel extension...' -ForegroundColor Green
    Write-Host "  extension : $ext"
    Write-Host "  profile   : $profile"
    Write-Host '  Log in once in this window - the session is kept for next time.'
    Start-Process -FilePath $edge -ArgumentList $args | Out-Null
}
else {
    # Chrome 137+ ignores --load-extension, so the extension has to be installed
    # once with "Load unpacked" in chrome://extensions. After that it is always on.
    $chromeArgs = @('--new-window', $Url)

    if ($chrome) {
        Write-Host 'Opening Google Chrome...' -ForegroundColor Green
        Write-Host ''
        Write-Host 'NOTE: Chrome 137+ no longer accepts --load-extension.'
        Write-Host '      Install the extension once (chrome://extensions -> Developer mode ->'
        Write-Host '      Load unpacked -> select the "extension" folder). After that it is'
        Write-Host '      loaded automatically every time you open PMRS.' -ForegroundColor Yellow
        Start-Process -FilePath $chrome -ArgumentList $chromeArgs | Out-Null
    }
    else {
        Write-Host "Chrome not found - opening the report in your default browser." -ForegroundColor Yellow
        Start-Process $Url | Out-Null
    }
}
