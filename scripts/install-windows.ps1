# ============================================================================
#  Hermes Agent - Windows Installer
#
#  One-liner -- paste this into any PowerShell terminal:
#    irm https://raw.githubusercontent.com/claudlos/hermes-windows-installer/main/scripts/install-windows.ps1 | iex
#
#  If you hit a TLS/SSL error, run this line first, then the one-liner above:
#    [Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }
#
#  Or run locally:
#    .\scripts\install-windows.ps1
#
#  Optional custom Python:
#    .\scripts\install-windows.ps1 -PythonExe C:\path\to\python.exe
# ============================================================================

param(
    [string]$PythonExe = "",
    [ValidateSet("none", "nous", "staff")]
    [string]$DesktopIcon = "staff"
)

# -- TLS fix ----------------------------------------------------------------
try {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
} catch {}
try {
    [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }
} catch {}

$ErrorActionPreference = "Stop"
if ($PSVersionTable.PSVersion.Major -ge 7) { $ProgressPreference = "Bypass" } else { $ProgressPreference = "SilentlyContinue" }

$INSTALL_DIR = "$env:LOCALAPPDATA\hermes-agent"
$VENV_DIR = "$INSTALL_DIR\.venv"
$BIN_DIR = "$env:LOCALAPPDATA\Programs\hermes"
$ICONS_DIR = "$INSTALL_DIR\icons"

$SCRIPT_PATH = $MyInvocation.MyCommand.Path
$SCRIPT_DIR = if ($SCRIPT_PATH) { Split-Path -Parent $SCRIPT_PATH } elseif ($PSScriptRoot) { $PSScriptRoot } else { $pwd.Path }
$REPO_ROOT = Split-Path -Parent $SCRIPT_DIR
$FROM_REPO = (Test-Path "$REPO_ROOT\.git") -and (Test-Path "$REPO_ROOT\pyproject.toml")

if ($FROM_REPO) {
    $SOURCE_DIR = $REPO_ROOT
    $BRANCH = (& git -C $REPO_ROOT rev-parse --abbrev-ref HEAD 2>$null)
    if (-not $BRANCH) { $BRANCH = "windows-qol-v2" }
    $REPO_URL = "https://github.com/claudlos/hermes-agent.git"
} else {
    $SOURCE_DIR = $null
    $REPO_URL = "https://github.com/claudlos/hermes-agent.git"
    $BRANCH = "windows-qol-v2"
}

function Write-Step($n, $msg) { Write-Host "`n  [$n] " -NoNewline -ForegroundColor DarkYellow; Write-Host $msg }
function Write-Ok($msg)      { Write-Host "      $msg" -ForegroundColor Green }
function Write-Dim($msg)     { Write-Host "      $msg" -ForegroundColor DarkGray }
function Write-Err($msg)     { Write-Host "      $msg" -ForegroundColor Red }
function Fail($msg)          { Write-Err $msg; exit 1 }
function Invoke-GitQuiet {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$Arguments
    )
    $prevEAP = $ErrorActionPreference
    try {
        $ErrorActionPreference = 'Continue'
        $null = & git @Arguments 2>&1
        return $LASTEXITCODE
    } finally {
        $ErrorActionPreference = $prevEAP
    }
}

Write-Host ""
Write-Host "  ============================================" -ForegroundColor DarkYellow
Write-Host "   Hermes Agent - Windows Installer" -ForegroundColor Yellow
Write-Host "  ============================================" -ForegroundColor DarkYellow
if ($FROM_REPO) {
    Write-Host "  Mode:   Local build (current branch)" -ForegroundColor DarkGray
    Write-Host "  Branch: $BRANCH" -ForegroundColor DarkGray
    Write-Host "  Source: $SOURCE_DIR" -ForegroundColor DarkGray
} else {
    Write-Host "  Mode:   Fresh install from GitHub" -ForegroundColor DarkGray
}

Write-Step 1 "Checking Python..."

$python = $null
if ($PythonExe) {
    if (-not (Test-Path $PythonExe)) {
        Write-Err "Requested Python not found: $PythonExe"
        exit 1
    }
    try {
        $ver = & $PythonExe --version 2>&1
        if ($ver -match "Python 3\.(\d+)") {
            $minor = [int]$Matches[1]
            if ($minor -ge 10) {
                $python = $PythonExe
                Write-Ok "Using explicit Python: $ver"
            }
        }
    } catch {}
    if (-not $python) {
        Write-Err "Explicit Python must be 3.10+: $PythonExe"
        exit 1
    }
} else {
    foreach ($cmd in @("python3", "python", "py")) {
        try {
            $ver = & $cmd --version 2>&1
            if ($ver -match "Python 3\.(\d+)") {
                $minor = [int]$Matches[1]
                if ($minor -ge 10) {
                    $python = $cmd
                    Write-Ok "Found $ver"
                    break
                }
            }
        } catch {}
    }
}

if (-not $python) {
    Write-Err "Python 3.10+ required but not found."
    Write-Host ""
    Write-Host "  Download from https://www.python.org/downloads/" -ForegroundColor Yellow
    Write-Host "  Check 'Add Python to PATH' during install." -ForegroundColor DarkGray
    Write-Host "  Or pass a custom interpreter: .\scripts\install-windows.ps1 -PythonExe C:\path\to\python.exe" -ForegroundColor DarkGray
    Write-Host ""
    exit 1
}

if (-not $FROM_REPO) {
    Write-Step 2 "Checking Git..."
    try {
        $gitVer = & git --version 2>&1
        Write-Ok $gitVer
    } catch {
        Write-Err "Git not found. Install from https://git-scm.com/download/win"
        exit 1
    }
} else {
    Write-Step 2 "Using local repo"
    Write-Ok $SOURCE_DIR
}

Write-Step 3 "Preparing source..."

if ($FROM_REPO) {
    if ($SOURCE_DIR -ne $INSTALL_DIR) {
        $clonedOk = $false
        if (Test-Path "$INSTALL_DIR\.git") {
            Write-Dim "Updating existing clone at $INSTALL_DIR..."
            Push-Location $INSTALL_DIR
            try {
                $fetchExit = Invoke-GitQuiet @('fetch', 'origin', $BRANCH, '--depth', '1')
                $checkoutExit = Invoke-GitQuiet @('checkout', $BRANCH)
                $resetExit = Invoke-GitQuiet @('reset', '--hard', "origin/$BRANCH")
                $clonedOk = ($fetchExit -eq 0 -and $checkoutExit -eq 0 -and $resetExit -eq 0)
            } finally {
                Pop-Location
            }
        } else {
            Write-Dim "Shallow-cloning $REPO_URL ($BRANCH) to $INSTALL_DIR..."
            if (Test-Path $INSTALL_DIR) {
                Remove-Item $INSTALL_DIR -Recurse -Force
            }
            $cloneExit = Invoke-GitQuiet @('clone', '--depth', '1', '--branch', $BRANCH, $REPO_URL, $INSTALL_DIR)
            $clonedOk = ($cloneExit -eq 0)
        }
        if ($clonedOk) {
            Write-Ok "Source ready via shallow clone ($BRANCH)"
        } else {
            Write-Dim "Shallow clone failed; falling back to robocopy sync..."
            if (-not (Test-Path $INSTALL_DIR)) {
                New-Item -ItemType Directory -Path $INSTALL_DIR -Force | Out-Null
            }
            & robocopy $SOURCE_DIR $INSTALL_DIR /MIR /XD .git .venv __pycache__ .pytest_cache node_modules PCbuild externals /XF "*.pyc" /NFL /NDL /NJH /NJS /NP 2>&1 | Out-Null
            $robocopyCode = $LASTEXITCODE
            if ($robocopyCode -ge 8) {
                Fail "robocopy sync failed with exit code $robocopyCode"
            }
            Push-Location $INSTALL_DIR
            try {
                if (-not (Test-Path ".git")) {
                    $initExit = Invoke-GitQuiet @('init', '--quiet')
                    if ($initExit -eq 0) {
                        $addExit = Invoke-GitQuiet @('add', '-A')
                        $commitExit = Invoke-GitQuiet @('commit', '-m', "install snapshot from $BRANCH", '--quiet')
                    }
                }
            } finally {
                Pop-Location
            }
            Write-Ok "Source synced via robocopy ($BRANCH)"
        }
    } else {
        Write-Ok "Installing from current directory"
    }
} else {
    if (Test-Path "$INSTALL_DIR\.git") {
        Write-Dim "Updating existing clone..."
        Push-Location $INSTALL_DIR
        try {
            $fetchExit = Invoke-GitQuiet @('fetch', 'origin', $BRANCH)
            if ($fetchExit -ne 0) { Fail "git fetch failed for branch $BRANCH" }
            $checkoutExit = Invoke-GitQuiet @('checkout', $BRANCH)
            if ($checkoutExit -ne 0) { Fail "git checkout failed for branch $BRANCH" }
            $pullExit = Invoke-GitQuiet @('pull', 'origin', $BRANCH)
            if ($pullExit -ne 0) { Fail "git pull failed for branch $BRANCH" }
        } finally {
            Pop-Location
        }
        Write-Ok "Updated to latest $BRANCH"
    } else {
        if (Test-Path $INSTALL_DIR) {
            Remove-Item $INSTALL_DIR -Recurse -Force
        }
        Write-Dim "Cloning..."
        $cloneExit = Invoke-GitQuiet @('clone', '--depth', '1', '--branch', $BRANCH, $REPO_URL, $INSTALL_DIR)
        if ($cloneExit -ne 0) { Write-Err "Clone failed"; exit 1 }
        Write-Ok "Cloned $BRANCH"
    }
}

Write-Step 4 "Setting up Python environment..."

if (-not (Test-Path "$VENV_DIR\Scripts\python.exe")) {
    Write-Dim "Creating virtual environment..."
    & $python -m venv $VENV_DIR
    if ($LASTEXITCODE -ne 0) { Write-Err "Failed to create venv"; exit 1 }
}

$venvPython = "$VENV_DIR\Scripts\python.exe"
$venvPip = "$VENV_DIR\Scripts\pip.exe"

& $venvPython -m pip install --upgrade pip --quiet 2>&1 | Out-Null
Write-Ok "Virtual environment ready"

Write-Step 5 "Installing Hermes Agent..."
Write-Dim "This may take a minute on first install..."
$installTarget = "$INSTALL_DIR[keyring,pty]"

& $venvPip install -e $installTarget --quiet 2>&1
if ($LASTEXITCODE -ne 0) {
    Write-Dim "Retrying with full output..."
    & $venvPip install -e $installTarget
    if ($LASTEXITCODE -ne 0) { Write-Err "Install failed"; exit 1 }
}

$hermesExe = "$VENV_DIR\Scripts\hermes.exe"
if (-not (Test-Path $hermesExe)) { Write-Err "hermes.exe not found after install"; exit 1 }
Write-Ok "Hermes Agent installed"

Write-Step 6 "Creating launcher..."

if (-not (Test-Path $BIN_DIR)) {
    New-Item -ItemType Directory -Path $BIN_DIR -Force | Out-Null
}

$launcherLines = @(
    '@echo off',
    ('"{0}" %*' -f $hermesExe)
)
Set-Content "$BIN_DIR\hermes.bat" -Value $launcherLines -Encoding ASCII
Write-Ok "hermes.bat -> $BIN_DIR"

Write-Step 7 "Configuring PATH..."

$userPath = [Environment]::GetEnvironmentVariable("Path", "User")
if ($userPath -notlike "*$BIN_DIR*") {
    [Environment]::SetEnvironmentVariable("Path", "$BIN_DIR;$userPath", "User")
    $env:Path = "$BIN_DIR;$env:Path"
    Write-Ok "Added to PATH"
    Write-Dim "Restart your terminal for PATH to take effect"
} else {
    Write-Ok "Already in PATH"
}

if ($DesktopIcon -ne 'none') {
    Write-Step 8 "Creating desktop shortcut..."

    $iconFile = if ($DesktopIcon -eq 'staff') { 'hermes-staff.ico' } else { 'hermes-nous.ico' }

    $iconPath = $null
    if ($SCRIPT_PATH) {
        $localIcon = Join-Path (Split-Path -Parent $SCRIPT_PATH) "icons\$iconFile"
        if (Test-Path $localIcon) { $iconPath = $localIcon }
    }
    if (-not $iconPath) {
        $localIcon = "$INSTALL_DIR\scripts\icons\$iconFile"
        if (Test-Path $localIcon) { $iconPath = $localIcon }
    }
    if (-not $iconPath) {
        $localIcon = "$ICONS_DIR\$iconFile"
        if (Test-Path $localIcon) { $iconPath = $localIcon }
    }

    if (-not $iconPath) {
        Write-Dim "Downloading desktop icon..."
        if (-not (Test-Path $ICONS_DIR)) {
            New-Item -ItemType Directory -Path $ICONS_DIR -Force | Out-Null
        }
        $iconUrl = "https://raw.githubusercontent.com/claudlos/hermes-windows-installer/main/scripts/icons/$iconFile"
        $iconDest = "$ICONS_DIR\$iconFile"
        try {
            $wc = New-Object System.Net.WebClient
            $wc.Headers.Add("User-Agent", "hermes-windows-installer")
            $wc.DownloadFile($iconUrl, $iconDest)
            $iconPath = $iconDest
            Write-Ok "Icon downloaded ($iconFile)"
        } catch {
            Write-Dim "Icon download failed - skipping shortcut"
        }
    }

    if ($iconPath -and (Test-Path $iconPath)) {
        $desktopPath = [Environment]::GetFolderPath('Desktop')
        $shortcutPath = Join-Path $desktopPath 'Hermes Agent.lnk'
        $ws = New-Object -ComObject WScript.Shell
        $sc = $ws.CreateShortcut($shortcutPath)
        $sc.TargetPath = 'cmd.exe'
        $sc.Arguments = '/k hermes'
        $sc.WorkingDirectory = '%USERPROFILE%'
        $sc.IconLocation = $iconPath
        $sc.Description = 'Hermes Agent'
        $sc.Save()
        Write-Ok "Desktop shortcut created ($DesktopIcon)"
    } else {
        Write-Dim "Icon not found - skipping shortcut"
    }
}

Write-Step 9 "Verifying install..."

$verifyOut = & $hermesExe --version 2>&1
if ($verifyOut) { Write-Ok $verifyOut } else { Write-Ok "Binary runs" }

Write-Host ""
Write-Host "  ============================================" -ForegroundColor Green
Write-Host "   Hermes Agent installed successfully" -ForegroundColor Green
Write-Host "  ============================================" -ForegroundColor Green
Write-Host ""
Write-Host "  Next step - open a NEW terminal and run:" -ForegroundColor White
Write-Host ""
Write-Host "      hermes setup" -ForegroundColor Yellow
Write-Host ""
Write-Host "  This configures your API key and preferences." -ForegroundColor DarkGray
Write-Host ""
if ($FROM_REPO) {
    Write-Host ("  Built from: {0} ({1})" -f $SOURCE_DIR, $BRANCH) -ForegroundColor DarkGray
}
Write-Host ("  Installed:  {0}" -f $INSTALL_DIR) -ForegroundColor DarkGray
Write-Host ("  Launcher:   {0}\hermes.bat" -f $BIN_DIR) -ForegroundColor DarkGray
Write-Host ""
