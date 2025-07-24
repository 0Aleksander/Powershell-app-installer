#Requires -Version 5.1

$asciiArt = @"
    __    ___    __ __ _______________   _____________  ____    ______
   / /   /   |  / //_// ____/_  __/   | / ___/_  __/\ \/ / /   / ____/
  / /   / /| | / ,<  / __/   / / / /| | \__ \ / /    \  / /   / __/   
 / /___/ ___ |/ /| |/ /___  / / / ___ |___/ // /     / / /___/ /___   
/_____/_/  |_/_/ |_/_____/ /_/ /_/  |_/____//_/     /_/_____/_____/   
"@

# --- Ensure latest App Installer (winget) is installed ---
$wingetVersion = (winget --version)
Write-Host "Current winget version: $wingetVersion" -ForegroundColor DarkGray

try {
    winget upgrade --id Microsoft.AppInstaller --accept-source-agreements --accept-package-agreements --silent
    Write-Host "Attempted to upgrade App Installer to latest version." -ForegroundColor DarkGray
} catch {
    Write-Host "Error upgrading App Installer: $_" -ForegroundColor Red
}
Start-Sleep -Seconds 1  # brief pause before continuing

function Show-Menu {
    param (
        [string]$Title = 'Installation Menu',
        [bool]$WingetAvailable = $false
    )
    Clear-Host
    Write-Host $asciiArt -ForegroundColor Magenta
    Write-Host "============ $Title ============" -ForegroundColor Cyan
    if ($WingetAvailable) {
        Write-Host "Winget is AVAILABLE (v$(winget --version))" -ForegroundColor Green
        Write-Host "1: Standard Installation (Winget)"
    } else {
        Write-Host "Winget is NOT AVAILABLE or version < 1.4" -ForegroundColor Yellow
    }
    Write-Host "2: Ninite Fallback Installation"
    Write-Host "Q: Quit"
}

function Test-Winget {
    try {
        $wingetCmd = Get-Command winget -ErrorAction Stop
        $versionOutput = winget --version | Out-String
        $versionString = $versionOutput.Trim()

        if (-not [string]::IsNullOrEmpty($versionString)) {
            $versionString = $versionString -replace '^v',''
            $versionParts = $versionString.Split('.')

            if ($versionParts.Count -ge 2) {
                $major = [int]$versionParts[0]
                $minor = [int]$versionParts[1]
                return ($major -gt 1) -or ($major -eq 1 -and $minor -ge 4)
            }
        }
        return $false
    } catch {
        return $false
    }
}

function Install-Apps {
    param (
        [string[]]$AppList,
        [hashtable]$CustomSources
    )

    foreach ($app in $AppList) {
        Write-Host "`nInstalling $app..." -ForegroundColor Cyan
        try {
            if ($CustomSources.ContainsKey($app)) {
                $source = $CustomSources[$app]
                winget install --id $app --source $source --silent --accept-package-agreements --accept-source-agreements
            } else {
                winget install --id $app --silent --accept-package-agreements --accept-source-agreements
            }

            if ($LASTEXITCODE -eq 0) {
                Write-Host "$app installed successfully." -ForegroundColor Green
            } else {
                Write-Host "$app installation failed (Exit code: $LASTEXITCODE)" -ForegroundColor Red
            }
        } catch {
            Write-Host "Error installing $app : $_" -ForegroundColor Red
        }
        Start-Sleep -Milliseconds 500
    }
}

function Invoke-NiniteFallback {
    Write-Host "`n=== Ninite Fallback Installation ===" -ForegroundColor Yellow
    
    $niniteUrl = "https://ninite.com/7zip-chrome-firefox-greenshot-notepadplusplus-teamviewer15-vlc/ninite.exe"
    $ninitePath = "$env:TEMP\NiniteInstaller.exe"
    
    try {
        Write-Host "Downloading Ninite installer..." -ForegroundColor Cyan
        Invoke-WebRequest -Uri $niniteUrl -OutFile $ninitePath -ErrorAction Stop
        
        Write-Host "Running Ninite installer..." -ForegroundColor Cyan
        $process = Start-Process -FilePath $ninitePath -Wait -PassThru
        
        if ($process.ExitCode -eq 0) {
            Write-Host "Ninite install completed successfully." -ForegroundColor Green
            Write-Host "Note: Lenovo Vantage is not supported by Ninite. Please download it from the Microsoft Store." -ForegroundColor Yellow
        } else {
            Write-Host "Ninite install failed with exit code $($process.ExitCode)." -ForegroundColor Red
        }
    } catch {
        Write-Host "Error during Ninite installation: $_" -ForegroundColor Red
    } finally {
        if (Test-Path $ninitePath) {
            Remove-Item $ninitePath -Force -ErrorAction SilentlyContinue
        }
    }
}

# App definitions
$apps = @(
    "Google.Chrome",
    "Mozilla.Firefox",
    "Microsoft.Teams",
    "VideoLAN.VLC",
    "Greenshot.Greenshot",
    "TeamViewer.TeamViewer",
    "7zip.7zip",
    "Notepad++.Notepad++",
    "9WZDNCRFJ4MV"  # Lenovo Vantage
)

$customSources = @{
    "9WZDNCRFJ4MV" = "msstore"
}

# === MAIN LOOP ===
Clear-Host
Write-Host $asciiArt -ForegroundColor Magenta
Write-Host "=== Application Installation Manager ===" -ForegroundColor Cyan

$wingetAvailable = Test-Winget

do {
    Show-Menu -WingetAvailable $wingetAvailable
    $selection = Read-Host "`nPlease make a selection"

    switch ($selection.ToLower()) {
        '1' {
            if ($wingetAvailable) {
                Install-Apps -AppList $apps -CustomSources $customSources
            } else {
                Write-Host "Winget not available. Please choose option 2 for Ninite." -ForegroundColor Red
                Start-Sleep -Seconds 2
            }
        }
        '2' {
            Invoke-NiniteFallback
        }
        'q' {
            break
        }
        default {
            Write-Host "Invalid selection, please try again." -ForegroundColor Red
            Start-Sleep -Seconds 1
        }
    }

    if ($selection -ne 'q') {
        Write-Host "`nPress any key to return to menu..."
        $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    }

} while ($selection -ne 'q')