# ================================
# Auto Release Script (Interactive Version)
# ================================

$initFile       = "app/__init__.py"
$installerFile  = "installer.iss"

# --- 1. Read current version from __init__.py ---
Write-Host "Reading current version..."
$initContent = Get-Content $initFile
$versionLine = $initContent | Select-String "__version__"
$currentVersion = ($versionLine -split '"')[1]

Write-Host "Current version detected: $currentVersion"

# --- 2. Ask user for next version ---
$nextVersion = Read-Host "Enter the next version (example: $currentVersion → 1.0.7)"

if (-not $nextVersion) {
    Write-Host "No version entered. Aborting."
    exit
}

Write-Host "Next version will be: $nextVersion"

# --- 3. Update __init__.py ---
Write-Host "Updating __init__.py..."
$updatedInit = $initContent -replace $currentVersion, $nextVersion
Set-Content $initFile $updatedInit

# --- 4. Update installer.iss ---
Write-Host "Updating installer.iss..."
$installerContent = Get-Content $installerFile

$installerContent = $installerContent `
    -replace "AppVersion=$currentVersion", "AppVersion=$nextVersion" `
    -replace "ReportGenerator-Setup-$currentVersion", "ReportGenerator-Setup-$nextVersion"

Set-Content $installerFile $installerContent

# --- 5. Print build commands ---
Write-Host ""
Write-Host "==============================="
Write-Host "BUILD COMMANDS"
Write-Host "==============================="
Write-Host "pyinstaller run.spec"
Write-Host "& 'C:\Program Files (x86)\Inno Setup 6\ISCC.exe' installer.iss"
Write-Host ""

# --- 6. Print git commands ---
Write-Host "==============================="
Write-Host "GIT COMMANDS"
Write-Host "==============================="
Write-Host "git add ."
Write-Host "git commit -m \"Release v$nextVersion\""
Write-Host "git tag v$nextVersion"
Write-Host "git push origin main"
Write-Host "git push origin v$nextVersion"
Write-Host ""

# --- 7. Release info ---
Write-Host "==============================="
Write-Host "GITHUB RELEASE INFO"
Write-Host "==============================="
Write-Host "Upload file: installer_output\\ReportGenerator-Setup-$nextVersion.exe"
Write-Host "Release tag: v$nextVersion"
Write-Host "Release title: v$nextVersion"
Write-Host "URL: https://github.com/alexbalak21/Report-Generator/releases/new"
Write-Host ""

Write-Host "Done. Version updated everywhere."
