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

# --- 5. Execute build commands ---
Write-Host "==============================="
Write-Host "BUILD PROCESS"
Write-Host "==============================="

Write-Host "Running PyInstaller..."
pyinstaller run.spec

Write-Host "Running Inno Setup Compiler..."
& 'C:\Program Files (x86)\Inno Setup 6\ISCC.exe' installer.iss

Write-Host "Build completed."

# --- 6. Execute git commands ---
Write-Host "==============================="
Write-Host "GIT PROCESS"
Write-Host "==============================="

Write-Host "Adding files..."
git add .

Write-Host "Committing..."
git commit -m "Release v$nextVersion"

Write-Host "Tagging..."
git tag v$nextVersion

Write-Host "Pushing main..."
git push origin main

Write-Host "Pushing tag..."
git push origin v$nextVersion

Write-Host "Git operations completed."

# --- 7. Open browser to GitHub release page ---
Write-Host "Opening GitHub release page..."
Start-Process "https://github.com/alexbalak21/Report-Generator/releases/new"

# --- 8. Final info ---
Write-Host "==============================="
Write-Host "GITHUB RELEASE INFO"
Write-Host "==============================="
Write-Host "Upload file: installer_output\\ReportGenerator-Setup-$nextVersion.exe"
Write-Host "Release tag: v$nextVersion"
Write-Host "Release title: v$nextVersion"
Write-Host "URL: https://github.com/alexbalak21/Report-Generator/releases/new"
Write-Host ""

Write-Host "Done. Version updated, installer built, git pushed, browser opened."
