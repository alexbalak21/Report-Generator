# ================================
# Auto Update Script for Report-Generator
# ================================

# -------- CONFIGURATION VARIABLES --------

# URL of the ZIP to download
$ZipUrl       = "https://github.com/alexbalak21/Report-Generator/archive/refs/heads/main.zip"

# Local folder where the app is installed
$InstallPath  = "D:\Report-Generator"

# Temporary paths
$ZipPath      = "$env:TEMP\ReportGeneratorUpdate.zip"
$ExtractPath  = "$env:TEMP\ReportGeneratorExtract"

# Name of the folder inside the ZIP
$ZipRootFolder = "Report-Generator-main"

# -------- START UPDATE PROCESS --------

Write-Host "Downloading latest version from $ZipUrl ..."
Invoke-WebRequest -Uri $ZipUrl -OutFile $ZipPath -UseBasicParsing

Write-Host "Extracting ZIP..."
if (Test-Path $ExtractPath) { Remove-Item $ExtractPath -Recurse -Force }
Expand-Archive -LiteralPath $ZipPath -DestinationPath $ExtractPath

# Build full path to extracted folder
$SourcePath = Join-Path $ExtractPath $ZipRootFolder

Write-Host "Copying files to $InstallPath ..."
Copy-Item -Path "$SourcePath\*" -Destination $InstallPath -Recurse -Force

Write-Host "Cleaning temporary files..."
Remove-Item $ZipPath -Force
Remove-Item $ExtractPath -Recurse -Force

Write-Host "Update completed successfully!"
