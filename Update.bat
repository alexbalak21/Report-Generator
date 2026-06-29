@echo off
setlocal

REM === Create temporary PS1 file ===
set "TempPS1=%TEMP%\Update-ReportGenerator.ps1"

REM === Write the PowerShell updater script into the temp file ===
> "%TempPS1%" echo # ================================
>>"%TempPS1%" echo # Auto Update Script for Report-Generator
>>"%TempPS1%" echo # ================================
>>"%TempPS1%" echo
>>"%TempPS1%" echo # -------- CONFIGURATION VARIABLES --------
>>"%TempPS1%" echo
>>"%TempPS1%" echo $ZipUrl       = "https://github.com/alexbalak21/Report-Generator/archive/refs/heads/main.zip"
>>"%TempPS1%" echo $InstallPath  = "D:\Report-Generator"
>>"%TempPS1%" echo $ZipPath      = "$env:TEMP\ReportGeneratorUpdate.zip"
>>"%TempPS1%" echo $ExtractPath  = "$env:TEMP\ReportGeneratorExtract"
>>"%TempPS1%" echo $ZipRootFolder = "Report-Generator-main"
>>"%TempPS1%" echo
>>"%TempPS1%" echo # -------- START UPDATE PROCESS --------
>>"%TempPS1%" echo
>>"%TempPS1%" echo Write-Host "Downloading latest version from $ZipUrl ..."
>>"%TempPS1%" echo Invoke-WebRequest -Uri $ZipUrl -OutFile $ZipPath -UseBasicParsing
>>"%TempPS1%" echo
>>"%TempPS1%" echo Write-Host "Extracting ZIP..."
>>"%TempPS1%" echo if (Test-Path $ExtractPath) { Remove-Item $ExtractPath -Recurse -Force }
>>"%TempPS1%" echo Expand-Archive -LiteralPath $ZipPath -DestinationPath $ExtractPath
>>"%TempPS1%" echo
>>"%TempPS1%" echo $SourcePath = Join-Path $ExtractPath $ZipRootFolder
>>"%TempPS1%" echo
>>"%TempPS1%" echo Write-Host "Copying files to $InstallPath ..."
>>"%TempPS1%" echo Copy-Item -Path "$SourcePath\*" -Destination $InstallPath -Recurse -Force
>>"%TempPS1%" echo
>>"%TempPS1%" echo Write-Host "Cleaning temporary files..."
>>"%TempPS1%" echo Remove-Item $ZipPath -Force
>>"%TempPS1%" echo Remove-Item $ExtractPath -Recurse -Force
>>"%TempPS1%" echo
>>"%TempPS1%" echo Write-Host "Update completed successfully!"

REM === Run the temporary PS1 script ===
powershell -ExecutionPolicy Bypass -File "%TempPS1%"

REM === Cleanup ===
del "%TempPS1%"

pause
