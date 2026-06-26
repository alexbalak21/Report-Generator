# Delete all __pycache__ folders recursively from the current directory

Write-Host "Searching for __pycache__ folders..." -ForegroundColor Cyan

$folders = Get-ChildItem -Path . -Recurse -Directory -Filter "__pycache__"

foreach ($folder in $folders) {
    Write-Host "Deleting: $($folder.FullName)" -ForegroundColor Yellow
    Remove-Item -Recurse -Force $folder.FullName
}

Write-Host "All __pycache__ folders removed." -ForegroundColor Green
