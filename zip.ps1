# Build ZIP archive with 7-Zip
7z a "Report-Generator.zip" `
"app" `
"data" `
"tests" `
"run.py" `
"run.spec" `
"installer.iss" `
"requirements.txt" `
"icon.ico" `
"-xr!__pycache__" `
"-xr!*.pyc"

Write-Host "Archive created: Report-Generator.zip"
