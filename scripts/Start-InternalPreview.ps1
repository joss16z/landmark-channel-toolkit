# Run on the company-approved computer. No credentials or live API required.
$ErrorActionPreference = 'Stop'
$Repository = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
Set-Location $Repository
$Python = Join-Path $Repository '.venv/Scripts/python.exe'
if (-not (Test-Path $Python)) { throw 'Complete the internal deployment setup first.' }
$env:LANDMARK_PREVIEW_ONLY = '1'
& $Python -m streamlit run app.py --server.address=127.0.0.1 --server.port=8501 --server.headless=true
exit $LASTEXITCODE
