$ErrorActionPreference = 'Stop'

$ProjectPath = $PSScriptRoot
$PythonPath = Join-Path $ProjectPath '.venv\Scripts\python.exe'

Push-Location $ProjectPath
try {
    if (-not (Test-Path -LiteralPath $PythonPath)) {
        $PythonLauncher = Get-Command py -ErrorAction SilentlyContinue
        if ($PythonLauncher) {
            & $PythonLauncher.Source -3 -m venv (Join-Path $ProjectPath '.venv')
        } elseif (Get-Command python -ErrorAction SilentlyContinue) {
            & python -m venv (Join-Path $ProjectPath '.venv')
        } else {
            throw 'Python launcher "py" was not found. Install Python 3.10 or later, then run this script again.'
        }
        if ($LASTEXITCODE -ne 0) { throw 'Could not create the project virtual environment.' }
    }

    & $PythonPath -c 'import streamlit, pypdf, docx, dotenv' 2>$null
    if ($LASTEXITCODE -ne 0) {
        & $PythonPath -m pip install -r (Join-Path $ProjectPath 'requirements.txt')
        if ($LASTEXITCODE -ne 0) { throw 'Dependency installation failed. Check your internet connection and try again.' }
    }

    & $PythonPath -m streamlit run (Join-Path $ProjectPath 'app.py')
}
finally {
    Pop-Location
}
