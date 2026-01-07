# CI/local helper: run run_compile.sql using SQLcl (PowerShell)
# Requires these environment variables to be set:
#   DB_USER, DB_PASSWORD, DB_HOST, DB_PORT, DB_SERVICE
# Optionally set SQLCL_DIR to the SQLcl installation folder (containing bin\sql.bat)

$LogFile = Join-Path -Path (Split-Path -Parent $PSScriptRoot) -ChildPath "compile_output.log"
New-Item -Path $LogFile -ItemType File -Force | Out-Null

function Write-Log { param($m) $m | Tee-Object -FilePath $LogFile -Append }

if (-not $env:DB_USER -or -not $env:DB_PASSWORD -or -not $env:DB_HOST -or -not $env:DB_PORT -or -not $env:DB_SERVICE) {
    Write-Log "Missing DB connection environment variables. Set DB_USER, DB_PASSWORD, DB_HOST, DB_PORT, DB_SERVICE."
    exit 1
}

$sqlclDir = $env:SQLCL_DIR
if (-not $sqlclDir) {
    # common default locations
    $candidates = @(
        "$Env:ProgramFiles\sqlcl",
        "$Env:ProgramFiles\Oracle\sqlcl",
        "C:\oracle\sqlcl",
        "C:\sqlcl"
    )
    foreach ($c in $candidates) {
        if (Test-Path $c) { $sqlclDir = $c; break }
    }
}

if (-not $sqlclDir -or -not (Test-Path $sqlclDir)) {
    Write-Log "SQLCL_DIR not set and no default SQLcl installation found. Set SQLCL_DIR environment variable to your SQLcl folder."
    exit 1
}

$sqlExe = Join-Path $sqlclDir "bin\sql.bat"
if (-not (Test-Path $sqlExe)) {
    Write-Log "sql executable not found at $sqlExe"
    Get-ChildItem -Path $sqlclDir -Recurse -Depth 2 | Out-String | Write-Log
    exit 1
}

$connect = "${env:DB_USER}/${env:DB_PASSWORD}@//${env:DB_HOST}:${env:DB_PORT}/${env:DB_SERVICE}"
Write-Log "Running compile script with: $connect"

try {
    $args = @('-L', $connect, '@run_compile.sql')
    & $sqlExe @args 2>&1 | Tee-Object -FilePath $LogFile -Append
    $ec = $LASTEXITCODE
    if ($ec -ne 0) {
        Write-Log "sqlcl exited with code $ec"
        exit $ec
    }
} catch {
    Write-Log "Exception running sqlcl: $_"
    exit 1
}

Write-Log "Compile finished; see $LogFile"
