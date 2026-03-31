#Requires -Version 5.1
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$ProjectDir = Split-Path -Parent $PSScriptRoot
$EnvFile = Join-Path $ProjectDir '.env'

# Load .env from project root
if (Test-Path $EnvFile) {
    Get-Content $EnvFile | ForEach-Object {
        if ($_ -match '^\s*([^#][^=]+)=(.*)$') {
            [System.Environment]::SetEnvironmentVariable($Matches[1].Trim(), $Matches[2].Trim(), 'Process')
        }
    }
}

$user   = $env:ORACLE_USER   ?? $(throw 'ORACLE_USER not set in .env')
$pass   = $env:ORACLE_PASSWORD ?? $(throw 'ORACLE_PASSWORD not set in .env')
$conn   = $env:ORACLE_CONNECTION_STRING ?? $(throw 'ORACLE_CONNECTION_STRING not set in .env')

Set-Location $ProjectDir

Write-Host "Compiling doc-parser packages via sqlplus..."
& sqlplus -s "${user}/${pass}@${conn}" '@compile_docs_parser.sql'
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
