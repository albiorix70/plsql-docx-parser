# PowerShell wrapper to invoke MCP/CLI (SQLcl-like) to run run_compile.sql
# Expects env vars: MCP_CLI (optional), DB_USER, DB_PASSWORD, DB_HOST, DB_PORT, DB_SERVICE

$LogFile = Join-Path -Path (Split-Path -Parent $PSScriptRoot) -ChildPath "mcp_compile_output.log"
New-Item -Path $LogFile -ItemType File -Force | Out-Null
function Write-Log { param($m) $m | Tee-Object -FilePath $LogFile -Append }

if (-not $env:DB_USER -or -not $env:DB_PASSWORD -or -not $env:DB_HOST -or -not $env:DB_PORT -or -not $env:DB_SERVICE) {
    Write-Log "Missing DB connection environment variables. Set DB_USER, DB_PASSWORD, DB_HOST, DB_PORT, DB_SERVICE."
    exit 1
}

$mcp = $env:MCP_CLI
if (-not $mcp) {
    $candidates = @('mcp','sdcli','sql')
    foreach ($c in $candidates) {
        $p = Get-Command $c -ErrorAction SilentlyContinue
        if ($p) { $mcp = $p.Path; break }
    }
}

if (-not $mcp) {
    Write-Log "MCP/CLI binary not found. Set MCP_CLI environment variable or install 'mcp'/'sdcli'/'sql' on PATH."
    exit 1
}

$connect = "${env:DB_USER}/${env:DB_PASSWORD}@//${env:DB_HOST}:${env:DB_PORT}/${env:DB_SERVICE}"
Write-Log "Using MCP/CLI: $mcp"
Write-Log "Running: $mcp -L $connect @run_compile.sql"

try {
    $mcpArgs = @('-L', $connect, '@run_compile.sql')
    & $mcp @mcpArgs 2>&1 | Tee-Object -FilePath $LogFile -Append
    $ec = $LASTEXITCODE
    if ($ec -ne 0) { Write-Log "MCP/CLI exited with code $ec"; exit $ec }
} catch {
    Write-Log "Exception running MCP/CLI: $_"
    exit 1
}

Write-Log "MCP/CLI compile finished; see $LogFile"
