param([string]$EvidenceRoot = $PSScriptRoot)
$ErrorActionPreference = 'Stop'
$audit = Get-Content -Raw -Encoding utf8 -LiteralPath (Join-Path $EvidenceRoot 'package-audit.json') | ConvertFrom-Json
if ($audit.status -ne 'passed') { throw 'Package audit must pass before execution.' }
$portableDir = $audit.package.temporary_extraction
$portableExe = Join-Path $portableDir 'ExcelTools.exe'
if ((Get-FileHash -LiteralPath $portableExe -Algorithm SHA256).Hash.ToLowerInvariant() -ne $audit.executable.sha256) {
    throw 'Extracted executable identity changed.'
}
$workRoot = Split-Path -Parent $portableDir
$workingDir = Join-Path $workRoot '独立 工作目录'
$scratchDir = Join-Path $workRoot '临时 解压目录'
New-Item -ItemType Directory -Path $workingDir, $scratchDir | Out-Null
$selfTestLog = Join-Path $EvidenceRoot 'fresh-portable-self-test.txt'
$saved = @{}
Get-ChildItem Env: | Where-Object {
    $_.Name -match '^(PYTHON.*|TCL.*|TK.*|VIRTUAL_ENV|CONDA.*|PATH|TEMP|TMP)$'
} | ForEach-Object { $saved[$_.Name] = $_.Value }
$startTime = [DateTime]::UtcNow
$signature = Get-AuthenticodeSignature -LiteralPath $portableExe
try {
    foreach ($key in $saved.Keys) { Remove-Item "Env:$key" -ErrorAction SilentlyContinue }
    $env:PATH = "$env:SystemRoot\System32;$env:SystemRoot"
    $env:TEMP = $scratchDir
    $env:TMP = $scratchDir
    $app = Start-Process -FilePath $portableExe -ArgumentList @('--self-test', '--self-test-log', ('"{0}"' -f $selfTestLog)) -WorkingDirectory $workingDir -WindowStyle Hidden -PassThru
    $processHandle = $app.Handle
    if (-not $app.WaitForExit(45000)) {
        $app.Kill($true)
        throw 'Portable self-test exceeded 45 seconds.'
    }
    $app.Refresh()
    $portableExit = $app.ExitCode
} finally {
    foreach ($key in @('PATH', 'TEMP', 'TMP')) { Remove-Item "Env:$key" -ErrorAction SilentlyContinue }
    foreach ($key in $saved.Keys) { Set-Item "Env:$key" $saved[$key] }
}
$message = (Get-Content -Raw -Encoding utf8 -LiteralPath $selfTestLog).Trim()
$record = [ordered]@{
    started_utc = $startTime.ToString('o')
    ended_utc = [DateTime]::UtcNow.ToString('o')
    executable_path = $portableExe
    executable_sha256 = (Get-FileHash -LiteralPath $portableExe -Algorithm SHA256).Hash.ToLowerInvariant()
    package_sha256_after = (Get-FileHash -LiteralPath $audit.package.path -Algorithm SHA256).Hash.ToLowerInvariant()
    package_unchanged = ((Get-FileHash -LiteralPath $audit.package.path -Algorithm SHA256).Hash.ToLowerInvariant() -eq $audit.package.sha256)
    process_id = $app.Id
    process_exited = $app.HasExited
    exit_code = $portableExit
    self_test_log = $message
    authenticode_status = $signature.Status.ToString()
    isolation = 'Fresh ZIP extraction; EXE+instructions only; Chinese and space paths; independent cwd; Python/Tcl variables removed; restricted PATH; distinct TEMP/TMP; hidden launch.'
    script_sha256 = (Get-FileHash -LiteralPath $PSCommandPath -Algorithm SHA256).Hash.ToLowerInvariant()
    status = $(if ($portableExit -eq 0 -and $message -eq 'ExcelTools v1.1 self-test passed.') { 'passed' } else { 'failed' })
}
$record | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $EvidenceRoot 'fresh-portable-process.json') -Encoding utf8
$record | ConvertTo-Json -Depth 5
if ($record.status -ne 'passed' -or -not $record.package_unchanged) { throw 'Fresh portable self-test or integrity check failed.' }
