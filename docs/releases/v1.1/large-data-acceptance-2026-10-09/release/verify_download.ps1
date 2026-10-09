param([Parameter(Mandatory=$true)][string]$Package)
$ErrorActionPreference = 'Stop'
$Package = (Resolve-Path -LiteralPath $Package).Path
$expectedZip = '07b801517060896c01ea9be2aaf1420784d2e87fe3a5f2ae4ea1b31f5b0488f6'
$expectedExe = '239b1456771ae417f044dfd1ac2642f7ac36de7bc3781c6d2ac3ddc75d2d7ea4'
if ((Get-FileHash -LiteralPath $Package -Algorithm SHA256).Hash.ToLowerInvariant() -ne $expectedZip) { throw 'Downloaded package identity differs from accepted candidate.' }
$workRoot = Join-Path $env:TEMP ('excel-tools-release-download-' + [guid]::NewGuid().ToString('N'))
$portableDir = Join-Path $workRoot '下载 解压'
$workingDir = Join-Path $workRoot '独立 工作目录'
$scratchDir = Join-Path $workRoot '临时 解压目录'
New-Item -ItemType Directory -Path $portableDir, $workingDir, $scratchDir | Out-Null
Expand-Archive -LiteralPath $Package -DestinationPath $portableDir
$portableExe = Join-Path $portableDir 'ExcelTools.exe'
if ((Get-FileHash -LiteralPath $portableExe -Algorithm SHA256).Hash.ToLowerInvariant() -ne $expectedExe) { throw 'Downloaded executable identity differs from accepted candidate.' }
$contents = @(Get-ChildItem -LiteralPath $portableDir -File | Select-Object -ExpandProperty Name)
if ($contents.Count -ne 2 -or $contents -notcontains 'ExcelTools.exe' -or $contents -notcontains '使用说明.txt') { throw 'Unexpected package contents.' }
$selfTestLog = Join-Path $PSScriptRoot 'download-self-test.txt'
$report = Join-Path $PSScriptRoot 'download-self-test.json'
if ((Test-Path -LiteralPath $report) -or (Test-Path -LiteralPath $selfTestLog)) { throw 'Preserve previous verification evidence.' }
$saved = @{}
Get-ChildItem Env: | Where-Object {
    $_.Name -match '^(PYTHON.*|TCL.*|TK.*|VIRTUAL_ENV|CONDA.*|PATH|TEMP|TMP)$'
} | ForEach-Object { $saved[$_.Name] = $_.Value }
$started = [DateTime]::UtcNow
try {
    foreach ($key in $saved.Keys) { Remove-Item "Env:$key" -ErrorAction SilentlyContinue }
    $env:PATH = "$env:SystemRoot\System32;$env:SystemRoot"
    $env:TEMP = $scratchDir
    $env:TMP = $scratchDir
    $app = Start-Process -FilePath $portableExe -ArgumentList @('--self-test', '--self-test-log', ('"{0}"' -f $selfTestLog)) -WorkingDirectory $workingDir -WindowStyle Hidden -PassThru
    $processHandle = $app.Handle
    if (-not $app.WaitForExit(45000)) {
        $app.Kill($true)
        throw 'Downloaded executable self-test timed out.'
    }
    $app.Refresh()
    $exitCode = $app.ExitCode
} finally {
    foreach ($key in @('PATH', 'TEMP', 'TMP')) { Remove-Item "Env:$key" -ErrorAction SilentlyContinue }
    foreach ($key in $saved.Keys) { Set-Item "Env:$key" $saved[$key] }
}
$message = (Get-Content -LiteralPath $selfTestLog -Raw -Encoding UTF8).Trim()
$result = [ordered]@{
    started_utc = $started.ToString('o')
    ended_utc = [DateTime]::UtcNow.ToString('o')
    downloaded_package = $Package
    package_sha256 = (Get-FileHash -LiteralPath $Package -Algorithm SHA256).Hash.ToLowerInvariant()
    executable_sha256 = (Get-FileHash -LiteralPath $portableExe -Algorithm SHA256).Hash.ToLowerInvariant()
    extracted_contents = $contents
    executable_path = $portableExe
    exit_code = $exitCode
    process_exited = $app.HasExited
    self_test_log = $message
    isolation = 'Downloaded Release asset; fresh extraction; Chinese/space paths; separate cwd; no Python/Tcl environment; restricted PATH; separate TEMP/TMP; hidden launch.'
    script_sha256 = (Get-FileHash -LiteralPath $PSCommandPath -Algorithm SHA256).Hash.ToLowerInvariant()
    passed = $exitCode -eq 0 -and $message -eq 'ExcelTools v1.1 self-test passed.'
}
$result | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $report -Encoding UTF8
$result | ConvertTo-Json -Depth 5
if (-not $result.passed) { throw 'Downloaded executable self-test failed.' }
