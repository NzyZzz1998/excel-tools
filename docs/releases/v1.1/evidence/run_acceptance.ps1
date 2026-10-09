param(
    [string]$Repo = 'E:/codex/excel-tools',
    [string]$EvidenceRoot = $PSScriptRoot
)
$ErrorActionPreference = 'Stop'
$reviewPython = Join-Path $EvidenceRoot 'venv/Scripts/python.exe'
$snapshot = Join-Path $EvidenceRoot 'candidate-source'
$testsFolder = Join-Path $snapshot 'tests'
if (Test-Path -LiteralPath $snapshot) { throw 'Candidate snapshot already exists. Use a fresh evidence root for a changed candidate.' }
New-Item -ItemType Directory -Path $snapshot, $testsFolder | Out-Null
Copy-Item -LiteralPath (Join-Path $Repo 'excel_unmerge_fill.py'), (Join-Path $Repo 'excel_unmerge_gui.py') -Destination $snapshot
Get-ChildItem -LiteralPath (Join-Path $Repo 'tests') -Filter '*.py' | Copy-Item -Destination $testsFolder
$version = & $reviewPython -c "import sys;sys.path.insert(0,r'$snapshot');from excel_unmerge_gui import APP_VERSION;print(APP_VERSION)"
if ($LASTEXITCODE -ne 0 -or $version -ne '1.1') { throw 'Candidate application version must be exactly 1.1.' }
$sourceHashes = Get-ChildItem -LiteralPath $snapshot -Recurse -Filter '*.py' | Get-FileHash -Algorithm SHA256 | Select-Object Path, Hash
$sourceHashes | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $EvidenceRoot 'source-snapshot-hashes.json') -Encoding utf8
$head = git -C $Repo rev-parse HEAD
$gitStatus = git -C $Repo status --short
$gitStatus | Set-Content -LiteralPath (Join-Path $EvidenceRoot 'source-status.txt') -Encoding utf8
$env:PYTHONDONTWRITEBYTECODE = '1'
$env:PYTHONUTF8 = '1'
Push-Location $snapshot
try {
    & $reviewPython -m unittest discover -s tests -v 2>&1 | Tee-Object -FilePath (Join-Path $EvidenceRoot 'unit-tests.log')
    if ($LASTEXITCODE -ne 0) { throw 'Source test suite failed.' }
    & $reviewPython (Join-Path $EvidenceRoot 'independent_acceptance.py') $snapshot 2>&1 | Tee-Object -FilePath (Join-Path $EvidenceRoot 'independent-acceptance.log')
    if ($LASTEXITCODE -ne 0) { throw 'Independent acceptance failed.' }
    & $reviewPython excel_unmerge_gui.py --self-test --self-test-log (Join-Path $EvidenceRoot 'source-self-test.log')
    if ($LASTEXITCODE -ne 0) { throw 'Source self-test failed.' }
    & $reviewPython -m PyInstaller --noconfirm --clean --onefile --windowed --name ExcelTools --distpath (Join-Path $EvidenceRoot 'built') --workpath (Join-Path $EvidenceRoot 'build') --specpath $snapshot excel_unmerge_gui.py 2>&1 | Tee-Object -FilePath (Join-Path $EvidenceRoot 'build.log')
    if ($LASTEXITCODE -ne 0) { throw 'Portable build failed.' }
} finally {
    Pop-Location
}
$portableDir = Join-Path $EvidenceRoot '免安装 验收'
$workingDir = Join-Path $EvidenceRoot '独立 工作目录'
$scratchDir = Join-Path $EvidenceRoot '临时 解压目录'
New-Item -ItemType Directory -Path $portableDir, $workingDir, $scratchDir | Out-Null
$portableExe = Join-Path $portableDir 'ExcelTools.exe'
Copy-Item -LiteralPath (Join-Path $EvidenceRoot 'built/ExcelTools.exe') -Destination $portableExe
$selfTestLog = Join-Path $EvidenceRoot 'portable-self-test.log'
$saved = @{}
Get-ChildItem Env: | Where-Object {
    $_.Name -match '^(PYTHON.*|TCL.*|TK.*|VIRTUAL_ENV|CONDA.*|PATH|TEMP|TMP)$'
} | ForEach-Object { $saved[$_.Name] = $_.Value }
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
    "PORTABLE_SELF_TEST_EXIT=$portableExit" | Tee-Object -FilePath (Join-Path $EvidenceRoot 'portable-process.log')
    if ($portableExit -ne 0) { throw 'Portable self-test failed.' }
} finally {
    foreach ($key in @('PATH', 'TEMP', 'TMP')) { Remove-Item "Env:$key" -ErrorAction SilentlyContinue }
    foreach ($key in $saved.Keys) { Set-Item "Env:$key" $saved[$key] }
}
$instructions = Get-Content -Raw -Encoding utf8 -LiteralPath (Join-Path $Repo 'packaging/使用说明.txt')
[System.IO.File]::WriteAllText((Join-Path $portableDir '使用说明.txt'), $instructions, [System.Text.UTF8Encoding]::new($true))
$destination = Join-Path $Repo 'dist'
New-Item -ItemType Directory -Force -Path $destination | Out-Null
$package = Join-Path $destination 'ExcelTools-v1.1-Windows-x64.zip'
if (Test-Path -LiteralPath $package) { throw 'Existing candidate ZIP must be reviewed before replacement.' }
Compress-Archive -LiteralPath $portableExe, (Join-Path $portableDir '使用说明.txt') -DestinationPath $package
$manifest = [ordered]@{
    source_head = $head
    source_status = $gitStatus
    source_snapshot = $snapshot
    source_hashes = $sourceHashes
    app_version = $version
    python = (& $reviewPython --version)
    pyinstaller = (& $reviewPython -m PyInstaller --version)
    openpyxl = (& $reviewPython -c 'import openpyxl; print(openpyxl.__version__)')
    portable_self_test_exit = $portableExit
    executable = (Get-FileHash -Algorithm SHA256 -LiteralPath $portableExe | Select-Object Path, Hash)
    package = (Get-FileHash -Algorithm SHA256 -LiteralPath $package | Select-Object Path, Hash)
    captured_utc = [DateTime]::UtcNow.ToString('o')
    isolation = 'EXE only, Chinese/space paths, separate cwd, stripped Python/Tcl environment, restricted PATH, separate extraction directory, hidden launch'
    limitations = 'Local Windows 11 validation; no remote CI or official Release run; no private workbook content inspected; source snapshot includes authorized uncommitted v1.1 implementation'
}
$manifest | ConvertTo-Json -Depth 8 | Tee-Object -FilePath (Join-Path $EvidenceRoot 'verification-manifest.json')
