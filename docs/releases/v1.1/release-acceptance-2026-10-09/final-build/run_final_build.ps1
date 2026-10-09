param(
    [string]$Repo = 'E:/codex/excel-tools',
    [string]$EvidenceRoot = $PSScriptRoot,
    [string]$BuildPython = 'C:/Users/win/AppData/Local/Temp/excel-tools-v1.1-acceptance-23bc299521804363abc3c99f5a99ed02/venv/Scripts/python.exe'
)
$ErrorActionPreference = 'Stop'
$expectedSource = @{
    'excel_unmerge_fill.py' = '877485D3C497DF2CDF83FE1A7729002EF0B91C1E964109CC2F684FC068227F75'
    'excel_unmerge_gui.py' = '0D59D53E067DA326984702ECE5939347448A829012B564363EB9469C42B25F87'
    'packaging/使用说明.txt' = 'B40115CBA8374F727E63D1CDD04C93560B7433BDC21A87C0833AAACF8312D8DE'
}
foreach ($key in $expectedSource.Keys) {
    if ((Get-FileHash -LiteralPath (Join-Path $Repo $key) -Algorithm SHA256).Hash -ne $expectedSource[$key]) {
        throw ('Source differs from authorized freeze: ' + $key)
    }
}
$env:PYTHONDONTWRITEBYTECODE = '1'
$env:PYTHONUTF8 = '1'
$expectedPreviousHash = '92870BDC52451341C7268F6E1937A3741ECD3E1012DE0CF2E4026C606CCDD4A9'
$package = Join-Path $Repo 'dist/ExcelTools-v1.1-Windows-x64.zip'
$archive = Join-Path $Repo 'dist/archive/ExcelTools-v1.1-pre-deep-review-92870bdc.zip'
if ((Get-FileHash -LiteralPath $package -Algorithm SHA256).Hash -ne $expectedPreviousHash) {
    throw 'Existing ZIP differs from the authorized pre-deep-review candidate.'
}
foreach ($name in @('unit-tests.txt', 'build.txt', 'verification-manifest.json')) {
    if (Test-Path -LiteralPath (Join-Path $EvidenceRoot $name)) {
        throw 'Evidence already exists. Choose a fresh EvidenceRoot instead of overwriting it.'
    }
}
$versions = & $BuildPython -c 'import json,sys,PyInstaller,openpyxl;print(json.dumps(dict(python=sys.version.split()[0],pyinstaller=PyInstaller.__version__,openpyxl=openpyxl.__version__)))'
if ($LASTEXITCODE -ne 0) { throw 'Existing build environment is unavailable.' }
$versions = $versions | ConvertFrom-Json
if ($versions.python -ne '3.12.8' -or $versions.pyinstaller -ne '6.22.3' -or $versions.openpyxl -ne '3.1.5') {
    throw 'Build dependency versions differ from the previously validated environment.'
}
$workRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('excel-tools-v1.1-final-build-' + [guid]::NewGuid().ToString('N'))
$snapshot = Join-Path $workRoot 'candidate-source'
$testsFolder = Join-Path $snapshot 'tests'
$auditFolder = Join-Path $workRoot 'independent-acceptance'
New-Item -ItemType Directory -Path $snapshot, $testsFolder, $auditFolder | Out-Null
Copy-Item -LiteralPath (Join-Path $Repo 'excel_unmerge_fill.py'), (Join-Path $Repo 'excel_unmerge_gui.py') -Destination $snapshot
Get-ChildItem -LiteralPath (Join-Path $Repo 'tests') -Filter '*.py' | Copy-Item -Destination $testsFolder
Copy-Item -LiteralPath (Join-Path $Repo 'packaging/使用说明.txt') -Destination (Join-Path $snapshot '使用说明.txt')
Copy-Item -LiteralPath (Join-Path $Repo 'docs/releases/v1.1/evidence/independent_acceptance.py'), (Join-Path $Repo 'docs/releases/v1.1/evidence/baseline_engine.py') -Destination $auditFolder
$version = & $BuildPython -c 'import sys;sys.path.insert(0,sys.argv[1]);from excel_unmerge_gui import APP_VERSION;print(APP_VERSION)' $snapshot
if ($LASTEXITCODE -ne 0 -or $version -ne '1.1') { throw 'Candidate application version must be exactly 1.1.' }
$sourceHashes = Get-ChildItem -LiteralPath $snapshot -Recurse -File | Get-FileHash -Algorithm SHA256 | Select-Object Path, Hash
$auditHashes = Get-ChildItem -LiteralPath $auditFolder -File | Get-FileHash -Algorithm SHA256 | Select-Object Path, Hash
$sourceHashes | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $EvidenceRoot 'source-snapshot-hashes.json') -Encoding utf8
$head = git -C $Repo rev-parse HEAD
$gitStatus = git -C $Repo status --short
$gitStatus | Set-Content -LiteralPath (Join-Path $EvidenceRoot 'source-status.txt') -Encoding utf8
$env:PYTHONDONTWRITEBYTECODE = '1'
$env:PYTHONUTF8 = '1'
Push-Location $snapshot
try {
    & $BuildPython -m unittest discover -s tests -v 2>&1 | Tee-Object -FilePath (Join-Path $EvidenceRoot 'unit-tests.txt')
    if ($LASTEXITCODE -ne 0) { throw 'Frozen source test suite failed.' }
    if (-not (Select-String -LiteralPath (Join-Path $EvidenceRoot 'unit-tests.txt') -Pattern '^Ran 49 tests in ' -Quiet)) { throw 'Expected 49 tests were not executed.' }
    & $BuildPython (Join-Path $auditFolder 'independent_acceptance.py') $snapshot 2>&1 | Tee-Object -FilePath (Join-Path $EvidenceRoot 'independent-acceptance.txt')
    if ($LASTEXITCODE -ne 0) { throw 'Independent acceptance failed.' }
    if (-not (Select-String -LiteralPath (Join-Path $EvidenceRoot 'independent-acceptance.txt') -Pattern '^Ran 9 tests in ' -Quiet)) { throw 'Expected 9 independent tests were not executed.' }
    & $BuildPython excel_unmerge_gui.py --self-test --self-test-log (Join-Path $EvidenceRoot 'source-self-test.txt')
    if ($LASTEXITCODE -ne 0) { throw 'Source self-test failed.' }
    & $BuildPython -m PyInstaller --noconfirm --clean --onefile --windowed --name ExcelTools --distpath (Join-Path $workRoot 'built') --workpath (Join-Path $workRoot 'build') --specpath $snapshot excel_unmerge_gui.py 2>&1 | Tee-Object -FilePath (Join-Path $EvidenceRoot 'build.txt')
    if ($LASTEXITCODE -ne 0) { throw 'Portable build failed.' }
} finally {
    Pop-Location
}
$portableDir = Join-Path $workRoot '免安装 验收'
$workingDir = Join-Path $workRoot '独立 工作目录'
$scratchDir = Join-Path $workRoot '临时 解压目录'
New-Item -ItemType Directory -Path $portableDir, $workingDir, $scratchDir | Out-Null
$portableExe = Join-Path $portableDir 'ExcelTools.exe'
Copy-Item -LiteralPath (Join-Path $workRoot 'built/ExcelTools.exe') -Destination $portableExe
$selfTestLog = Join-Path $EvidenceRoot 'portable-self-test.txt'
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
        throw 'Portable synthetic self-test exceeded 45 seconds.'
    }
    $app.Refresh()
    $portableExit = $app.ExitCode
    "PORTABLE_SELF_TEST_EXIT=$portableExit" | Tee-Object -FilePath (Join-Path $EvidenceRoot 'portable-process.txt')
    if ($portableExit -ne 0) { throw 'Portable isolated self-test failed.' }
} finally {
    foreach ($key in @('PATH', 'TEMP', 'TMP')) { Remove-Item "Env:$key" -ErrorAction SilentlyContinue }
    foreach ($key in $saved.Keys) { Set-Item "Env:$key" $saved[$key] }
}
foreach ($item in $sourceHashes) {
    if ((Get-FileHash -LiteralPath $item.Path -Algorithm SHA256).Hash -ne $item.Hash) {
        throw 'Frozen source changed during build.'
    }
}
foreach ($item in $sourceHashes) {
    $relative = $item.Path.Substring($snapshot.Length + 1)
    if ($relative -eq '使用说明.txt') { $relative = 'packaging/使用说明.txt' }
    if ((Get-FileHash -LiteralPath (Join-Path $Repo $relative) -Algorithm SHA256).Hash -ne $item.Hash) {
        throw ('Live source/test/instructions changed after freeze: ' + $relative)
    }
}
$instructions = Get-Content -Raw -Encoding utf8 -LiteralPath (Join-Path $snapshot '使用说明.txt')
[System.IO.File]::WriteAllText((Join-Path $portableDir '使用说明.txt'), $instructions, [System.Text.UTF8Encoding]::new($true))
$newPackage = Join-Path $workRoot 'ExcelTools-v1.1-Windows-x64.zip'
Compress-Archive -LiteralPath $portableExe, (Join-Path $portableDir '使用说明.txt') -DestinationPath $newPackage
# Preserve the exact previous candidate before replacing the current download.
New-Item -ItemType Directory -Force -Path (Split-Path -Parent $archive) | Out-Null
if (-not (Test-Path -LiteralPath $archive)) { Copy-Item -LiteralPath $package -Destination $archive }
if ((Get-FileHash -LiteralPath $archive -Algorithm SHA256).Hash -ne $expectedPreviousHash) {
    throw 'Previous candidate archive does not match the original bytes.'
}
if ((Get-FileHash -LiteralPath $package -Algorithm SHA256).Hash -ne $expectedPreviousHash) {
    throw 'Current candidate was changed by another process before replacement.'
}
Copy-Item -LiteralPath $newPackage -Destination $package -Force
if ((Get-FileHash -LiteralPath $package -Algorithm SHA256).Hash -ne (Get-FileHash -LiteralPath $newPackage -Algorithm SHA256).Hash) {
    throw 'Published local candidate differs from the validated new ZIP.'
}
$manifest = [ordered]@{
    source_head = $head
    expected_frozen_source_hashes = $expectedSource
    unit_test_count = 49
    independent_test_count = 9
    source_status = $gitStatus
    source_snapshot = $snapshot
    source_hashes = $sourceHashes
    independent_acceptance_hashes = $auditHashes
    build_script = (Get-FileHash -LiteralPath $PSCommandPath -Algorithm SHA256 | Select-Object Path, Hash)
    app_version = $version
    build_python = $BuildPython
    dependencies = $versions
    portable_self_test_exit = $portableExit
    executable = (Get-FileHash -Algorithm SHA256 -LiteralPath $portableExe | Select-Object Path, Hash)
    package = (Get-FileHash -Algorithm SHA256 -LiteralPath $package | Select-Object Path, Hash)
    previous_candidate_archive = (Get-FileHash -Algorithm SHA256 -LiteralPath $archive | Select-Object Path, Hash)
    captured_utc = [DateTime]::UtcNow.ToString('o')
    isolation = 'EXE only; Chinese/space paths; separate cwd; stripped Python/Tcl environment; restricted PATH; separate extraction directory; hidden launch'
    limitations = 'Local Windows 11; no clean-machine/remote CI/official Release claim; build checks use synthetic fixtures only; business-client retesting is recorded separately.'
}
$manifest | ConvertTo-Json -Depth 10 | Tee-Object -FilePath (Join-Path $EvidenceRoot 'verification-manifest.json')
