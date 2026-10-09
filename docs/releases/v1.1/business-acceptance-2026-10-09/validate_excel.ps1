param([ValidateSet('engine', 'exe')][string]$Track = 'engine', [string]$EvidenceName = '', [string]$Revision = '')
$ErrorActionPreference = 'Stop'
$repo = (Resolve-Path (Join-Path $PSScriptRoot '../../../..')).Path
$work = Join-Path $repo ('testfile/验收_v1.1_2026-10-09/' + $Revision + '/' + $Track)
$inventory = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'input-inventory.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$outPath = Join-Path $PSScriptRoot $(if ($EvidenceName) { $EvidenceName } else { 'excel-client-' + $Track + '.json' })
if (Test-Path -LiteralPath $outPath) { throw 'Preserve previous evidence; choose a new evidence file before rerunning.' }
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type @'
using System;
using System.Runtime.InteropServices;
public static class ExcelAcceptanceWindow {
 [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
}
'@
function Release-Com($value) {
    if ($null -ne $value -and [Runtime.InteropServices.Marshal]::IsComObject($value)) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($value)
    }
}
function Cell-Point([string]$address) {
    if ($address -notmatch '^([A-Z]+)([0-9]+)$') { throw 'Unexpected coordinate' }
    $column = 0
    foreach ($char in $Matches[1].ToCharArray()) { $column = $column * 26 + [int]$char - 64 }
    return @{ row = [int]$Matches[2]; column = $column }
}
function Read-Book([string]$path, $sheetInfo) {
    $book = $null; $sheets = $null
    try {
        $missing = [Type]::Missing
        # Normal load (0), read-only, no link updates, explicit empty passwords.
        $book = $script:books.Open($path, 0, $true, 5, '', '', $true, 2, '', $false, $false, 0, $false, $true, 0)
        $sheets = $book.Worksheets
        $snapshots = @()
        for ($index = 1; $index -le $sheets.Count; $index++) {
            $sheet = $null; $range = $null; $used = $null
            try {
                $sheet = $sheets.Item($index)
                $shape = $sheetInfo[$index - 1].actual_cell_bounds
                $columnNumber = [int]$shape.max_column; $columnName = ''
                while ($columnNumber -gt 0) {
                    $columnNumber--
                    $columnName = [char](65 + ($columnNumber % 26)) + $columnName
                    $columnNumber = [int][Math]::Floor($columnNumber / 26)
                }
                $range = $sheet.Range('A1:' + $columnName + $shape.max_row)
                $values = $range.Value2
                $used = $sheet.UsedRange
                $snapshots += ,@{ values = $values; rows = [int]$shape.max_row; columns = [int]$shape.max_column; used_range = $used.Address(); sheet_index = $index }
            } finally { Release-Com $used; Release-Com $range; Release-Com $sheet }
        }
        return @{ sheets = $snapshots; opened_read_only = [bool]$book.ReadOnly; worksheet_count = $sheets.Count }
    } finally {
        if ($null -ne $book) { $book.Close($false) }
        Release-Com $sheets; Release-Com $book
    }
}
$preExisting = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
$excel = $null; $script:books = $null; $ownedPid = 0; $ownsInstance = $false
$results = @(); $failure = $null; $version = $null; $build = $null
$before = @{}
foreach ($file in $inventory.files) {
    $before[$file.source_path] = (Get-FileHash -LiteralPath $file.source_path -Algorithm SHA256).Hash
    foreach ($mode in @('default', 'all')) {
        $copy = Join-Path (Join-Path $work $mode) $file.file_name
        $output = Join-Path (Join-Path $work $mode) ([IO.Path]::GetFileNameWithoutExtension($file.file_name) + '_拆分填充.xlsx')
        $before[$copy] = (Get-FileHash -LiteralPath $copy -Algorithm SHA256).Hash
        $before[$output] = (Get-FileHash -LiteralPath $output -Algorithm SHA256).Hash
    }
}
try {
    $excel = New-Object -ComObject Excel.Application
    [uint32]$excelProcess = 0
    [void][ExcelAcceptanceWindow]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$excelProcess)
    $ownedPid = [int]$excelProcess
    if ($preExisting -contains $ownedPid -or $ownedPid -eq 0) { throw 'Could not establish a separate owned Excel instance.' }
    $ownsInstance = $true
    $excel.Visible = $false
    $excel.AutomationSecurity = 3
    $excel.EnableEvents = $false
    $excel.AskToUpdateLinks = $false
    $excel.DisplayAlerts = $false
    $version = $excel.Version; $build = $excel.Build
    $clientFile = Join-Path $excel.Path 'EXCEL.EXE'
    $clientFileVersion = (Get-Item -LiteralPath $clientFile).VersionInfo.FileVersion
    $script:books = $excel.Workbooks
    foreach ($file in $inventory.files) {
        $copy = Join-Path (Join-Path $work 'default') $file.file_name
        $baseline = Read-Book $copy $file.sheets
        $archive = [IO.Compression.ZipFile]::OpenRead($copy)
        $mergeSets = @()
        try {
            foreach ($sheet in $file.sheets) {
                $stream = $archive.GetEntry($sheet.part).Open()
                $reader = [IO.StreamReader]::new($stream)
                try { [xml]$xml = $reader.ReadToEnd() } finally { $reader.Dispose(); $stream.Dispose() }
                $refs = @($xml.SelectNodes('//*[local-name()="mergeCell"]') | ForEach-Object { $_.GetAttribute('ref') })
                $mergeSets += ,$refs
            }
        } finally { $archive.Dispose() }
        foreach ($mode in @('default', 'all')) {
            $output = Join-Path (Join-Path $work $mode) ([IO.Path]::GetFileNameWithoutExtension($file.file_name) + '_拆分填充.xlsx')
            $actual = Read-Book $output $file.sheets
            $issues = @(); $sheetResults = @(); $checked = 0
            if (-not $baseline.opened_read_only -or -not $actual.opened_read_only) { $issues += 'Not read-only' }
            if ($baseline.worksheet_count -ne $actual.worksheet_count) { $issues += 'Worksheet count mismatch' }
            for ($s = 0; $s -lt $baseline.sheets.Count; $s++) {
                $src = $baseline.sheets[$s]; $dst = $actual.sheets[$s]
                $expected = $src.values.Clone()
                foreach ($merge in $mergeSets[$s]) {
                    $ends = $merge.Split(':'); $start = Cell-Point $ends[0]; $end = Cell-Point $ends[-1]
                    if ($mode -eq 'default' -and $start.row -eq $end.row) { continue }
                    $lastColumn = if ($mode -eq 'all') { $end.column } else { $start.column }
                    $anchor = $src.values.GetValue($start.row, $start.column)
                    for ($r = $start.row; $r -le $end.row; $r++) {
                        for ($c = $start.column; $c -le $lastColumn; $c++) { $expected.SetValue($anchor, $r, $c) }
                    }
                }
                for ($r = 1; $r -le $src.rows; $r++) {
                    for ($c = 1; $c -le $src.columns; $c++) {
                        $a = $expected.GetValue($r, $c); $b = $dst.values.GetValue($r, $c)
                        $checked++
                        if (-not [object]::Equals($a, $b)) { $issues += ('Sheet {0} R{1}C{2}: Excel Value2 mismatch' -f ($s + 1), $r, $c) }
                    }
                }
                $sheetResults += @{ sheet_index = $s + 1; source_used_range = $src.used_range; output_used_range = $dst.used_range; grid_cells_checked = $src.rows * $src.columns }
            }
            $results += @{ file = $file.file_name; mode = $mode; output = $output; opened_normal_mode = $true; opened_read_only = $actual.opened_read_only; grid_cells_checked = $checked; sheet_results = $sheetResults; passed = $issues.Count -eq 0; issues = $issues }
            Write-Output ('Excel {0}: {1} PASS={2}, cells={3}' -f $mode, $file.file_name, ($issues.Count -eq 0), $checked)
        }
    }
} catch { $failure = $_.Exception.Message; $failureStack = $_.ScriptStackTrace }
finally {
    Release-Com $script:books
    if ($ownsInstance -and $null -ne $excel) { $excel.Quit() }
    Release-Com $excel
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
$hashResults = @()
foreach ($path in $before.Keys) {
    $after = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
    $hashResults += @{ path = $path; sha256_before = $before[$path].ToLowerInvariant(); sha256_after = $after.ToLowerInvariant(); unchanged = $before[$path] -eq $after }
}
if ($ownedPid -gt 0 -and $ownsInstance) {
    for ($attempt = 0; $attempt -lt 25 -and (Get-Process -Id $ownedPid -ErrorAction SilentlyContinue); $attempt++) { Start-Sleep -Milliseconds 200 }
}
$report = @{
    track = $Track; recorded_at = (Get-Date).ToString('o'); excel_version = $version; excel_build = $build
    revision = $Revision; validation_script_sha256 = (Get-FileHash -LiteralPath $PSCommandPath -Algorithm SHA256).Hash.ToLowerInvariant()
    candidate_package_sha256 = (Get-FileHash -LiteralPath (Join-Path $repo 'dist/ExcelTools-v1.1-Windows-x64.zip') -Algorithm SHA256).Hash.ToLowerInvariant()
    client = 'Microsoft Excel COM, local installed client'; owned_excel_pid = $ownedPid; preexisting_excel_pids = $preExisting
    excel_executable = $clientFile; excel_file_version = $clientFileVersion
    owned_excel_exited = -not [bool](Get-Process -Id $ownedPid -ErrorAction SilentlyContinue)
    normal_load_no_recovery = $true; update_links = 0; automation_security = 3; saves_performed = 0
    results = $results; file_hash_checks = $hashResults; failure = $failure; failure_stack = $failureStack
    passed = $null -eq $failure -and $results.Count -eq 10 -and @($results | Where-Object { -not $_.passed }).Count -eq 0 -and @($hashResults | Where-Object { -not $_.unchanged }).Count -eq 0
    scope_limit = 'Normal-mode read-only opening and Excel values, not visual layout or repair-dialog observation; DisplayAlerts disabled. No formula/macros/charts samples present.'
}
$report | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $outPath -Encoding UTF8
Write-Output ('Report={0}; PASS={1}; Failure={2}' -f $outPath, $report.passed, $failure)
if (-not $report.passed) { exit 1 }
