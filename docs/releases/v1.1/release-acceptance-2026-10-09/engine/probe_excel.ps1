param([Parameter(Mandatory=$true)][string]$ManifestPath, [Parameter(Mandatory=$true)][string]$EvidencePath)
$ErrorActionPreference = 'Stop'
if (Test-Path -LiteralPath $EvidencePath) { throw 'Choose a new evidence path; previous evidence is retained.' }
$manifest = Get-Content -LiteralPath $ManifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
Add-Type @'
using System;
using System.Runtime.InteropServices;
public static class EngineProbeWindow {
 [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
}
'@
function Release-ProbeCom($value) {
    if ($null -ne $value -and [Runtime.InteropServices.Marshal]::IsComObject($value)) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($value)
    }
}
$preExisting = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
$excel = $null; $books = $null; $owned = $false; $results = @(); $hashes = @{}
foreach ($case in $manifest.cases) {
    foreach ($kind in @('source','before','after')) {
        $hashes[$case.$kind] = (Get-FileHash -LiteralPath $case.$kind -Algorithm SHA256).Hash
    }
}
try {
    $excel = New-Object -ComObject Excel.Application
    [uint32]$excelProcess = 0
    [void][EngineProbeWindow]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$excelProcess)
    if ($preExisting -contains $excelProcess -or $excelProcess -eq 0) { throw 'Separate Excel instance not established.' }
    $owned = $true
    $excel.Visible = $false; $excel.AutomationSecurity = 3; $excel.EnableEvents = $false
    $excel.AskToUpdateLinks = $false; $excel.DisplayAlerts = $false
    $books = $excel.Workbooks
    foreach ($case in $manifest.cases) {
        $snap = [ordered]@{ name=$case.name; files=@() }
        foreach ($kind in @('source','before','after')) {
            $book = $null; $sheets = $null; $sheet = $null
            try {
                $book = $books.Open($case.$kind,0,$true,5,'','',$true,2,'',$false,$false,0,$false,$true,0)
                $sheets = $book.Worksheets; $sheet = $sheets.Item(1)
                $cells = @()
                foreach ($ref in @('A1',$case.target)) {
                    $range = $null
                    try {
                        $range = $sheet.Range($ref)
                        $cells += [ordered]@{ address=$ref; value=$range.Value2; format=$range.NumberFormat; merge=[bool]$range.MergeCells }
                    } finally { Release-ProbeCom $range }
                }
                $snap.files += [ordered]@{ kind=$kind; cells=$cells; readOnly=[bool]$book.ReadOnly }
            } catch {
                $snap.files += [ordered]@{ kind=$kind; error=$_.Exception.Message }
                if ($kind -eq 'source') { throw }
            } finally {
                if ($null -ne $book) { $book.Close($false) }
                Release-ProbeCom $sheet; Release-ProbeCom $sheets; Release-ProbeCom $book
            }
        }
        $snap['before_target_matches_source_anchor'] = -not $snap.files[1].error -and [object]::Equals($snap.files[0].cells[0].value,$snap.files[1].cells[1].value)
        $snap['after_target_matches_source_anchor'] = -not $snap.files[2].error -and [object]::Equals($snap.files[0].cells[0].value,$snap.files[2].cells[1].value)
        $snap['after_anchor_unchanged'] = -not $snap.files[2].error -and [object]::Equals($snap.files[0].cells[0].value,$snap.files[2].cells[0].value)
        $results += $snap
    }
    $unchanged = $true
    foreach ($entry in $hashes.GetEnumerator()) {
        if ((Get-FileHash -LiteralPath $entry.Key -Algorithm SHA256).Hash -ne $entry.Value) { $unchanged = $false }
    }
    $allPassed = @($results | Where-Object { -not $_.after_target_matches_source_anchor -or -not $_.after_anchor_unchanged }).Count -eq 0
    $report = [ordered]@{
        synthetic_only=$true; baseline_revision=$manifest.baseline_revision
        baseline_engine_sha256=$manifest.baseline_engine_sha256; current_engine_sha256=$manifest.current_engine_sha256
        manifest_path=$ManifestPath; excel_version=$excel.Version; excel_build=$excel.Build; owned_pid=$excelProcess
        files_unchanged=$unchanged; files_checked=$hashes.Count; all_fixed_cases_passed=$allPassed; results=$results
        limitations='Read-only normal-load Value2 comparison; no assertion about repair dialogs or visual rendering.'
    }
    $report | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $EvidencePath -Encoding UTF8
    [ordered]@{ cases=$results.Count; baseline_failures=@($results | Where-Object {-not $_.before_target_matches_source_anchor}).Count; after_all_passed=$allPassed; unchanged=$unchanged; evidence=$EvidencePath } | ConvertTo-Json
    if (-not $allPassed -or -not $unchanged) { throw 'Synthetic Excel probe failed.' }
} finally {
    Release-ProbeCom $books
    if ($owned) { $excel.Quit() }
    Release-ProbeCom $excel
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
