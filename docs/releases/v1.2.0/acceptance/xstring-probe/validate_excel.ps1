param([Parameter(Mandatory=$true)][string]$WorkDir)
$ErrorActionPreference = 'Stop'
$root = (Resolve-Path -LiteralPath $WorkDir).Path
$expected = Get-Content (Join-Path $root 'expected.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$beforePids = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
Add-Type @'
using System;
using System.Runtime.InteropServices;
public static class SyntheticExcelPid {
 [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint pid);
}
'@
$excel = $null
$book = $null
$owned = $false
$result = [ordered]@{ synthetic_only = $true; passed = $false; files = @(); cleanup_errors = @() }
try {
    $excel = New-Object -ComObject Excel.Application
    [uint32]$ownedPid = 0
    [void][SyntheticExcelPid]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$ownedPid)
    if ($ownedPid -eq 0 -or $beforePids -contains [int]$ownedPid) { throw 'Cannot verify an independent newly created Excel process.' }
    $owned = $true
    $result.independent_process_confirmed = $true
    $result.excel_version = [string]$excel.Version
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AskToUpdateLinks = $false
    $excel.AutomationSecurity = 3
    foreach ($name in @('synthetic-source.xlsx', 'synthetic-output.xlsx')) {
        $path = Join-Path $root $name
        $hash = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
        $book = $excel.Workbooks.Open($path, 0, $true)
        $sheet = $book.Worksheets.Item(1)
        $cells = @()
        for ($column = 1; $column -le 6; $column++) {
            $cell = $sheet.Cells.Item(2, $column)
            $value = [string]$cell.Value2
            $equal = [string]::Equals($value, [string]$expected[$column - 1], [StringComparison]::Ordinal)
            $hasFormula = [bool]$cell.HasFormula
            $cells += [ordered]@{ column = $column; value_equal_expected = $equal; utf16_length = $value.Length; has_formula = $hasFormula }
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($cell)
            if (!$equal -or $hasFormula) { throw "Synthetic cell mismatch, column $column" }
        }
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($sheet)
        $book.Close($false)
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)
        $book = $null
        $unchanged = $hash -eq (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
        if (!$unchanged) { throw 'Read-only synthetic source changed.' }
        $result.files += [ordered]@{ name = $name; cells = $cells; unchanged = $unchanged }
    }
    $result.passed = $true
} catch {
    $result.error_type = $_.Exception.GetType().FullName
    $result.error = $_.Exception.Message
} finally {
    if ($null -ne $book) {
        try { $book.Close($false) } catch { $result.cleanup_errors += $_.Exception.GetType().FullName }
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)
    }
    if ($null -ne $excel -and $owned) {
        try { $excel.Quit() } catch { $result.cleanup_errors += $_.Exception.GetType().FullName }
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    $result | ConvertTo-Json -Depth 9 | Set-Content (Join-Path $root 'excel-result.json') -Encoding UTF8
}
$result | ConvertTo-Json -Depth 9
if (!$result.passed -or $result.cleanup_errors.Count -ne 0) { exit 1 }
