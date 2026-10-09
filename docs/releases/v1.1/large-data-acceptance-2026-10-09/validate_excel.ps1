param(
    [Parameter(Mandatory=$true)][string]$Source,
    [Parameter(Mandatory=$true)][string]$DefaultOutput,
    [Parameter(Mandatory=$true)][string]$AllOutput,
    [Parameter(Mandatory=$true)][string]$Report
)
$ErrorActionPreference = 'Stop'
if (Test-Path -LiteralPath $Report) { throw 'Preserve previous report; select a new path.' }
$inventory = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'input-inventory.json') -Raw -Encoding UTF8 | ConvertFrom-Json
if ((Get-FileHash -LiteralPath $Source -Algorithm SHA256).Hash.ToLowerInvariant() -ne $inventory.source_sha256_before) { throw 'Source does not match independently inventoried daily workbook.' }
Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Xml;
public static class LargeExcelCheck {
 [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid);
 public static int[] Point(string address) {
   address=address.Replace("$", "").ToUpperInvariant(); int i=0, col=0;
   while (i<address.Length && address[i]>='A' && address[i]<='Z') { col=col*26+address[i]-'A'+1; i++; }
   return new int[]{int.Parse(address.Substring(i)),col};
 }
 public static string[] Merges(string path, string part) {
   var result=new List<string>();
   using(var zip=ZipFile.OpenRead(path)) using(var stream=zip.GetEntry(part).Open())
   using(var reader=XmlReader.Create(stream, new XmlReaderSettings { DtdProcessing=DtdProcessing.Prohibit, XmlResolver=null })) {
     while(reader.Read()) if(reader.NodeType==XmlNodeType.Element && reader.LocalName=="mergeCell") result.Add(reader.GetAttribute("ref"));
   }
   return result.ToArray();
 }
 public static string[] Compare(Array source, Array actual, string[] merges, bool all, int rows, int columns) {
   var issues=new List<string>(); var expected=(Array)source.Clone();
   foreach(var merge in merges) {
     var ends=merge.Split(':'); var a=Point(ends[0]); var b=Point(ends[ends.Length-1]);
     if (!all && a[0]==b[0]) continue;
     var anchor=source.GetValue(a[0],a[1]);
     for(int r=a[0];r<=b[0];r++) for(int c=a[1];c<=(all?b[1]:a[1]);c++) expected.SetValue(anchor,r,c);
   }
   for(int r=1;r<=rows;r++) for(int c=1;c<=columns;c++) {
     if (!object.Equals(expected.GetValue(r,c), actual.GetValue(r,c)) && issues.Count<50) issues.Add("R"+r+"C"+c+": Value2 mismatch");
   }
   return issues.ToArray();
 }
}
'@
function Release-Com($value) {
    if ($null -ne $value -and [Runtime.InteropServices.Marshal]::IsComObject($value)) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($value) }
}
function Read-Book([string]$Path) {
    $book = $null; $sheets = $null; $snapshots = @()
    try {
        $book = $script:books.Open($Path, 0, $true, 5, '', '', $true, 2, '', $false, $false, 0, $false, $true, 0)
        if (-not $book.ReadOnly) { throw 'Expected read-only workbook.' }
        $sheets = $book.Worksheets
        if ($sheets.Count -ne $inventory.sheets.Count) { throw 'Sheet count differs from inventory.' }
        foreach ($info in $inventory.sheets) {
            $sheet = $null; $range = $null; $used = $null
            try {
                $sheet = $sheets.Item([int]$info.sheet_index)
                $range = $sheet.Range($info.actual_bounds)
                $values = $range.Value2
                $used = $sheet.UsedRange
                $snapshots += ,@{ values=$values; rows=[int]$info.max_row; columns=[int]$info.max_column; used_range=$used.Address(); part=$info.part; index=[int]$info.sheet_index }
            } finally { Release-Com $used; Release-Com $range; Release-Com $sheet }
        }
        return @{ sheets=$snapshots; read_only=[bool]$book.ReadOnly }
    } finally {
        if ($null -ne $book) { $book.Close($false) }
        Release-Com $sheets; Release-Com $book
    }
}
$Source = (Resolve-Path -LiteralPath $Source).Path
$DefaultOutput = (Resolve-Path -LiteralPath $DefaultOutput).Path
$AllOutput = (Resolve-Path -LiteralPath $AllOutput).Path
$before=@{}
foreach ($path in @($Source,$DefaultOutput,$AllOutput)) { $before[$path]=(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash }
$preExisting=@(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
$excel=$null; $script:books=$null; $ownedPid=0; $ownsInstance=$false; $failure=$null; $results=@()
try {
    $excel=New-Object -ComObject Excel.Application
    [uint32]$instancePid=0
    [void][LargeExcelCheck]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$instancePid)
    $ownedPid=[int]$instancePid
    if ($ownedPid -eq 0 -or $preExisting -contains $ownedPid) { throw 'Could not establish separate Excel ownership.' }
    $ownsInstance=$true
    $excel.Visible=$false; $excel.AutomationSecurity=3; $excel.EnableEvents=$false; $excel.AskToUpdateLinks=$false; $excel.DisplayAlerts=$false
    $version=$excel.Version; $build=$excel.Build
    $clientFile=Join-Path $excel.Path 'EXCEL.EXE'
    $clientVersion=(Get-Item -LiteralPath $clientFile).VersionInfo.FileVersion
    $script:books=$excel.Workbooks
    $baseline=Read-Book $Source
    foreach ($mode in @('default','all')) {
        $output=if($mode -eq 'default'){$DefaultOutput}else{$AllOutput}
        $actual=Read-Book $output
        $sheetResults=@(); $checked=0; $issues=@()
        for($i=0;$i -lt $baseline.sheets.Count;$i++) {
            $a=$baseline.sheets[$i]; $b=$actual.sheets[$i]
            $merges=[LargeExcelCheck]::Merges($Source,$a.part)
            $errors=[LargeExcelCheck]::Compare($a.values,$b.values,$merges,($mode -eq 'all'),$a.rows,$a.columns)
            $issues+=@($errors | ForEach-Object { 'sheet '+$a.index+': '+$_ })
            $count=$a.rows*$a.columns; $checked+=$count
            $sheetResults+=@{ sheet_index=$a.index; checked_cells=$count; source_used_range=$a.used_range; output_used_range=$b.used_range }
        }
        $results+=@{ mode=$mode; output_sha256=$before[$output].ToLowerInvariant(); checked_cells=$checked; sheets=$sheetResults; issues=$issues; passed=$issues.Count -eq 0 }
        Write-Output ('Excel '+$mode+': checked='+$checked+'; issues='+$issues.Count)
        $actual=$null
    }
} catch { $failure=$_.Exception.Message }
finally {
    Release-Com $script:books
    if($ownsInstance -and $null -ne $excel){$excel.Quit()}
    Release-Com $excel
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
$hashes=@()
foreach($path in $before.Keys){
    $after=(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
    $hashes+=@{ file_name=[IO.Path]::GetFileName($path); sha256=$before[$path].ToLowerInvariant(); unchanged=$after -eq $before[$path] }
}
if($ownsInstance){for($i=0;$i -lt 25 -and (Get-Process -Id $ownedPid -ErrorAction SilentlyContinue);$i++){Start-Sleep -Milliseconds 200}}
$exited=$ownsInstance -and -not [bool](Get-Process -Id $ownedPid -ErrorAction SilentlyContinue)
$passed=$null -eq $failure -and $results.Count -eq 2 -and @($results | Where-Object {-not $_.passed}).Count -eq 0 -and @($hashes | Where-Object {-not $_.unchanged}).Count -eq 0 -and $exited
@{ recorded_utc=[DateTime]::UtcNow.ToString('o'); script_sha256=(Get-FileHash -LiteralPath $PSCommandPath -Algorithm SHA256).Hash.ToLowerInvariant(); excel_version=$version; excel_build=$build; excel_file_version=$clientVersion; source_sha256=$before[$Source].ToLowerInvariant(); results=$results; file_hash_checks=$hashes; owned_pid=$ownedPid; owned_excel_exited=$exited; preexisting_excel_pids=$preExisting; normal_read_only_load=$true; macro_security=3; update_links=0; saves=0; failure=$failure; passed=$passed; limitations='Local Excel Value2 comparison with DisplayAlerts disabled; not a visual or repair-dialog observation, not WPS or macro execution.' } | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $Report -Encoding UTF8
if(-not $passed){throw ('Excel validation failed; see '+$Report)}
