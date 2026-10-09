param([Parameter(Mandatory=$true)][string]$FixtureRoot,[Parameter(Mandatory=$true)][string]$Report)
$ErrorActionPreference='Stop'
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class QualityStyleWin32 {
 [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hwnd,out uint pid);
}
'@
$preexisting=@(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object Id)
$excel=$null;$owned=$false;$ownedPid=0
$result=@{cases=@();preexisting_excel_pids=$preexisting;fixture_scope='synthetic only';saves=0}
try {
 $excel=New-Object -ComObject Excel.Application
 [uint32]$detected=0
 [void][QualityStyleWin32]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$detected)
 if($detected -eq 0 -or $preexisting -contains $detected){throw 'Cannot establish own Excel process; no workbook opened.'}
 $owned=$true;$ownedPid=$detected
 $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.AutomationSecurity=3;$excel.AskToUpdateLinks=$false
 $result.excel_version=$excel.Version;$result.excel_build=$excel.Build
 $cases=Get-Content -LiteralPath (Join-Path $FixtureRoot 'fixtures.json') -Raw | ConvertFrom-Json
 foreach($case in $cases){
  $entry=@{name=$case.name;target=$case.target}
  foreach($kind in @('source','output')){
   $path=$case.$kind;$book=$null;$sheet=$null;$anchor=$null;$target=$null
   try {
    $book=$excel.Workbooks.Open($path,0,$true)
    $sheet=$book.Worksheets.Item(1);$anchor=$sheet.Range('A1');$target=$sheet.Range($case.target)
    $entry[$kind]=@{anchor_number_format=$anchor.NumberFormat;anchor_text=$anchor.Text;anchor_value2=$anchor.Value2;target_number_format=$target.NumberFormat;target_text=$target.Text;target_value2=$target.Value2}
   } finally {
    if($null -ne $book){$book.Close($false)}
    foreach($object in @($target,$anchor,$sheet,$book)){if($null -ne $object){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($object)}}
   }
  }
  $result.cases+=,$entry
 }
} finally {
 if($null -ne $excel){if($owned){$excel.Quit()};[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}
 [GC]::Collect();[GC]::WaitForPendingFinalizers()
 $result.owned_pid=$ownedPid
 $result.owned_excel_exited=$owned -and -not (Get-Process -Id $ownedPid -ErrorAction SilentlyContinue)
 [IO.File]::WriteAllText($Report,($result | ConvertTo-Json -Depth 8),[Text.UTF8Encoding]::new($false))
}
$result | ConvertTo-Json -Depth 8
