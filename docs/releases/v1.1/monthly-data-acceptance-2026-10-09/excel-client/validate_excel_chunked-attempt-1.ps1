[CmdletBinding(DefaultParameterSetName='Validate')]
param(
    [Parameter(Mandatory=$true,ParameterSetName='Validate')][string]$Source,
    [Parameter(Mandatory=$true,ParameterSetName='Validate')][string]$DefaultOutput,
    [Parameter(Mandatory=$true,ParameterSetName='Validate')][string]$Report,
    [Parameter(ParameterSetName='Validate')][string]$Inventory=(Join-Path $PSScriptRoot 'input-inventory.json'),
    [Parameter(ParameterSetName='Validate')][string]$Progress,
    [Parameter(ParameterSetName='Validate')][string]$RunId,
    [Parameter(ParameterSetName='Validate')][ValidateRange(1,5000)][int]$ChunkRows=5000,
    [Parameter(Mandatory=$true,ParameterSetName='SelfTest')][switch]$SelfTest
)
$ErrorActionPreference='Stop'
Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Xml;
public sealed class MonthlyMerge {
 public int R1,C1,R2,C2; public object Anchor; public bool Captured;
 public MonthlyMerge(string reference) {
  var ends=reference.Split(':'); if(ends.Length>2) throw new ArgumentException("Invalid merge reference");
  var a=MonthlyExcelCheck.Point(ends[0]); var b=MonthlyExcelCheck.Point(ends[ends.Length-1]);
  R1=a[0];C1=a[1];R2=b[0];C2=b[1];
  if(R2<R1 || C2<C1) throw new ArgumentException("Reversed merge reference");
 }
}
public sealed class MonthlyBlockResult {
 public long Checked,Mismatches; public string[] IssueAddresses;
}
public sealed class MonthlyComparer {
 readonly Dictionary<int,List<MonthlyMerge>> columns=new Dictionary<int,List<MonthlyMerge>>();
 readonly Dictionary<int,int> cursors=new Dictionary<int,int>();
 readonly List<MonthlyMerge> starts=new List<MonthlyMerge>(); int nextStart=0, nextRow=1;
 public MonthlyComparer(MonthlyMerge[] merges) {
  foreach(var p in merges) if(p.R2>p.R1) {
   if(!columns.ContainsKey(p.C1)){columns[p.C1]=new List<MonthlyMerge>();cursors[p.C1]=0;}
   columns[p.C1].Add(p);starts.Add(p);
  }
  starts.Sort((a,b)=>a.R1.CompareTo(b.R1));
  foreach(var list in columns.Values) list.Sort((a,b)=>a.R1.CompareTo(b.R1));
 }
 static object At(object value,int row,int col,int rows,int cols) {
  var a=value as Array;
  if(a==null){if(rows!=1 || cols!=1)throw new ArgumentException("Expected rectangular Value2 array");return value;}
  if(a.Rank!=2 || a.GetLength(0)!=rows || a.GetLength(1)!=cols)throw new ArgumentException("Value2 dimensions differ");
  return a.GetValue(row+a.GetLowerBound(0),col+a.GetLowerBound(1));
 }
 static bool Equal(object a,object b) {
  var ea=a as ErrorWrapper;var eb=b as ErrorWrapper;
  if(ea!=null || eb!=null)return ea!=null && eb!=null && ea.ErrorCode==eb.ErrorCode;
  return object.Equals(a,b);
 }
 public MonthlyBlockResult Compare(object source,object actual,int firstRow,int rows,int cols) {
  if(firstRow!=nextRow || rows<1 || cols<1)throw new ArgumentException("Chunks must be contiguous from row one");
  int lastRow=firstRow+rows-1;
  while(nextStart<starts.Count && starts[nextStart].R1<=lastRow) {
   var p=starts[nextStart++];
   if(p.R1<firstRow || p.C1>cols)throw new ArgumentException("Merge anchor outside source chunk");
   p.Anchor=At(source,p.R1-firstRow,p.C1-1,rows,cols);p.Captured=true;
  }
  var result=new MonthlyBlockResult();var issues=new List<string>();
  for(int col=1;col<=cols;col++) {
   List<MonthlyMerge> list;columns.TryGetValue(col,out list);int cursor=list==null?0:cursors[col];
   for(int local=0;local<rows;local++) {
    int row=firstRow+local;object expected=At(source,local,col-1,rows,cols);
    if(list!=null) {
     while(cursor<list.Count && list[cursor].R2<row)cursor++;
     if(cursor<list.Count && list[cursor].R1<=row) {
      if(!list[cursor].Captured)throw new InvalidOperationException("Merge anchor was not captured");
      expected=list[cursor].Anchor;
     }
    }
    if(!Equal(expected,At(actual,local,col-1,rows,cols))) {
     result.Mismatches++;if(issues.Count<50)issues.Add("R"+row+"C"+col);
    }
    result.Checked++;
   }
   if(list!=null)cursors[col]=cursor;
  }
  nextRow=lastRow+1;result.IssueAddresses=issues.ToArray();return result;
 }
}
public static class MonthlyExcelCheck {
 [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd,out uint pid);
 public static int[] Point(string address) {
  address=address.Replace("$","").ToUpperInvariant();int i=0,col=0,row;
  while(i<address.Length && address[i]>='A' && address[i]<='Z'){col=col*26+address[i]-'A'+1;i++;}
  if(i==0 || !int.TryParse(address.Substring(i),out row) || row<1 || row>1048576 || col<1 || col>16384)throw new ArgumentException("Invalid coordinate");
  return new[]{row,col};
 }
 public static string Address(int row,int col) {
  string s="";while(col>0){col--;s=(char)('A'+col%26)+s;col/=26;}return s+row;
 }
 public static MonthlyMerge[] ReadMerges(string path,string part) {
  var result=new List<MonthlyMerge>();
  using(var zip=ZipFile.OpenRead(path))using(var stream=zip.GetEntry(part).Open())
  using(var reader=XmlReader.Create(stream,new XmlReaderSettings{DtdProcessing=DtdProcessing.Prohibit,XmlResolver=null})) {
   while(reader.Read())if(reader.NodeType==XmlNodeType.Element && reader.LocalName=="mergeCell" && reader.NamespaceURI=="http://schemas.openxmlformats.org/spreadsheetml/2006/main")result.Add(new MonthlyMerge(reader.GetAttribute("ref")));
  }
  return result.ToArray();
 }
 public static string SelfTest() {
  var checker=new MonthlyComparer(new[]{new MonthlyMerge("A1:A10002"),new MonthlyMerge("C1:D10002"),new MonthlyMerge("E2:F2")});
  long checkedCells=0;
  for(int first=1;first<=10002;first+=5000) {
   int rows=Math.Min(5000,10003-first);var a=new object[rows,6];var b=new object[rows,6];
   if(first==1){a[0,0]="  synthetic  ";a[0,2]=0.0;a[1,4]="retained";b[1,4]="retained";}
   for(int r=0;r<rows;r++){b[r,0]="  synthetic  ";b[r,2]=0.0;}
   a[rows-1,1]=" ";b[rows-1,1]=" ";
   var good=checker.Compare(a,b,first,rows,6);
   if(good.Mismatches!=0)throw new Exception("Cross-chunk or ordinary-cell regression");checkedCells+=good.Checked;
  }
  var negative=new MonthlyComparer(new[]{new MonthlyMerge("A1:A2")});
  var source=new object[,]{{"x",null},{null," "}};var actual=new object[,]{{"x",null},{"wrong",null}};
  var bad=negative.Compare(source,actual,1,2,2);
  if(bad.Mismatches!=2 || bad.IssueAddresses.Length!=2)throw new Exception("Mismatch detection regression");
  var scalar=new MonthlyComparer(new MonthlyMerge[0]);
  if(scalar.Compare(3.0,3.0,1,1,1).Mismatches!=0)throw new Exception("Scalar regression");
  return "PASS: cross-5000-row anchors, ordinary blanks/spaces, zero, retained horizontal merge, negative mismatches, scalar; checked="+checkedCells;
 }
}
'@
if($SelfTest){[MonthlyExcelCheck]::SelfTest();return}
if(Test-Path -LiteralPath $Report){throw 'Preserve previous report; select a new path.'}
if($Progress -and (Test-Path -LiteralPath $Progress)){throw 'Preserve previous progress; select a new path.'}
$producerCreation=(Get-Process -Id $PID).StartTime.ToUniversalTime().ToFileTimeUtc()
function Write-ProgressRecord([string]$Phase,[long]$Checked=0,[int]$SheetIndex=0){
    if(-not $Progress){return}
    $record=@{run_id=$RunId;producer_pid=$PID;producer_creation_filetime=$producerCreation;owned_excel_pid=$ownedPid;owned_excel_creation_filetime=$ownedCreation;ownership_confirmed=$ownsInstance;stage=$Phase;checked_cells_in_sheet=$Checked;sheet_index=$SheetIndex;recorded_utc=[DateTime]::UtcNow.ToString('o')}
    $temporaryProgress=$Progress+'.writing'
    [IO.File]::WriteAllText($temporaryProgress,($record | ConvertTo-Json -Compress),[Text.UTF8Encoding]::new($false))
    if(Test-Path -LiteralPath $Progress){[IO.File]::Replace($temporaryProgress,$Progress,$null)}else{[IO.File]::Move($temporaryProgress,$Progress)}
}
$Source=(Resolve-Path -LiteralPath $Source).Path
$DefaultOutput=(Resolve-Path -LiteralPath $DefaultOutput).Path
if($Source -eq $DefaultOutput){throw 'Source and result must be different paths.'}
$inventoryData=Get-Content -LiteralPath $Inventory -Raw -Encoding UTF8 | ConvertFrom-Json
$sourceBefore=(Get-FileHash -LiteralPath $Source -Algorithm SHA256).Hash.ToLowerInvariant()
if($sourceBefore -ne $inventoryData.source_sha256_before){throw 'Source differs from independently inventoried workbook.'}
$outputBefore=(Get-FileHash -LiteralPath $DefaultOutput -Algorithm SHA256).Hash.ToLowerInvariant()
$plans=@();$selected=0
foreach($info in $inventoryData.sheets){
    $merges=[MonthlyExcelCheck]::ReadMerges($Source,$info.part)
    $rows=[int]$info.max_row;$cols=[int]$info.max_column
    foreach($merge in $merges){$rows=[Math]::Max($rows,$merge.R2);$cols=[Math]::Max($cols,$merge.C2);if($merge.R2 -gt $merge.R1){$selected++}}
    $plans+=@{index=[int]$info.sheet_index;rows=$rows;cols=$cols;merges=$merges;comparer=[MonthlyComparer]::new($merges)}
}
if($selected -eq 0){throw 'No selected source merges; do not launch an output comparison.'}
function Release-Com($value){
    if($null -ne $value -and [Runtime.InteropServices.Marshal]::IsComObject($value)){
        try{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($value)}
        catch{$script:cleanupFailures+=$_.Exception.GetType().FullName}
    }
}
$preExisting=@(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object {$_.Id})
$excel=$null;$books=$null;$sourceBook=$null;$outputBook=$null;$sourceSheets=$null;$outputSheets=$null
$ownedPid=0;$ownedCreation=0;$ownsInstance=$false;$readOnlyVerified=$false;$failure=$null;$cleanupFailures=@();$sheetResults=@();$stage='create-owned-excel'
try{
    $beforeCreate=[DateTime]::UtcNow.ToFileTimeUtc()
    $excel=New-Object -ComObject Excel.Application
    [uint32]$instancePid=0
    [void][MonthlyExcelCheck]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$instancePid)
    $ownedPid=[int]$instancePid
    if($ownedPid -eq 0 -or $preExisting -contains $ownedPid){throw 'Could not establish separate Excel ownership.'}
    $ownedCreation=(Get-Process -Id $ownedPid).StartTime.ToUniversalTime().ToFileTimeUtc()
    if($ownedCreation -lt $beforeCreate){throw 'Excel process predates this COM creation request.'}
    $ownsInstance=$true
    Write-ProgressRecord 'owned-excel-confirmed'
    $excel.Visible=$false;$excel.AutomationSecurity=3;$excel.EnableEvents=$false;$excel.AskToUpdateLinks=$false;$excel.DisplayAlerts=$false
    $version=$excel.Version;$build=$excel.Build
    $clientVersion=(Get-Item -LiteralPath (Join-Path $excel.Path 'EXCEL.EXE')).VersionInfo.FileVersion
    $books=$excel.Workbooks;$stage='open-source-read-only';Write-ProgressRecord $stage
    $sourceBook=$books.Open($Source,0,$true,5,'','',$true,2,'',$false,$false,0,$false,$true,0)
    $stage='open-output-read-only';Write-ProgressRecord $stage
    $outputBook=$books.Open($DefaultOutput,0,$true,5,'','',$true,2,'',$false,$false,0,$false,$true,0)
    if(-not $sourceBook.ReadOnly -or -not $outputBook.ReadOnly){throw 'Both workbooks must be read-only.'}
    $readOnlyVerified=$true
    $sourceSheets=$sourceBook.Worksheets;$outputSheets=$outputBook.Worksheets
    if($sourceSheets.Count -ne $plans.Count -or $outputSheets.Count -ne $plans.Count){throw 'Worksheet count differs from inventory.'}
    foreach($plan in $plans){
        $sourceSheet=$null;$outputSheet=$null;$sourceUsed=$null;$outputUsed=$null
        $checked=[long]0;$mismatches=[long]0;$issues=@();$chunks=0
        try{
            $sourceSheet=$sourceSheets.Item($plan.index);$outputSheet=$outputSheets.Item($plan.index)
            $sourceUsed=$sourceSheet.UsedRange;$outputUsed=$outputSheet.UsedRange
            $sourceUsedAddress=$sourceUsed.Address();$outputUsedAddress=$outputUsed.Address()
            $sourceEnd=[MonthlyExcelCheck]::Point($sourceUsedAddress.Split(':')[-1])
            $outputEnd=[MonthlyExcelCheck]::Point($outputUsedAddress.Split(':')[-1])
            if($sourceEnd[0] -gt $plan.rows -or $outputEnd[0] -gt $plan.rows -or $sourceEnd[1] -gt $plan.cols -or $outputEnd[1] -gt $plan.cols){throw 'UsedRange extends beyond independently inventoried comparison bounds.'}
            for($first=1;$first -le $plan.rows;$first+=$ChunkRows){
                $stage='sheet-'+$plan.index+'-rows-'+$first
                if($chunks % 5 -eq 0){Write-ProgressRecord $stage $checked $plan.index}
                $count=[Math]::Min($ChunkRows,$plan.rows-$first+1)
                $rangeAddress='A'+$first+':'+[MonthlyExcelCheck]::Address($first+$count-1,$plan.cols)
                $sourceRange=$null;$outputRange=$null;$sourceValues=$null;$outputValues=$null
                try{
                    $sourceRange=$sourceSheet.Range($rangeAddress);$outputRange=$outputSheet.Range($rangeAddress)
                    $sourceValues=$sourceRange.Value2;$outputValues=$outputRange.Value2
                    $check=$plan.comparer.Compare($sourceValues,$outputValues,$first,$count,$plan.cols)
                    $checked+=$check.Checked;$mismatches+=$check.Mismatches;$chunks++
                    foreach($address in $check.IssueAddresses){if($issues.Count -lt 50){$issues+=$address}}
                }finally{$sourceValues=$null;$outputValues=$null;Release-Com $sourceRange;Release-Com $outputRange}
            }
            $expected=[long]$plan.rows*[long]$plan.cols
            if($checked -ne $expected){throw 'Grid comparison count differs from expected bounds.'}
            $sheetResults+=@{sheet_index=$plan.index;rows=$plan.rows;columns=$plan.cols;chunks=$chunks;checked_cells=$checked;mismatches=$mismatches;issue_addresses=$issues;source_used_range=$sourceUsedAddress;output_used_range=$outputUsedAddress;passed=$mismatches -eq 0}
            Write-ProgressRecord ('sheet-'+$plan.index+'-complete') $checked $plan.index
            Write-Output ('Excel sheet='+$plan.index+'; checked='+$checked+'; mismatches='+$mismatches)
        }finally{Release-Com $sourceUsed;Release-Com $outputUsed;Release-Com $sourceSheet;Release-Com $outputSheet}
    }
}catch{$failure=@{stage=$stage;exception_type=$_.Exception.GetType().FullName}}
finally{
    try{if($ownsInstance){Write-ProgressRecord 'closing-read-only-workbooks'}}catch{$cleanupFailures+=$_.Exception.GetType().FullName}
    Release-Com $sourceSheets;Release-Com $outputSheets
    foreach($book in @($sourceBook,$outputBook)){
        if($null -ne $book){try{$book.Close($false)}catch{$cleanupFailures+=$_.Exception.GetType().FullName}finally{Release-Com $book}}
    }
    Release-Com $books
    if($ownsInstance -and $null -ne $excel){try{$excel.Quit()}catch{$cleanupFailures+=$_.Exception.GetType().FullName}}
    Release-Com $excel
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
}
$sourceAfter=(Get-FileHash -LiteralPath $Source -Algorithm SHA256).Hash.ToLowerInvariant()
$outputAfter=(Get-FileHash -LiteralPath $DefaultOutput -Algorithm SHA256).Hash.ToLowerInvariant()
if($ownsInstance){for($i=0;$i -lt 25 -and (Get-Process -Id $ownedPid -ErrorAction SilentlyContinue);$i++){Start-Sleep -Milliseconds 200}}
$exited=$ownsInstance -and -not [bool](Get-Process -Id $ownedPid -ErrorAction SilentlyContinue)
try{if($ownsInstance){Write-ProgressRecord 'validation-complete'}}catch{$cleanupFailures+=$_.Exception.GetType().FullName}
$passed=$null -eq $failure -and $cleanupFailures.Count -eq 0 -and $sheetResults.Count -eq $plans.Count -and @($sheetResults | Where-Object {-not $_.passed}).Count -eq 0 -and $sourceBefore -eq $sourceAfter -and $outputBefore -eq $outputAfter -and $exited
@{recorded_utc=[DateTime]::UtcNow.ToString('o');script_sha256=(Get-FileHash -LiteralPath $PSCommandPath -Algorithm SHA256).Hash.ToLowerInvariant();inventory_sha256=(Get-FileHash -LiteralPath $Inventory -Algorithm SHA256).Hash.ToLowerInvariant();mode='default';chunk_rows=$ChunkRows;selected_merge_count=$selected;excel_version=$version;excel_build=$build;excel_file_version=$clientVersion;source_sha256=$sourceBefore;output_sha256=$outputBefore;source_unchanged=$sourceBefore -eq $sourceAfter;output_unchanged=$outputBefore -eq $outputAfter;normal_read_only_load=$readOnlyVerified;macro_security=3;update_links=0;saves=0;owned_pid=$ownedPid;owned_excel_exited=$exited;preexisting_excel_pids=$preExisting;sheets=$sheetResults;failure=$failure;cleanup_failures=$cleanupFailures;passed=$passed;privacy='Only coordinates, counts, identities and exception types are recorded; no cell values.';limitations='Local Excel Value2 comparison; at most two 5000-row Value2 arrays plus selected merge anchors are retained by the comparator. Excel itself opens both workbooks and its own memory is not bounded by the chunk size. DisplayAlerts is disabled; no visual/repair-dialog/WPS/macro-execution claim.'} | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $Report -Encoding UTF8
if(-not $passed){throw ('Excel validation failed; see '+$Report)}
