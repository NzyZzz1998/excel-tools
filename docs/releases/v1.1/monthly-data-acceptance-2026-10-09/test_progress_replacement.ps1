# Exercise the actual progress writer without starting Excel or reading business files.
$ErrorActionPreference='Stop'
$scriptPath=Join-Path $PSScriptRoot 'validate_excel_chunked.ps1'
$tokens=$null;$errors=$null
$ast=[Management.Automation.Language.Parser]::ParseFile($scriptPath,[ref]$tokens,[ref]$errors)
if($errors.Count){throw 'PowerShell parse errors'}
$writer=$ast.Find({param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Write-ProgressRecord'},$true)
Invoke-Expression $writer.Extent.Text
$directory=Join-Path ([IO.Path]::GetTempPath()) ('excel-progress-regression-'+[guid]::NewGuid().ToString())
[void][IO.Directory]::CreateDirectory($directory)
try{
    $Progress=Join-Path $directory 'progress.json';$RunId='synthetic';$producerCreation=100
    $ownedPid=123;$ownedCreation=200;$ownsInstance=$true
    Write-ProgressRecord 'first'
    Write-ProgressRecord 'second' 25000 1
    Write-ProgressRecord 'third' 50000 1
    $actual=Get-Content -LiteralPath $Progress -Raw | ConvertFrom-Json
    if($actual.stage -ne 'third' -or $actual.checked_cells_in_sheet -ne 50000 -or $actual.owned_excel_pid -ne 123){throw 'Progress replacement assertion failed'}
    if(Test-Path -LiteralPath ($Progress+'.writing')){throw 'Temporary progress residue'}
    'PASS: actual progress function created and replaced JSON twice; final identity/stage/count retained; no temporary residue; no Excel launched.'
}finally{[IO.Directory]::Delete($directory,$true)}
