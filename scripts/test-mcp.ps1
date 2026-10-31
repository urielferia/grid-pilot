param([switch]$Excel,[switch]$Published)
$ErrorActionPreference='Stop'
Set-Location (Split-Path $PSScriptRoot -Parent)
$sdk=Join-Path $PWD '.tools/dotnet/dotnet.exe'
$info=[System.Diagnostics.ProcessStartInfo]::new($sdk)
if($Published){$info.FileName=Join-Path $PWD 'artifacts/release/server/GridPilot.Mcp.exe'}
else{$info.FileName=Join-Path $PWD 'src/GridPilot.Mcp/bin/Release/net10.0-windows/GridPilot.Mcp.exe'}
$info.UseShellExecute=$false
$info.CreateNoWindow=$true
$info.RedirectStandardInput=$true
$info.RedirectStandardOutput=$true
[Console]::InputEncoding=[System.Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding=[System.Text.UTF8Encoding]::new($false)
$server=[System.Diagnostics.Process]::Start($info)
try {
 $requests=@(
  @{jsonrpc='2.0';id=1;method='initialize';params=@{protocolVersion='2025-06-18';capabilities=@{};clientInfo=@{name='test';version='1'}}},
  @{jsonrpc='2.0';id=2;method='tools/list'},
  @{jsonrpc='2.0';id=3;method='ping'}
 )
 if($Excel){
  $original=Join-Path $PWD 'assets/GridPilot-Demo.xlsx'
  $before=(Get-FileHash $original).Hash
  $requests+=@(
   @{jsonrpc='2.0';id=4;method='gridpilot/open';params=@{path=$original;root=(Join-Path $PWD 'artifacts/mcp-smoke');profile=(Get-Content assets/profile.json -Raw -Encoding UTF8|ConvertFrom-Json)}},
   @{jsonrpc='2.0';id=5;method='tools/call';params=@{name='recalculate_validate';arguments=@{}}},
   @{jsonrpc='2.0';id=6;method='gridpilot/close'}
  )
 }
 foreach($request in $requests){$server.StandardInput.WriteLine(($request|ConvertTo-Json -Depth 12 -Compress));$server.StandardInput.Flush();$line=$server.StandardOutput.ReadLineAsync();if(!$line.Wait(120000)){throw 'MCP timeout.'};$response=$line.Result|ConvertFrom-Json;if($response.error){throw $response.error.message};if($response.id -ne $request.id){throw 'Out-of-sequence RPC response.'};if($request.id -eq 2){$names=$response.result.tools.name;if($names.Count -ne 7 -or ($names -match 'approve|export')){throw 'Invalid tool catalog.'}};if($request.id -eq 4 -and !$response.result.Editable){throw 'Template is not editable.'};if($request.id -eq 5 -and $response.result.isError){throw $response.result.content[0].text}}
 if($Excel -and (Get-FileHash $original).Hash -ne $before){throw 'Original file was modified.'}
 if($Excel){ 'PASS MCP Excel: editable copy, recalculate/validate, close, and original intact.' }
 'PASS MCP: initialize, 7-tool catalog without approval/export, ping, and EOF shutdown.'
}finally{$server.StandardInput.Close();if(!$server.WaitForExit(10000)){throw 'Server did not exit on EOF.'};$server.Dispose()}
