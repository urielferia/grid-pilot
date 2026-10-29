$ErrorActionPreference='Stop'
Set-Location (Split-Path $PSScriptRoot -Parent)
$sdk=Join-Path $PWD '.tools/dotnet/dotnet.exe'
$info=[System.Diagnostics.ProcessStartInfo]::new($sdk)
$info.ArgumentList.Add((Join-Path $PWD 'src/GridPilot.Mcp/bin/Release/net10.0-windows/GridPilot.Mcp.dll'))
$info.UseShellExecute=$false
$info.CreateNoWindow=$true
$info.RedirectStandardInput=$true
$info.RedirectStandardOutput=$true
$server=[System.Diagnostics.Process]::Start($info)
try {
 $requests=@(
  @{jsonrpc='2.0';id=1;method='initialize';params=@{protocolVersion='2025-06-18';capabilities=@{};clientInfo=@{name='test';version='1'}}},
  @{jsonrpc='2.0';id=2;method='tools/list'},
  @{jsonrpc='2.0';id=3;method='ping'}
 )
 foreach($request in $requests){$server.StandardInput.WriteLine(($request|ConvertTo-Json -Depth 12 -Compress));$server.StandardInput.Flush();$response=$server.StandardOutput.ReadLine()|ConvertFrom-Json;if($response.error){throw 'Error MCP.'};if($response.id -ne $request.id){throw 'Secuencia RPC incorrecta.'};if($request.id -eq 2){$names=$response.result.tools.name;if($names.Count -ne 7 -or ($names -match 'approve|export')){throw 'Catálogo de herramientas incorrecto.'}}}
 'PASS MCP: initialize, catálogo de 7 herramientas sin aprobación/exportación, ping y cierre por EOF.'
}finally{$server.StandardInput.Close();if(!$server.WaitForExit(10000)){throw 'El servidor no cerró por EOF.'};$server.Dispose()}
