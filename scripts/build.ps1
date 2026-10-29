param([switch]$Publish)
$ErrorActionPreference='Stop'
Set-Location (Split-Path $PSScriptRoot -Parent)
$env:DOTNET_CLI_HOME="$PWD\.tools\cli"
$env:NUGET_PACKAGES="$PWD\.tools\packages"
$env:DOTNET_CLI_TELEMETRY_OPTOUT='1'
$sdk=Join-Path $PWD '.tools/dotnet/dotnet.exe'
if(!(Test-Path $sdk)){$sdk='dotnet'}
& $sdk build GridPilot.sln -c Release
if($LASTEXITCODE -ne 0){throw 'Falló la compilación.'}
if($Publish){
 & $sdk publish src/GridPilot.App -c Release -r win-x64 --self-contained true -o artifacts/release
 if($LASTEXITCODE -ne 0){throw 'Falló publicación de interfaz.'}
 & $sdk publish src/GridPilot.Mcp -c Release -r win-x64 --self-contained true -o artifacts/release/server
 if($LASTEXITCODE -ne 0){throw 'Falló publicación del servidor.'}
}else{
 $dest='src/GridPilot.App/bin/Release/net10.0-windows/server'
 New-Item -ItemType Directory -Force $dest | Out-Null
 Copy-Item src/GridPilot.Mcp/bin/Release/net10.0-windows/* $dest -Recurse -Force
}
