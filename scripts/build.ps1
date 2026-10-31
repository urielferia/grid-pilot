param([switch]$Publish,[string]$Output='artifacts/release')
$ErrorActionPreference='Stop'
Set-Location (Split-Path $PSScriptRoot -Parent)
$env:DOTNET_CLI_HOME="$PWD\.tools\cli"
$env:APPDATA="$PWD\.tools\cli\AppData\Roaming"
$env:LOCALAPPDATA="$PWD\.tools\cli\AppData\Local"
$env:NUGET_PACKAGES="$PWD\.tools\packages"
$env:DOTNET_CLI_TELEMETRY_OPTOUT='1'
New-Item -ItemType Directory -Force $env:APPDATA,$env:LOCALAPPDATA | Out-Null
$sdk=Join-Path $PWD '.tools/dotnet/dotnet.exe'
if(!(Test-Path $sdk)){$sdk='dotnet'}
& $sdk build GridPilot.sln -c Release -nodeReuse:false -m:1
if($LASTEXITCODE -ne 0){throw 'Build failed.'}
if($Publish){
 & $sdk publish src/GridPilot.App -c Release -r win-x64 --self-contained true -nodeReuse:false -m:1 -o $Output
 if($LASTEXITCODE -ne 0){throw 'UI publish failed.'}
 & $sdk publish src/GridPilot.Mcp -c Release -r win-x64 --self-contained true -nodeReuse:false -m:1 -o "$Output/server"
 if($LASTEXITCODE -ne 0){throw 'Server publish failed.'}
 New-Item -ItemType Directory -Force "$Output/docs" | Out-Null
 Copy-Item README.md "$Output/README.md" -Force
 Copy-Item docs/*.md "$Output/docs/" -Force
 Copy-Item tests/evaluation-30.json "$Output/docs/evaluation-30.json" -Force
}else{
 $dest='src/GridPilot.App/bin/Release/net10.0-windows/server'
 New-Item -ItemType Directory -Force $dest | Out-Null
 Copy-Item src/GridPilot.Mcp/bin/Release/net10.0-windows/* $dest -Recurse -Force
}
