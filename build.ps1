$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $here
New-Item -ItemType Directory -Force -Path dist | Out-Null
$env:GOOS = 'windows'
$env:GOARCH = 'amd64'
$env:CGO_ENABLED = '0'
go build -trimpath -ldflags '-s -w -H=windowsgui' -o 'dist\hMailServer-Migration-Assistant-v1.0.exe' .
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
Get-FileHash 'dist\hMailServer-Migration-Assistant-v1.0.exe' -Algorithm SHA256 |
    Format-List Algorithm,Hash,Path
