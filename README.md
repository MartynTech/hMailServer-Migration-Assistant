# hMailServer Migration Assistant v1.0

Windows x64 migration utility for hMailServer installations using Microsoft SQL Server Compact Edition (MSSQL CE).

## Build

Requirements:
- Go 1.23 or later
- Windows, or Go cross-compilation support

From PowerShell on Windows:

```powershell
.\build.ps1
```

Or from CMD:

```cmd
build.cmd
```

The output is written to `dist\hMailServer-Migration-Assistant-v1.0.exe`.

## Scope

The application supports:
- Source/destination hMailServer installation discovery
- Manual INI/database/Data/Events path overrides
- MSSQL CE validation
- Robocopy `/L` mail-store enumeration
- Standard Robocopy staged migration
- Optional 7-Zip source-side archive staging through PowerShell remoting
- Final cutover with database/configuration backup
- Destination-only `.eml` preservation
- Certificate/DKIM discovery via the hMailServer COM API
- Verification, DNS/mail-flow checks, HTML reporting and rollback

## Important

Always maintain an independent, verified backup before migrating a production mail system.

The software is an independent administration utility and is not affiliated with the hMailServer project.

Author: Martyn Beech
