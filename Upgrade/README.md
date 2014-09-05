# PostSharp batch upgrade
This script upgrades PostSharp to version specified by -postSharpVersion parameter in all
projects (csproj, vbproj) recursively found in directory specified by -path parameter.
Supported are projects referencing PostSharp 2.x, 3.x and 4.x. Projects without PostSharp are
not processed.

Script requires at least PowerShell 3 and Microsoft.Build v4.0. Script can run on machine
without Visual Studio.

To get command line help run:
Get-Help .\Upgrade-PostSharp.ps1

# Limitations
- Only .NET Framework targets are supported.
- Projects referencing PostSharp toolkits are skipped.
- Script must be run from directory with nuget.exe.