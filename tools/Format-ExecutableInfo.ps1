<#PSScriptInfo .VERSION 1.0.0#>

[CmdletBinding()]
param ([string[]] $LiteralPath)
(Get-Item $LiteralPath).VersionInfo | Format-List -Property @{ Name = "AssemblyName"; Expression = { $_.OriginalFilename } },@{ Name = "AssemblyProduct"; Expression = { $_.ProductName } },@{ Name = "AssemblyTitle"; Expression = { $_.FileDescription } },@{ Name = "AssemblyCompany"; Expression = { $_.CompanyName } },@{ Name = "AssemblyFileVersion"; Expression = { $_.FileVersion } },@{ Name = "AssemblyProductVersion"; Expression = { $_.ProductVersion } },@{ Name = "AssemblyCopyright"; Expression = { $_.LegalCopyright } } -Force