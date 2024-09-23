<#PSScriptInfo .VERSION 1.0.0#>

[CmdletBinding()]
param ([string] $PartialName)
return "$Env:windir\Microsoft.NET\Framework$(if ([Environment]::Is64BitOperatingSystem) { '64' })\v4.0.30319\WPF\$PartialName.dll"