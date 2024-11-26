<#PSScriptInfo .VERSION 1.0.3#>

[CmdletBinding()]
param ([string] $Root, [string] $DirName = 'bin')
[void] (New-Item ($LibDir = "$Root\$DirName") -ItemType Directory -ErrorAction SilentlyContinue)
function ImportTypeLibrary([string] $TypeLibPath, [string] $Namespace) { TlbImp.exe /nologo /silent $TypeLibPath /out:"$LibDir\Interop.$Namespace.dll" /namespace:$Namespace }
ImportTypeLibrary 'C:\Windows\System32\wshom.ocx' 'IWshRuntimeLibrary'