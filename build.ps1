<#PSScriptInfo .VERSION 1.0.1#>

using namespace System.Management.Automation
[CmdletBinding()]
param ()

& {
  Import-Module "$PSScriptRoot\tools"
  Format-ProjectCode @('*.fs','*.fsproj','*.ps*1','.gitignore'| ForEach-Object { "$PSScriptRoot\$_" })

  $HostColorArgs = @{
    ForegroundColor = 'Black'
    BackgroundColor = 'Green'
  }

  try { Remove-Item "$(($BinDir = "$PSScriptRoot\bin"))\*" -Recurse -ErrorAction Stop }
  catch [ItemNotFoundException] { Write-Host $_.Exception.Message @HostColorArgs }
  catch {
    $HostColorArgs.BackgroundColor = 'Red'
    Write-Host $_.Exception.Message @HostColorArgs
    return
  }
  New-Item $BinDir -ItemType Directory -ErrorAction SilentlyContinue | Out-Null
  Copy-Item "$PSScriptRoot\rsc" -Destination $BinDir -Recurse
  Save-ProjectPackage $PSScriptRoot
  Set-ProjectVersion $PSScriptRoot

  # Compile the source code with dotnet fsc.dll.
  $fsc_dll = dotnet.exe --list-sdks | Select-Object -Last 1 | ForEach-Object { $_ -replace '([\d\.]+) \[(.+)\]','$2\$1\FSharp\fsc.dll' }
  dotnet.exe $fsc_dll /nologo /target:$($DebugPreference -eq 'Continue' ? 'exe':'winexe') /win32icon:"$PSScriptRoot\menu.ico" /win32manifest:"$PSScriptRoot\cvmd2html.manifest" /reference:"$BinDir\Interop.IWshRuntimeLibrary.dll" /reference:$(Get-WpfLibrary PresentationFramework) /reference:$(Get-WpfLibrary PresentationCore) /reference:$(Get-WpfLibrary WindowsBase) /reference:$(Get-CoreLibrary System.Xaml) /reference:$(Get-CoreLibrary System.Management) /reference:$(Get-CoreLibrary System.Text.RegularExpressions) /reference:$(Get-CoreLibrary System.Linq) /reference:$(Get-CoreLibrary System.Core) /reference:$(Get-CoreLibrary System) /reference:$(Get-CoreLibrary mscorlib) /out:$(($ConvertExe = "$BinDir\cvmd2html.exe")) "$(($SrcDir = "$PSScriptRoot\src"))\AssemblyInfo.fs" "$SrcDir\Util.fs" "$SrcDir\Parameters.fs" "$SrcDir\Package.fs" "$SrcDir\Setup.fs" "$SrcDir\ErrorLog.fs" "$PSScriptRoot\Program.fs"

  if ($LASTEXITCODE -eq 0) {
    Write-Host "Output file $ConvertExe written." @HostColorArgs -NoNewline
    Format-ExecutableInfo $ConvertExe
  }

  Remove-Module tools
}