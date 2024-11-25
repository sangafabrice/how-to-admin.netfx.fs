<#PSScriptInfo .VERSION 1.0.3#>

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

  $AssemblyInfoFs = "$(($SrcDir = "$PSScriptRoot\src"))\AssemblyInfo.fs"

  function ImportMgmtClass([string] $ClassName) {
    $FileName = $ClassName.Replace('_', '.')
    fsc.exe /nologo /warn:0 /target:library /out:$(($ClassDll = "$BinDir\$FileName.dll")) /reference:"$BinDir\Interop.WbemScripting.dll" $AssemblyInfoFs "$SrcDir\$FileName.fs"
    return $ClassDll
  }

  # Compile the source code with fsc.exe.
  $EnvPath = $Env:Path
  $Env:Path = "$Env:ProgramFiles\Microsoft Visual Studio\2022\Community\Common7\IDE\CommonExtensions\Microsoft\FSharp\Tools\;$Env:Path"
  fsc.exe /nologo /target:$($DebugPreference -eq 'Continue' ? 'exe':'winexe') /win32icon:"$PSScriptRoot\menu.ico" /reference:$(ImportMgmtClass StdRegProv) /reference:"$BinDir\Interop.WbemScripting.dll" /reference:"$BinDir\Interop.IWshRuntimeLibrary.dll" /reference:$(Get-WpfLibrary PresentationFramework) /reference:$(Get-WpfLibrary PresentationCore) /reference:$(Get-WpfLibrary WindowsBase) /reference:System.Xaml.dll /out:$(($ConvertExe = "$BinDir\cvmd2html.exe")) $AssemblyInfoFs "$SrcDir\Util.fs" "$SrcDir\Parameters.fs" "$SrcDir\Package.fs" "$SrcDir\Setup.fs" "$SrcDir\ErrorLog.fs" "$PSScriptRoot\Program.fs"
  $Env:Path = $EnvPath

  if ($LASTEXITCODE -eq 0) {
    Write-Host "Output file $ConvertExe written." @HostColorArgs -NoNewline
    Format-ExecutableInfo $ConvertExe
  }

  Remove-Module tools
}