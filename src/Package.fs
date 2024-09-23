/// <summary>Information about the resource files used by the project.</summary>
/// <version>0.0.1.0</version>

module cvmd2html.Package

open System
open System.IO
open IWshRuntimeLibrary
open Microsoft.Win32
open Util

let [<Literal>] private POWERSHELL_SUBKEY = @"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\pwsh.exe"

/// <summary>The project root path</summary>
let private root = AppContext.BaseDirectory

/// <summary>The project resources directory path.</summary>
let private resourcePath = Path.Combine(root, "rsc")

/// <summary>The powershell core runtime path.</summary>
let internal PwshExePath = Registry.GetValue(POWERSHELL_SUBKEY, null, null) |> string

/// <summary>The shortcut target powershell script path.</summary>
let internal PwshScriptPath = Path.Combine(resourcePath, "cvmd2html.ps1")

let private randomLinkPath = GenerateRandomPath ".lnk"

/// <summary>adapted link object.</summary>
[<AbstractClass>]
type internal IconLink =
  /// <summary>The custom icon link full path.</summary>
  static member Path = randomLinkPath

  /// <summary>Create the custom icon link file.</summary>
  /// <param name="markdownPath">The input markdown file path.</param>
  static member Create (markdownPath: string) =
    let mutable shell = new WshShellClass()
    let mutable link = shell.CreateShortcut(IconLink.Path) :?> IWshShortcut
    link.TargetPath <- PwshExePath
    link.Arguments <- String.Format("""-ep Bypass -nop -w Hidden -f "{0}" -Markdown "{1}" """, PwshScriptPath, markdownPath)
    link.IconLocation <- AssemblyLocation
    link.Save()
    ReleaseComObject &link
    ReleaseComObject &shell

  /// <summary>Delete the custom icon link file.</summary>
  static member Delete () =
    DeleteFile IconLink.Path