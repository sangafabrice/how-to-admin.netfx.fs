/// <summary>Launch the shortcut's target PowerShell script with the markdown.</summary>
/// <version>0.0.1.1</version>

module cvmd2html.Program

open System
open System.Diagnostics
open System.Management
open Util
open Parameters
open Package
open Setup
open ErrorLog

/// <summary>Wait for the process exit.</summary>
/// <param name="processId">The process identifier.</param>
/// <returns>The exit status of the process.</returns>
let private WaitForExit (processId: int) =
  // The process termination event query. Win32_ProcessStopTrace requires admin rights to be used.
  let wqlQuery = "SELECT * FROM Win32_ProcessStopTrace WHERE ProcessName='cmd.exe' AND ProcessId=" + string(processId)
  // Wait for the process to exit.
  (new ManagementEventWatcher(wqlQuery)).WaitForNextEvent().["ExitStatus"] :?> uint

[<EntryPoint>]
let main args =
  match ParseCommandLine args with
  | Markdown(path) ->
    let CMD_EXE = @"C:\Windows\System32\cmd.exe"
    let CMD_LINE_FORMAT = @"/d /c """"{0}"" 2> ""{1}"""""
    IconLink.Create(path)
    let startInfo = new ProcessStartInfo(
      CMD_EXE,
      String.Format(CMD_LINE_FORMAT, IconLink.Path, ErrorLog.Path)
    )
    startInfo.WindowStyle <- ProcessWindowStyle.Hidden
    if (WaitForExit (Process.Start(startInfo).Id)) <> 0u then
      ErrorLog.Read()
      ErrorLog.Delete()
    IconLink.Delete()
    Quit 0
    0
  | Setting(config, noIcon) when config ->
    SetShortcut()
    if noIcon = Some true then RemoveIcon()
    else AddIcon()
    Quit 0
    0
  | Setting(config, _) when not config ->
    UnsetShortcut()
    Quit 0
    0
  | _ ->
  Quit 1
  1