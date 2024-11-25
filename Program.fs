/// <summary>Launch the shortcut's target PowerShell script with the markdown.</summary>
/// <version>0.0.1.3</version>

module cvmd2html.Program

open System
open System.Diagnostics
open System.ComponentModel
open System.Reflection
open WbemScripting
open ROOT.CIMV2
open Util
open Parameters
open Package
open Setup
open ErrorLog

[<assembly: AssemblyTitle("CvMd2Html")>]

do ()

/// <summary>Check if the process is elevated.</summary>
/// <returns>True if the running process is elevated, false otherwise.</returns>
let private IsCurrentProcessElevated () =
  let HKU = 0x80000003u
  let mutable bGranted = false
  StdRegProv.CheckAccess (HKU, @"S-1-5-19\Environment", &bGranted) |> ignore
  bGranted

/// <summary>Request administrator privileges.</summary>
/// <param name="args">The command line arguments.</param>
let private RequestAdminPrivileges (args: string array) =
  if not (IsCurrentProcessElevated()) then
    try
      let startInfo = new ProcessStartInfo(
        AssemblyLocation,
        (if args.Length > 0 then String.Format(@"""{0}""", String.Join(@""" """, args)) else "")
      )
      startInfo.UseShellExecute <- true
      startInfo.Verb <- "runas"
      startInfo.WindowStyle <- ProcessWindowStyle.Hidden
      Process.Start startInfo |> ignore
      Quit 0
    with
      | :? Win32Exception -> Quit 0
      | _ -> Quit 1

/// <summary>Wait for the process exit.</summary>
/// <param name="processId">The process identifier.</param>
/// <returns>The exit status of the process.</returns>
let private WaitForExit (processId: int) =
  // The process termination event query. Win32_ProcessStopTrace requires admin rights to be used.
  let wqlQuery = "SELECT * FROM Win32_ProcessStopTrace WHERE ProcessName='cmd.exe' AND ProcessId=" + string(processId)
  // Wait for the process to exit.
  let mutable wbemLocator = new SWbemLocatorClass()
  let mutable wmiService = wbemLocator.ConnectServer()
  let mutable watcher = wmiService.ExecNotificationQuery(wqlQuery)
  let mutable cmdProcess = watcher.NextEvent()
  let mutable cmdProcessProperties = cmdProcess.Properties_
  let mutable ExitStatus = cmdProcessProperties.["ExitStatus"]
  try
    ExitStatus.Value :?> int
  finally
    ReleaseComObject &ExitStatus
    ReleaseComObject &cmdProcessProperties
    ReleaseComObject &cmdProcess
    ReleaseComObject &watcher
    ReleaseComObject &wmiService
    ReleaseComObject &wbemLocator

[<EntryPoint>]
let main args =
  RequestAdminPrivileges args
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
    if (WaitForExit (Process.Start(startInfo).Id)) <> 0 then
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