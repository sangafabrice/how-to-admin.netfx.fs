/// <summary>Launch the shortcut's target PowerShell script with the markdown.</summary>
/// <version>0.0.1.0</version>

module cvmd2html.Program

open System
open System.Reflection
open System.Diagnostics
open Microsoft.VisualBasic
open Util
open Parameters
open Package
open Setup
open ErrorLog

/// <summary>Check if the process is elevated.</summary>
/// <returns>True if the running process is elevated, false otherwise.</returns>
let private IsCurrentProcessElevated () =
  let HKU = 0x80000003
  let mutable stdRegProvMethods =
    StdRegProv.GetType().InvokeMember(
      "Methods_",
      BindingFlags.GetProperty,
      null,
      StdRegProv,
      [||]
    )
  let mutable checkAccessMethod =
    stdRegProvMethods.GetType().InvokeMember(
      "Item",
      BindingFlags.InvokeMethod,
      null,
      stdRegProvMethods,
      [|"CheckAccess"|]
    )
  let mutable checkAccessMethodParams =
    checkAccessMethod.GetType().InvokeMember(
      "InParameters",
      BindingFlags.GetProperty,
      null,
      checkAccessMethod,
      [||]
    )
  let mutable inParams =
    checkAccessMethodParams.GetType().InvokeMember(
      "SpawnInstance_",
      BindingFlags.InvokeMethod,
      null,
      checkAccessMethodParams,
      [||]
    )
  let mutable inParamProperties =
    inParams.GetType().InvokeMember(
      "Properties_",
      BindingFlags.GetProperty,
      null,
      inParams,
      [||]
    )
  let mutable hDefKey =
    inParamProperties.GetType().InvokeMember(
      "Item",
      BindingFlags.InvokeMethod,
      null,
      inParamProperties,
      [|"hDefKey"|]
    )
  hDefKey.GetType().InvokeMember(
    "Value",
    BindingFlags.SetProperty,
    null,
    hDefKey,
    [|HKU|]
  ) |> ignore
  let mutable sSubKeyName =
    inParamProperties.GetType().InvokeMember(
      "Item",
      BindingFlags.InvokeMethod,
      null,
      inParamProperties,
      [|"sSubKeyName"|]
    )
  sSubKeyName.GetType().InvokeMember(
    "Value",
    BindingFlags.SetProperty,
    null,
    sSubKeyName,
    [|@"S-1-5-19\Environment"|]
  ) |> ignore
  let mutable outParams =
    StdRegProv.GetType().InvokeMember(
      "ExecMethod_",
      BindingFlags.InvokeMethod,
      null,
      StdRegProv,
      [|"CheckAccess"; inParams|]
    )
  let mutable outParamsProperties =
    outParams.GetType().InvokeMember(
      "Properties_",
      BindingFlags.GetProperty,
      null,
      outParams,
      [||]
    )
  let mutable bGranted =
    outParamsProperties.GetType().InvokeMember(
      "Item",
      BindingFlags.InvokeMethod,
      null,
      outParamsProperties,
      [|"bGranted"|]
    )
  try
    bGranted.GetType().InvokeMember(
      "Value",
      BindingFlags.GetProperty,
      null,
      bGranted,
      [||]
    ) :?> bool
  finally
    ReleaseComObject &bGranted
    ReleaseComObject &outParamsProperties
    ReleaseComObject &outParams
    ReleaseComObject &sSubKeyName
    ReleaseComObject &hDefKey
    ReleaseComObject &inParamProperties
    ReleaseComObject &inParams
    ReleaseComObject &checkAccessMethodParams
    ReleaseComObject &checkAccessMethod
    ReleaseComObject &stdRegProvMethods

/// <summary>Request administrator privileges.</summary>
let private RequestAdminPrivileges () =
  if not (IsCurrentProcessElevated()) then
    let mutable shell = WSH.CreateObject "Shell.Application"
    shell.GetType().InvokeMember(
      "ShellExecute",
      BindingFlags.InvokeMethod,
      null,
      shell,
      [|AssemblyLocation; Interaction.Command(); Missing.Value; "runas"; Constants.vbHidden|]
    ) |> ignore
    ReleaseComObject &shell
    Quit 0

/// <summary>Wait for the process exit.</summary>
/// <param name="processId">The process identifier.</param>
/// <returns>The exit status of the process.</returns>
let private WaitForExit (processId: int) =
  // The process termination event query. Win32_ProcessStopTrace requires admin rights to be used.
  let wqlQuery = "SELECT * FROM Win32_ProcessStopTrace WHERE ProcessName='cmd.exe' AND ProcessId=" + string(processId)
  // Wait for the process to exit.
  let mutable watcher =
    WSH.GetObject().GetType().InvokeMember(
      "ExecNotificationQuery",
      BindingFlags.InvokeMethod,
      null,
      WSH.GetObject(),
      [|wqlQuery|]
    )
  let mutable cmdProcess =
    watcher.GetType().InvokeMember(
      "NextEvent",
      BindingFlags.InvokeMethod,
      null,
      watcher,
      [||]
    )
  let mutable cmdProcessProperties =
    cmdProcess.GetType().InvokeMember(
      "Properties_",
      BindingFlags.GetProperty,
      null,
      cmdProcess,
      [||]
    )
  let mutable ExitStatus =
    cmdProcessProperties.GetType().InvokeMember(
      "Item",
      BindingFlags.InvokeMethod,
      null,
      cmdProcessProperties,
      [|"ExitStatus"|]
    )
  try
    ExitStatus.GetType().InvokeMember(
      "Value",
      BindingFlags.GetProperty,
      null,
      ExitStatus,
      [||]
    ) :?> int
  finally
    ReleaseComObject &ExitStatus
    ReleaseComObject &cmdProcessProperties
    ReleaseComObject &cmdProcess
    ReleaseComObject &watcher

[<EntryPoint>]
let main args =
  RequestAdminPrivileges()
  match CommandLineArgs with
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