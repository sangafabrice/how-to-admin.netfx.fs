/// <summary>Some utility methods.</summary>
/// <version>0.0.1.0</version>

module cvmd2html.Util

open System
open System.IO
open System.Runtime.InteropServices
open System.Windows
open System.Reflection

let internal AssemblyLocation = Assembly.GetExecutingAssembly().Location

let mutable private wmiService = null

type internal WSH =
  /// <summary>Create object.</summary>
  /// <param name="progId">The com class ProgId.</param>
  /// <returns>A COM object.</returns>
  static member CreateObject (progId) =
    Activator.CreateInstance(Type.GetTypeFromProgID(progId))

  /// <summary>Get a WMI object or class.</summary>
  /// <param name="monikerPath">The moniker path.</param>
  /// <returns>A WMI object or class.</returns>
  static member GetObject (?monikerPath: string) =
    let monikerPath = defaultArg monikerPath ""
    if String.IsNullOrEmpty(monikerPath) then wmiService
    else
      wmiService.GetType().InvokeMember(
        "Get",
        BindingFlags.InvokeMethod,
        null,
        wmiService,
        [|monikerPath|]
      )

let mutable private wbemLocator = WSH.CreateObject "WbemScripting.SWbemLocator"

wmiService <-
  wbemLocator.GetType().InvokeMember(
    "ConnectServer",
    BindingFlags.InvokeMethod,
    null,
    wbemLocator,
    [||]
  )

/// <summary>The registry com object.</summary>
let mutable internal StdRegProv = WSH.GetObject "StdRegProv"

/// <summary>Generate a random file path.</summary>
/// <param name="extension">The file extension.</param>
/// <returns>A random file path.</returns>
let internal GenerateRandomPath (extension): string =
  Path.Combine(Path.GetTempPath(), (string (Guid.NewGuid())) + ".tmp" + extension)

/// <summary>Delete the specified file.</summary>
/// <param name="filePath">The file path.</param>
let internal DeleteFile (filePath) =
  try
    File.Delete filePath
  with
    | _ -> ()

type internal Dialog =
  /// <summary>Show the application message box.</summary>
  /// <param name="messageText">The message text to show.</param>
  /// <param name="popupType">The type of popup box.</param>
  /// <param name="popupButtons">The buttons of the message box.</param>
  static member Popup (messageText, ?popupType: MessageBoxImage, ?popupButtons: MessageBoxButton) =
    let popupType = defaultArg popupType MessageBoxImage.None
    let popupButtons = defaultArg popupButtons MessageBoxButton.OK
    MessageBox.Show(messageText, "Convert to HTML", popupButtons, popupType) |> ignore

/// <summary>Release the specified COM object.</summary>
/// <param name="comObject">The COM object to destroy.</param>
let internal ReleaseComObject<'a> (comObject: byref<'a>) =
  Marshal.FinalReleaseComObject comObject |> ignore
  comObject <- Unchecked.defaultof<'a>

/// <summary>Destroy the COM objects.</summary>
let internal Dispose () =
  ReleaseComObject &StdRegProv
  ReleaseComObject &wmiService
  ReleaseComObject &wbemLocator

/// <summary>Clean up and quit.</summary>
/// <param name="exitCode">The exit code.</param>
let internal Quit (exitCode: int) =
  Dispose()
  GC.Collect()
  Environment.Exit exitCode