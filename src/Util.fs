/// <summary>Some utility methods.</summary>
/// <version>0.0.1.0</version>

module cvmd2html.Util

open System
open System.IO
open System.Runtime.InteropServices
open System.Windows
open System.Reflection

let internal AssemblyLocation = Assembly.GetExecutingAssembly().Location

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

/// <summary>Clean up and quit.</summary>
/// <param name="exitCode">The exit code.</param>
let internal Quit (exitCode: int) =
  GC.Collect()
  Environment.Exit exitCode