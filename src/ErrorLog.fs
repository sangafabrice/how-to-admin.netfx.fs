/// <summary>ErrorLog manages the error log file and content.</summary>
/// <version>0.0.1.0</version>

module cvmd2html.ErrorLog

open System.IO
open System.Text.RegularExpressions
open System.Windows
open Util

let private randomLogPath = GenerateRandomPath ".log"

/// <summary>adapted link object.</summary>
[<AbstractClass>]
type internal ErrorLog =
  /// <summary>The path to the generated error log file.</summary>
  static member Path = randomLogPath

  /// <summary>Display the content of the error log file in a message box.</summary>
  static member Read () =
    try
      using (File.OpenText(ErrorLog.Path)) (
        fun txtStream ->
          let errorMessage = Regex.Replace(txtStream.ReadToEnd(), @"(\x1B\[31;1m)|(\x1B\[0m)", "")
          if errorMessage.Length > 0 then Dialog.Popup(errorMessage, MessageBoxImage.Error)
      )
    with
      | _ -> ()

  /// <summary>Delete the error log file.</summary>
  static member Delete () =
    DeleteFile ErrorLog.Path