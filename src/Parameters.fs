/// <summary>The parsed parameters.</summary>
/// <version>0.0.1.1</version>

module cvmd2html.Parameters

open System
open System.Text
open Util

type internal InParams =
  /// <summary>The selected markdown file path.</summary>
  | Markdown of Path: string
  /// <summary>The settings to install or uninstall the shortcut menu.</summary>
  /// <param name="Set">Install the shortcut menu if True, Uninstall otherwise.</param>
  /// <param name="NoIcon">Optionally install the shortcut menu without icon.</param>
  | Setting of Set: bool * NoIcon: bool option
  /// <summary>Represent invalidated parameters</summary>
  | Empty

/// <summary>Parse the command line arguments.</summary>
/// <param name="args">Command line arguments.</param>
/// <returns>The parsed input parameters.</returns>
let private Parse (args: string[]) =
  if args.Length = 1 then
    let arg = args.[0].Trim()
    let paramNameValue = arg.Split([|':'|], 2)
    let mutable paramMarkdown = ""
    if paramNameValue.Length = 2 && paramNameValue.[0].Equals("/Markdown", StringComparison.OrdinalIgnoreCase) then
      paramMarkdown <- paramNameValue.[1].Trim()
      if paramMarkdown.Length > 0 then Markdown paramMarkdown
      else Empty
    else
      match arg.ToLower() with
      | "/set" -> Setting(true, Some false)
      | "/set:noicon" -> Setting(true, Some true)
      | "/unset" -> Setting(false, None)
      | _ -> Markdown arg
  elif args.Length = 0 then Setting(true, Some false)
  else Empty

/// <summary>Show the help message when the parameters when they are invalid.</summary>
/// <param name="inParams">Parsed input parameters.</param>
/// <returns>The parsed input parameters.</returns>
let private ShowHelp (inParams: InParams) =
  match inParams with
  | Empty ->
    Dialog.Popup((new StringBuilder())
    .AppendLine("The MarkdownToHtml shortcut launcher.")
    .AppendLine("It starts the shortcut menu target script in a hidden window.")
    .AppendLine()
    .AppendLine("Syntax:")
    .AppendLine("  Convert-MarkdownToHtml [/Markdown:]<markdown file path>")
    .AppendLine("  Convert-MarkdownToHtml [/Set[:NoIcon]]")
    .AppendLine("  Convert-MarkdownToHtml /Unset")
    .AppendLine("  Convert-MarkdownToHtml /Help")
    .AppendLine()
    .AppendLine("<markdown file path>  The selected markdown's file path.")
    .AppendLine("                 Set  Configure the shortcut menu in the registry.")
    .AppendLine("              NoIcon  Specifies that the icon is not configured.")
    .AppendLine("               Unset  Removes the shortcut menu.")
    .AppendLine("                Help  Show the help doc.")
    .ToString()
    )
    Quit 1
  | _ -> ()
  inParams

/// <summary>Parse the command line arguments.</summary>
/// <param name="args">Command line arguments.</param>
/// <returns>The parsed input parameters.</returns>
let internal ParseCommandLine = Parse >> ShowHelp