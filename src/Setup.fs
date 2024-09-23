/// <summary>The methods for managing the shortcut menu option: install and uninstall.</summary>
/// <version>0.0.1.0</version>

module cvmd2html.Setup

open System
open Microsoft.Win32
open Util

let [<Literal>] private SHELL_SUBKEY = @"SOFTWARE\Classes\SystemFileAssociations\.md\shell"

let [<Literal>] private VERB = "cthtml"

let private VERB_SUBKEY = String.Format(@"{0}\{1}", SHELL_SUBKEY, VERB)

let private VERB_KEY = String.Format(@"{0}\{1}", Registry.CurrentUser, VERB_SUBKEY)

let [<Literal>] private ICON_VALUENAME = "Icon"

/// <summary>Configure the shortcut menu in the registry.</summary>
let internal SetShortcut () =
  let COMMAND_KEY = VERB_KEY + @"\command";
  let command = String.Format(@"""{0}"" /Markdown:""%1""", AssemblyLocation);
  Registry.SetValue(COMMAND_KEY, null, command);
  Registry.SetValue(VERB_KEY, null, "Convert to &HTML");

/// <summary>Add an icon to the shortcut menu in the registry.</summary>
let internal AddIcon () =
  Registry.SetValue(VERB_KEY, ICON_VALUENAME, AssemblyLocation)

/// <summary>Remove the shortcut icon menu.</summary>
let internal RemoveIcon () =
  use VERB_KEY_OBJ = Registry.CurrentUser.OpenSubKey(VERB_SUBKEY, true)
  if VERB_KEY_OBJ <> null then VERB_KEY_OBJ.DeleteValue(ICON_VALUENAME, false)

/// <summary>Remove the shortcut menu by removing the verb key and subkeys.</summary>
let internal UnsetShortcut () =
  use SHELL_KEY_OBJ = Registry.CurrentUser.OpenSubKey(SHELL_SUBKEY, true)
  if SHELL_KEY_OBJ <> null then SHELL_KEY_OBJ.DeleteSubKeyTree(VERB, false)