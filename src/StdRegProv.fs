/// <summary>StdRegProv WMI class as inspired by mgmclassgen.exe.</summary>
/// <version>0.0.1.1</version>

namespace ROOT.CIMV2

open System
open System.Diagnostics
open System.Reflection
open System.Management

[<assembly: AssemblyTitle("StdRegProv")>]

do ()

module StdRegProv =
  let private CreatedClassName = (new StackTrace()).GetFrame(0).GetMethod().DeclaringType.Name.Substring(1)

  /// <summary>Get the name of the method calling this method.</summary>
  /// <param name="stackTrace">The stack trace from the calling method.</param>
  /// <returns>The name of the caller method.</returns>
  let private GetMethodName (stackTrace: StackTrace) =
    stackTrace.GetFrame(0).GetMethod().Name

  let CheckAccess (hDefKey: uint) (sSubKeyName: string) (bGranted: bool byref) =
    let methodName = GetMethodName (new StackTrace())
    let mutable classObj = new ManagementClass(CreatedClassName)
    let mutable inParams = classObj.GetMethodParameters methodName
    inParams.["hDefKey"] <- hDefKey
    inParams.["sSubKeyName"] <- sSubKeyName
    let mutable outParams = classObj.InvokeMethod(methodName, inParams, null)
    bGranted <- outParams.["bGranted"] :?> bool
    outParams.["ReturnValue"] :?> uint

  let CreateKey (hDefKey: uint) (sSubKeyName: string) =
    let methodName = GetMethodName (new StackTrace())
    let mutable classObj = new ManagementClass(CreatedClassName)
    let mutable inParams = classObj.GetMethodParameters methodName
    inParams.["hDefKey"] <- hDefKey
    inParams.["sSubKeyName"] <- sSubKeyName
    let mutable outParams = classObj.InvokeMethod(methodName, inParams, null)
    outParams.["ReturnValue"] :?> uint

  let DeleteKey (hDefKey: uint) (sSubKeyName: string) =
    let methodName = GetMethodName (new StackTrace())
    let mutable classObj = new ManagementClass(CreatedClassName)
    let mutable inParams = classObj.GetMethodParameters methodName
    inParams.["hDefKey"] <- hDefKey
    inParams.["sSubKeyName"] <- sSubKeyName
    let mutable outParams = classObj.InvokeMethod(methodName, inParams, null)
    outParams.["ReturnValue"] :?> uint

  let DeleteValue (hDefKey: uint) (sSubKeyName: string) (sValueName: string | null) =
    let methodName = GetMethodName (new StackTrace())
    let mutable classObj = new ManagementClass(CreatedClassName)
    let mutable inParams = classObj.GetMethodParameters methodName
    inParams.["hDefKey"] <- hDefKey
    inParams.["sSubKeyName"] <- sSubKeyName
    inParams.["sValueName"] <- sValueName
    let mutable outParams = classObj.InvokeMethod(methodName, inParams, null)
    outParams.["ReturnValue"] :?> uint

  let EnumKey (hDefKey: uint) (sSubKeyName: string) (sNames: string array byref) =
    let methodName = GetMethodName (new StackTrace())
    let mutable classObj = new ManagementClass(CreatedClassName)
    let mutable inParams = classObj.GetMethodParameters methodName
    inParams.["hDefKey"] <- hDefKey
    inParams.["sSubKeyName"] <- sSubKeyName
    let mutable outParams = classObj.InvokeMethod(methodName, inParams, null)
    let sNamesObj = outParams.["sNames"]
    sNames <-
      match sNamesObj with
      | :? array<obj> as names -> Array.map (fun name -> name.ToString()) names
      | _ -> [||]
    outParams.["ReturnValue"] :?> uint

  let GetStringValue (hDefKey: uint Nullable) (sSubKeyName: string) (sValueName: string | null) (sValue: string byref) =
    let methodName = GetMethodName (new StackTrace())
    let mutable classObj = new ManagementClass(CreatedClassName)
    let mutable inParams = classObj.GetMethodParameters methodName
    if hDefKey.HasValue then inParams.["hDefKey"] <- (hDefKey.Value)
    inParams.["sSubKeyName"] <- sSubKeyName
    inParams.["sValueName"] <- sValueName
    let mutable outParams = classObj.InvokeMethod(methodName, inParams, null)
    sValue <- outParams.["sValue"] :?> string
    outParams.["ReturnValue"] :?> uint

  let SetStringValue (hDefKey: uint) (sSubKeyName: string) (sValueName: string | null) (sValue: string | null) =
    let methodName = GetMethodName (new StackTrace())
    let mutable classObj = new ManagementClass(CreatedClassName)
    let mutable inParams = classObj.GetMethodParameters methodName
    inParams.["hDefKey"] <- hDefKey
    inParams.["sSubKeyName"] <- sSubKeyName
    inParams.["sValueName"] <- sValueName
    inParams.["sValue"] <- sValue
    let mutable outParams = classObj.InvokeMethod(methodName, inParams, null)
    outParams.["ReturnValue"] :?> uint

  /// <summary>Remove the key and all descendant subkeys.</summary>
  let rec DeleteKeyTree hDefKey sSubKeyName =
    let mutable sNames = Array.Empty()
    let mutable returnValue = EnumKey hDefKey sSubKeyName &sNames
    if sNames.Length > 0 then
      for sName in sNames do
        returnValue <- returnValue + DeleteKeyTree hDefKey $"{sSubKeyName}\\{sName}"
    returnValue <- returnValue + DeleteKey hDefKey sSubKeyName
    returnValue