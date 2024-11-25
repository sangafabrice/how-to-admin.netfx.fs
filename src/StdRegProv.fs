/// <summary>StdRegProv WMI class as inspired by mgmclassgen.exe.</summary>
/// <version>0.0.1.0</version>

namespace ROOT.CIMV2

open System
open System.Diagnostics
open System.Runtime.InteropServices
open System.Reflection
open WbemScripting

[<assembly: AssemblyTitle("StdRegProv")>]

do ()

module StdRegProv =
  let private CreatedClassName = (new StackTrace()).GetFrame(0).GetMethod().DeclaringType.Name.Substring(1)

  /// <summary>Get the name of the method calling this method.</summary>
  /// <param name="stackTrace">The stack trace from the calling method.</param>
  /// <returns>The name of the caller method.</returns>
  let private GetMethodName (stackTrace: StackTrace) =
    stackTrace.GetFrame(0).GetMethod().Name

  /// <summary>Release the specified COM object.</summary>
  /// <param name="comObject">The COM object to destroy.</param>
  let private ReleaseComObject<'a> (comObject: 'a byref) =
    Marshal.FinalReleaseComObject comObject |> ignore
    comObject <- Unchecked.defaultof<'a>

  /// <summary>Set the specified property of a SWbemObject instance.</summary>
  /// <param name="inParams">The object to set the property value from.</param>
  /// <param name="propertyName">The property name.</param>
  /// <param name="propertyValue">The property value.</param>
  let private SetWBemObjectProperty (inParams: SWbemObject byref) (propertyName: string) (propertyValue) =
    let mutable inProperties = inParams.Properties_
    let mutable property = inProperties.[propertyName]
    property.Value <- ref propertyValue
    ReleaseComObject &property
    ReleaseComObject &inProperties

  /// <summary>Get the specified property of a SWbemObject instance.</summary>
  /// <param name="outParams">The object to get the property value from.</param>
  /// <param name="propertyName">The property name.</param>
  /// <returns>The property value.</returns>
  let private GetWBemObjectProperty<'a> (outParams: SWbemObject byref) (propertyName: string): 'a =
    let mutable outProperties = outParams.Properties_
    let mutable property = outProperties.Item(propertyName)
    try
      property.Value :?> 'a
    finally
      ReleaseComObject &property
      ReleaseComObject &outProperties

  /// <summary>Get the return value from the specified output parameter object.</summary>
  /// <param name="outParams">The object to get the property value from.</param>
  /// <returns>The property value.</returns>
  let private GetReturnValue (outParams: SWbemObject byref) =
    GetWBemObjectProperty<int> &outParams "ReturnValue"

  type private StdRegInput (classObj: SWbemObject byref, methodName: string) =
    let mutable wbemMethodSet = classObj.Methods_
    let mutable wbemMethod = wbemMethodSet.[methodName]
    let mutable inParamsClass = wbemMethod.InParameters
    let mutable _params = inParamsClass.SpawnInstance_()
    member val Params = _params with get

    interface IDisposable with
      member this.Dispose (): unit =
        ReleaseComObject &_params
        ReleaseComObject &inParamsClass
        ReleaseComObject &wbemMethod
        ReleaseComObject &wbemMethodSet

  type private StdRegProvider () =
    let mutable wbemLocator = new SWbemLocatorClass()
    let mutable wbemService = wbemLocator.ConnectServer()
    let mutable provider = wbemService.Get(CreatedClassName)
    member val Provider = provider with get

    interface IDisposable with
      member this.Dispose (): unit =
        ReleaseComObject &provider
        ReleaseComObject &wbemService
        ReleaseComObject &wbemLocator

  let CheckAccess (hDefKey: uint, sSubKeyName: string, bGranted: bool byref) =
    let methodName = GetMethodName (new StackTrace())
    use Registry = new StdRegProvider()
    let mutable classObj = Registry.Provider
    use Input = new StdRegInput(&classObj, methodName)
    let mutable inParams = Input.Params
    SetWBemObjectProperty &inParams "hDefKey" (int hDefKey)
    SetWBemObjectProperty &inParams "sSubKeyName" sSubKeyName
    let mutable outParams = classObj.ExecMethod_(methodName, inParams)
    bGranted <- GetWBemObjectProperty &outParams "bGranted"
    try
      GetReturnValue &outParams
    finally
      ReleaseComObject &inParams
      ReleaseComObject &outParams
      ReleaseComObject &classObj

  let CreateKey (hDefKey: uint, sSubKeyName: string) =
    let methodName = GetMethodName (new StackTrace())
    use Registry = new StdRegProvider()
    let mutable classObj = Registry.Provider
    use Input = new StdRegInput(&classObj, methodName)
    let mutable inParams = Input.Params
    SetWBemObjectProperty &inParams "hDefKey" (int hDefKey)
    SetWBemObjectProperty &inParams "sSubKeyName" sSubKeyName
    let mutable outParams = classObj.ExecMethod_(methodName, inParams)
    try
      GetReturnValue &outParams
    finally
      ReleaseComObject &inParams
      ReleaseComObject &outParams
      ReleaseComObject &classObj

  let DeleteKey (hDefKey: uint, sSubKeyName: string) =
    let methodName = GetMethodName (new StackTrace())
    use Registry = new StdRegProvider()
    let mutable classObj = Registry.Provider
    use Input = new StdRegInput(&classObj, methodName)
    let mutable inParams = Input.Params
    SetWBemObjectProperty &inParams "hDefKey" (int hDefKey)
    SetWBemObjectProperty &inParams "sSubKeyName" sSubKeyName
    let mutable outParams = classObj.ExecMethod_(methodName, inParams)
    try
      GetReturnValue &outParams
    finally
      ReleaseComObject &inParams
      ReleaseComObject &outParams
      ReleaseComObject &classObj

  let DeleteValue (hDefKey: uint, sSubKeyName: string, sValueName: string | null) =
    let methodName = GetMethodName (new StackTrace())
    use Registry = new StdRegProvider()
    let mutable classObj = Registry.Provider
    use Input = new StdRegInput(&classObj, methodName)
    let mutable inParams = Input.Params
    SetWBemObjectProperty &inParams "hDefKey" (int hDefKey)
    SetWBemObjectProperty &inParams "sSubKeyName" sSubKeyName
    SetWBemObjectProperty &inParams "sValueName" sValueName
    let mutable outParams = classObj.ExecMethod_(methodName, inParams)
    try
      GetReturnValue &outParams
    finally
      ReleaseComObject &inParams
      ReleaseComObject &outParams
      ReleaseComObject &classObj

  let EnumKey (hDefKey: uint, sSubKeyName: string, sNames: string array byref) =
    let methodName = GetMethodName (new StackTrace())
    use Registry = new StdRegProvider()
    let mutable classObj = Registry.Provider
    use Input = new StdRegInput(&classObj, methodName)
    let mutable inParams = Input.Params
    SetWBemObjectProperty &inParams "hDefKey" (int hDefKey)
    SetWBemObjectProperty &inParams "sSubKeyName" sSubKeyName
    let mutable outParams = classObj.ExecMethod_(methodName, inParams)
    try
      let sNamesObj = GetWBemObjectProperty<obj> &outParams "sNames"
      sNames <-
        match sNamesObj with
        | :? DBNull -> Array.Empty()
        | :? array<obj> as names -> Array.map (fun name -> name.ToString()) names
        | _ -> [||]
      GetReturnValue &outParams
    finally
      ReleaseComObject &inParams
      ReleaseComObject &outParams
      ReleaseComObject &classObj

  let GetStringValue (hDefKey: uint Nullable, sSubKeyName: string, sValueName: string | null, sValue: string byref) =
    let methodName = GetMethodName (new StackTrace())
    use Registry = new StdRegProvider()
    let mutable classObj = Registry.Provider
    use Input = new StdRegInput(&classObj, methodName)
    let mutable inParams = Input.Params
    if hDefKey.HasValue then SetWBemObjectProperty &inParams "hDefKey" (int hDefKey.Value)
    SetWBemObjectProperty &inParams "sSubKeyName" sSubKeyName
    SetWBemObjectProperty &inParams "sValueName" sValueName
    let mutable outParams = classObj.ExecMethod_(methodName, inParams)
    sValue <- GetWBemObjectProperty &outParams "sValue"
    try
      GetReturnValue &outParams
    finally
      ReleaseComObject &inParams
      ReleaseComObject &outParams
      ReleaseComObject &classObj

  let SetStringValue (hDefKey: uint, sSubKeyName: string, sValueName: string | null, sValue: string | null) =
    let methodName = GetMethodName (new StackTrace())
    use Registry = new StdRegProvider()
    let mutable classObj = Registry.Provider
    use Input = new StdRegInput(&classObj, methodName)
    let mutable inParams = Input.Params
    SetWBemObjectProperty &inParams "hDefKey" (int hDefKey)
    SetWBemObjectProperty &inParams "sSubKeyName" sSubKeyName
    SetWBemObjectProperty &inParams "sValueName" sValueName
    SetWBemObjectProperty &inParams "sValue" sValue
    let mutable outParams = classObj.ExecMethod_(methodName, inParams)
    try
      GetReturnValue &outParams
    finally
      ReleaseComObject &inParams
      ReleaseComObject &outParams
      ReleaseComObject &classObj

  /// <summary>Remove the key and all descendant subkeys.</summary>
  let rec DeleteKeyTree (hDefKey, sSubKeyName) =
    let mutable sNames = Array.Empty()
    let mutable returnValue = EnumKey(hDefKey, sSubKeyName, &sNames)
    if sNames.Length > 0 then
      for sName in sNames do
        returnValue <- returnValue + DeleteKeyTree (hDefKey, $"{sSubKeyName}\\{sName}")
    returnValue <- returnValue + DeleteKey(hDefKey, sSubKeyName)
    returnValue