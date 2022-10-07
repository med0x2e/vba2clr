### TLDR:
Just experimenting with different ways to load CLR (.NET) assemblies on VBA locally and remotely by using `AppDomain.ExecuteAssembly` or any other handy method after getting arround `AccessVBOM` (programmatic access to visual basic project is not trusted) from VBA.

 * `vba2clr.*.vba`: 
   * Sets AccessVBOM regkey to 1
   * Instantiate a `Word.Application` COM object. (could be `Excel.Application` MS PowerPoint, Access ..etc).
   * Add Macro From String (Macro corresponds to b64/hex encoded ExecuteAssembly.vba)
   * Run ExecuteAssembly.vba Macro using `wordObj.Application.Run...`

### .NET from VBA:
* `ExecuteAssembly.vba`: 
   - Adds the required mscordlib references
   - Instantiates the required objects (`IDomain`, `ICRHost`)
   - Pack the required `AppDomain.ExecuteAssembly` arguments into two separate arrays (variables, types).
   - Use DispCallFunc to call `AppDomain.ExecuteAssembly(Arg1, Arg2)` (VFTable offset 51) where `Arg1` is the ".NET Assembly URL" or "Local Path" and `Arg2` is the return value.
   - AppDomain methods VFTable offsets can be checked on the AppDomain IDL <a href="https://github.com/med0x2e/VBACLR/blob/main/_AppDomain.idl.not.cs">_AppDomain.idl</a>, just keep in mind that the AppDomain interface inherits from the IUnknown interface, so functions/methods VTable offsets start from the third offset onwards, this is because interfaces inheriting from IUnknown have the first 3 entries in their vtable set to `QueryInterface`, `AddRef`, `Release` methods.
   - WinDbg or IDA can be also used as alternatives for extracting functions/methods VTable offsets.
 
### OPSEC Notes:
- Creating a COM object for `Word.Application` (or `Excel.Application` ..etc), will result spawning an additional `WinWord.exe` as a child process of svchost.exe instead of the main `WinWord.exe` process.
- `AccessVBOM` Registry key is modified/restored via COM using `WScript.Shell` , using win32 APIs could be a better alternative.
- Hosting the CLR using win32 APIs on VBA is obviously safer than updating the `AccessVBOM` registry key, will leave this for an other day...
- Other .NET APIs such as `System.CodeDom.Compiler` to compile/execute c# code from VBA, check reference below;

 
### References:

* https://github.com/jet2jet/vb2clr/blob/master/CLRHost.cls



