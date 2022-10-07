interface _AppDomain : IUnknown {
    HRESULT _stdcall GetTypeInfoCount([out] unsigned long* pcTInfo);

    HRESULT _stdcall GetTypeInfo(
      [in] unsigned long iTInfo, 
      [in] unsigned long lcid, 
      [in] long ppTInfo
    );

    HRESULT _stdcall GetIDsOfNames(
      [in] GUID* riid, 
      [in] long rgszNames, 
      [in] unsigned long cNames, 
      [in] unsigned long lcid, 
      [in] long rgDispId
    );

    HRESULT _stdcall Invoke(
      [in] unsigned long dispIdMember, 
      [in] GUID* riid, 
      [in] unsigned long lcid, 
      [in] short wFlags, 
      [in] long pDispParams, 
      [in] long pVarResult, 
      [in] long pExcepInfo, 
      [in] long puArgErr
    );

    [propget,custom(54FC8F55-38DE-4703-9C4E-250351302B1C, 1)]
    HRESULT _stdcall ToString([out, retval] BSTR* pRetVal);

    HRESULT _stdcall Equals(
      [in] VARIANT other, 
      [out, retval] VARIANT_BOOL* pRetVal
    );

    HRESULT _stdcall GetHashCode([out, retval] long* pRetVal);

    HRESULT _stdcall GetType([out, retval] _Type** pRetVal);

    HRESULT _stdcall InitializeLifetimeService([out, retval] VARIANT* pRetVal);

    HRESULT _stdcall GetLifetimeService([out, retval] VARIANT* pRetVal);

    [propget]
    HRESULT _stdcall Evidence([out, retval] _Evidence** pRetVal);

    HRESULT _stdcall add_DomainUnload([in] _EventHandler* value);

    HRESULT _stdcall remove_DomainUnload([in] _EventHandler* value);

    HRESULT _stdcall add_AssemblyLoad([in] _AssemblyLoadEventHandler* value);

    HRESULT _stdcall remove_AssemblyLoad([in] _AssemblyLoadEventHandler* value);

    HRESULT _stdcall add_ProcessExit([in] _EventHandler* value);

    HRESULT _stdcall remove_ProcessExit([in] _EventHandler* value);

    HRESULT _stdcall add_TypeResolve([in] _ResolveEventHandler* value);

    HRESULT _stdcall remove_TypeResolve([in] _ResolveEventHandler* value);

    HRESULT _stdcall add_ResourceResolve([in] _ResolveEventHandler* value);

    HRESULT _stdcall remove_ResourceResolve([in] _ResolveEventHandler* value);

    HRESULT _stdcall add_AssemblyResolve([in] _ResolveEventHandler* value);

    HRESULT _stdcall remove_AssemblyResolve([in] _ResolveEventHandler* value);

    HRESULT _stdcall add_UnhandledException([in] _UnhandledExceptionEventHandler* value);

    HRESULT _stdcall remove_UnhandledException([in] _UnhandledExceptionEventHandler* value);

    HRESULT _stdcall DefineDynamicAssembly(
      [in] _AssemblyName* name, 
      [in] AssemblyBuilderAccess access, 
      [out, retval] _AssemblyBuilder** pRetVal
    );

    custom(0F21F359-AB84-41E8-9A78-36D110E6D2F9, "DefineDynamicAssembly")
    HRESULT _stdcall DefineDynamicAssembly_2(
      [in] _AssemblyName* name, 
      [in] AssemblyBuilderAccess access, 
      [in] BSTR dir, 
      [out, retval] _AssemblyBuilder** pRetVal
    );
    
    custom(0F21F359-AB84-41E8-9A78-36D110E6D2F9, "DefineDynamicAssembly")
    HRESULT _stdcall DefineDynamicAssembly_3(
      [in] _AssemblyName* name, 
      [in] AssemblyBuilderAccess access, 
      [in] _Evidence* Evidence, 
      [out, retval] _AssemblyBuilder** pRetVal
    );

    custom(0F21F359-AB84-41E8-9A78-36D110E6D2F9, "DefineDynamicAssembly")
    HRESULT _stdcall DefineDynamicAssembly_4(
      [in] _AssemblyName* name, 
      [in] AssemblyBuilderAccess access, 
      [in] _PermissionSet* requiredPermissions, 
      [in] _PermissionSet* optionalPermissions, 
      [in] _PermissionSet* refusedPermissions, 
      [out, retval] _AssemblyBuilder** pRetVal
    );

    custom(0F21F359-AB84-41E8-9A78-36D110E6D2F9, "DefineDynamicAssembly")
    HRESULT _stdcall DefineDynamicAssembly_5(
      [in] _AssemblyName* name, 
      [in] AssemblyBuilderAccess access, 
      [in] BSTR dir, 
      [in] _Evidence* Evidence, 
      [out, retval] _AssemblyBuilder** pRetVal
    );

    custom(0F21F359-AB84-41E8-9A78-36D110E6D2F9, "DefineDynamicAssembly")
    HRESULT _stdcall DefineDynamicAssembly_6(
      [in] _AssemblyName* name, 
      [in] AssemblyBuilderAccess access, 
      [in] BSTR dir, 
      [in] _PermissionSet* requiredPermissions, 
      [in] _PermissionSet* optionalPermissions, 
      [in] _PermissionSet* refusedPermissions, 
      [out, retval] _AssemblyBuilder** pRetVal
    );
    
    custom(0F21F359-AB84-41E8-9A78-36D110E6D2F9, "DefineDynamicAssembly")
    HRESULT _stdcall DefineDynamicAssembly_7(
      [in] _AssemblyName* name, 
      [in] AssemblyBuilderAccess access, 
      [in] _Evidence* Evidence, 
      [in] _PermissionSet* requiredPermissions, 
      [in] _PermissionSet* optionalPermissions, 
      [in] _PermissionSet* refusedPermissions, 
      [out, retval] _AssemblyBuilder** pRetVal
    );
    
    custom(0F21F359-AB84-41E8-9A78-36D110E6D2F9, "DefineDynamicAssembly")
    HRESULT _stdcall DefineDynamicAssembly_8(
      [in] _AssemblyName* name, 
      [in] AssemblyBuilderAccess access, 
      [in] BSTR dir, 
      [in] _Evidence* Evidence, 
      [in] _PermissionSet* requiredPermissions, 
      [in] _PermissionSet* optionalPermissions, 
      [in] _PermissionSet* refusedPermissions, 
      [out, retval] _AssemblyBuilder** pRetVal
    );
    
    custom(0F21F359-AB84-41E8-9A78-36D110E6D2F9, "DefineDynamicAssembly")
    HRESULT _stdcall DefineDynamicAssembly_9(
      [in] _AssemblyName* name, 
      [in] AssemblyBuilderAccess access, 
      [in] BSTR dir, 
      [in] _Evidence* Evidence, 
      [in] _PermissionSet* requiredPermissions, 
      [in] _PermissionSet* optionalPermissions, 
      [in] _PermissionSet* refusedPermissions, 
      [in] VARIANT_BOOL IsSynchronized, 
      [out, retval] _AssemblyBuilder** pRetVal
    );

    HRESULT _stdcall CreateInstance(
      [in] BSTR AssemblyName, 
      [in] BSTR typeName, 
      [out, retval] _ObjectHandle** pRetVal
    );

    HRESULT _stdcall CreateInstanceFrom(
      [in] BSTR assemblyFile, 
      [in] BSTR typeName, 
      [out, retval] _ObjectHandle** pRetVal
    );
    
    custom(0F21F359-AB84-41E8-9A78-36D110E6D2F9, "CreateInstance")
    HRESULT _stdcall CreateInstance_2(
      [in] BSTR AssemblyName, 
      [in] BSTR typeName, 
      [in] SAFEARRAY(VARIANT) activationAttributes, 
      [out, retval] _ObjectHandle** pRetVal
    );
    
    custom(0F21F359-AB84-41E8-9A78-36D110E6D2F9, "CreateInstanceFrom")
    HRESULT _stdcall CreateInstanceFrom_2(
      [in] BSTR assemblyFile, 
      [in] BSTR typeName, 
      [in] SAFEARRAY(VARIANT) activationAttributes, 
      [out, retval] _ObjectHandle** pRetVal
    );
    
    custom(0F21F359-AB84-41E8-9A78-36D110E6D2F9, "CreateInstance")
    HRESULT _stdcall CreateInstance_3(
      [in] BSTR AssemblyName, 
      [in] BSTR typeName, 
      [in] VARIANT_BOOL ignoreCase, 
      [in] BindingFlags bindingAttr, 
      [in] _Binder* Binder, 
      [in] SAFEARRAY(VARIANT) args, 
      [in] _CultureInfo* culture, 
      [in] SAFEARRAY(VARIANT) activationAttributes, 
      [in] _Evidence* securityAttributes, 
      [out, retval] _ObjectHandle** pRetVal
    );
    
    custom(0F21F359-AB84-41E8-9A78-36D110E6D2F9, "CreateInstanceFrom")
    HRESULT _stdcall CreateInstanceFrom_3(
      [in] BSTR assemblyFile, 
      [in] BSTR typeName, 
      [in] VARIANT_BOOL ignoreCase, 
      [in] BindingFlags bindingAttr, 
      [in] _Binder* Binder, 
      [in] SAFEARRAY(VARIANT) args, 
      [in] _CultureInfo* culture, 
      [in] SAFEARRAY(VARIANT) activationAttributes, 
      [in] _Evidence* securityAttributes, 
      [out, retval] _ObjectHandle** pRetVal
    );


    HRESULT _stdcall Load(
      [in] _AssemblyName* assemblyRef, 
      [out, retval] _Assembly** pRetVal
    );
    
    custom(0F21F359-AB84-41E8-9A78-36D110E6D2F9, "Load")
    HRESULT _stdcall Load_2(
      [in] BSTR assemblyString, 
      [out, retval] _Assembly** pRetVal
    );
    
    custom(0F21F359-AB84-41E8-9A78-36D110E6D2F9, "Load")
    HRESULT _stdcall Load_3(
      [in] SAFEARRAY(unsigned char) rawAssembly, 
      [out, retval] _Assembly** pRetVal
    );
    
    custom(0F21F359-AB84-41E8-9A78-36D110E6D2F9, "Load")
    HRESULT _stdcall Load_4(
      [in] SAFEARRAY(unsigned char) rawAssembly, 
      [in] SAFEARRAY(unsigned char) rawSymbolStore, 
      [out, retval] _Assembly** pRetVal
    );
    
    custom(0F21F359-AB84-41E8-9A78-36D110E6D2F9, "Load")
    HRESULT _stdcall Load_5(
      [in] SAFEARRAY(unsigned char) rawAssembly, 
      [in] SAFEARRAY(unsigned char) rawSymbolStore, 
      [in] _Evidence* securityEvidence, 
      [out, retval] _Assembly** pRetVal
    );
    
    custom(0F21F359-AB84-41E8-9A78-36D110E6D2F9, "Load")
    HRESULT _stdcall Load_6(
      [in] _AssemblyName* assemblyRef, 
      [in] _Evidence* assemblySecurity, 
      [out, retval] _Assembly** pRetVal
    );
    
    custom(0F21F359-AB84-41E8-9A78-36D110E6D2F9, "Load")
    HRESULT _stdcall Load_7(
      [in] BSTR assemblyString, 
      [in] _Evidence* assemblySecurity, 
      [out, retval] _Assembly** pRetVal
    );

    HRESULT _stdcall ExecuteAssembly(
      [in] BSTR assemblyFile, 
      [in] _Evidence* assemblySecurity, 
      [out, retval] long* pRetVal
    );

    custom(0F21F359-AB84-41E8-9A78-36D110E6D2F9, "ExecuteAssembly")
    HRESULT _stdcall ExecuteAssembly_2(
      [in] BSTR assemblyFile, 
      [out, retval] long* pRetVal
    );

    custom(0F21F359-AB84-41E8-9A78-36D110E6D2F9, "ExecuteAssembly")
    HRESULT _stdcall ExecuteAssembly_3(
      [in] BSTR assemblyFile, 
      [in] _Evidence* assemblySecurity, 
      [in] SAFEARRAY(BSTR) args, 
      [out, retval] long* pRetVal
    );

    [propget]
    HRESULT _stdcall FriendlyName([out, retval] BSTR* pRetVal);
    
    [propget]
    HRESULT _stdcall BaseDirectory([out, retval] BSTR* pRetVal);
    
    [propget]
    HRESULT _stdcall RelativeSearchPath([out, retval] BSTR* pRetVal);
    
    [propget]
    HRESULT _stdcall ShadowCopyFiles([out, retval] VARIANT_BOOL* pRetVal);

    HRESULT _stdcall GetAssemblies([out, retval] SAFEARRAY(_Assembly*)* pRetVal);

    HRESULT _stdcall AppendPrivatePath([in] BSTR Path);

    HRESULT _stdcall ClearPrivatePath();

    HRESULT _stdcall SetShadowCopyPath([in] BSTR s);

    HRESULT _stdcall ClearShadowCopyPath();

    HRESULT _stdcall SetCachePath([in] BSTR s);

    HRESULT _stdcall SetData(
      [in] BSTR name, 
      [in] VARIANT data
    );

    HRESULT _stdcall GetData(
      [in] BSTR name, 
      [out, retval] VARIANT* pRetVal
    );
    
    HRESULT _stdcall SetAppDomainPolicy([in] _PolicyLevel* domainPolicy);
    
    HRESULT _stdcall SetThreadPrincipal([in] IPrincipal* principal);
    
    HRESULT _stdcall SetPrincipalPolicy([in] PrincipalPolicy policy);
    
    HRESULT _stdcall DoCallBack([in] _CrossAppDomainDelegate* theDelegate);
    
    [propget]
    HRESULT _stdcall DynamicDirectory([out, retval] BSTR* pRetVal);
};
