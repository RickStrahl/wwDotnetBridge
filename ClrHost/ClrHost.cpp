#include <atlbase.h>
#include <mscoree.h>

#import "mscorlib.tlb" raw_interfaces_only	no_smart_pointers high_property_prefixes("_get","_put","_putref") 

using namespace mscorlib;
//using namespace System;
//using namespace System::Security;
//using namespace System::Security::Policy;


CComPtr<ICorRuntimeHost>	spRuntimeHost = NULL;
CComPtr<_AppDomain>			spDefAppDomain = NULL;
CComPtr<_Object>             spDefObject = NULL;

BSTR ClrVersion = NULL;

DWORD WINAPI SetClrVersion(char *version)
{
	ClrVersion = CComBSTR(version);
	return 1;
}

/// Assigns an error message
DWORD WINAPI SetError(HRESULT hr, char *ErrorMessage)
{
	if (ErrorMessage)
	{ 			
		int len= strlen(ErrorMessage);		
		LoadStringRC(hr & 0xffff,(LPWSTR)ErrorMessage,len/2,0);																
		sprintf((char *)ErrorMessage,"%ws",ErrorMessage);
				
		return strlen(ErrorMessage);
	}
	 
	strcpy(ErrorMessage,"");
	return 0;
}

/// Starts up the CLR and creates a Default AppDomain
DWORD WINAPI ClrLoad(char *ErrorMessage, DWORD *dwErrorSize)
{
	if (spDefAppDomain)
		return 1;

	
	//Retrieve a pointer to the ICorRuntimeHost interface
	HRESULT hr = CorBindToRuntimeEx(
					ClrVersion,	//Retrieve latest version by default
					L"wks",	//Request a WorkStation build of the CLR
					STARTUP_LOADER_OPTIMIZATION_MULTI_DOMAIN | STARTUP_CONCURRENT_GC, 
					CLSID_CorRuntimeHost,
					IID_ICorRuntimeHost,
					(void**)&spRuntimeHost
					);

	if (FAILED(hr)) 
	{
		*dwErrorSize = SetError(hr,ErrorMessage);	
		return hr;
	}

	//Start the CLR
	hr = spRuntimeHost->Start();

	if (FAILED(hr))
		return hr;

	CComPtr<IUnknown> pUnk;

	//Retrieve the IUnknown default AppDomain
	//hr = spRuntimeHost->GetDefaultDomain(&pUnk);
	//if (FAILED(hr)) 
	//	return hr;

	
	WCHAR domainId[50];
	swprintf(domainId,L"%s_%i",L"wwDotNetBridge",GetTickCount());
	hr = spRuntimeHost->CreateDomain(domainId,NULL,&pUnk);	

	//spRuntimeHost->CreateDomainSetup(&pUnk);
	//

	hr = pUnk->QueryInterface(&spDefAppDomain.p);
	if (FAILED(hr)) 
		return hr;
	
	return 1;
}



// *** Unloads the CLR from the process
DWORD WINAPI ClrUnload()
{
	if (spDefAppDomain)
	{
		spRuntimeHost->UnloadDomain(spDefAppDomain.p);
		spDefAppDomain.Release();
		spDefAppDomain = NULL;

		spRuntimeHost->Stop();		
		spRuntimeHost.Release();
		spRuntimeHost = NULL;
	}
	
	return 1;
}


// *** Creates an instance by Name (ie. local path assembly without extension or GAC'd FullName of 
//      any signed assemblies.
DWORD WINAPI ClrCreateInstance( char *AssemblyName, char *className, char *ErrorMessage, DWORD *dwErrorSize)
{
	CComPtr<_ObjectHandle> spObjectHandle;
	
	if (!spDefAppDomain)
	{		
		if (ClrLoad(ErrorMessage,dwErrorSize) != 1)
			return -1;
	}

	DWORD hr;

	//Creates an instance of the type specified in the Assembly
	hr = spDefAppDomain->CreateInstance(		
										_bstr_t(AssemblyName), 
										_bstr_t(className),
										&spObjectHandle
	);


	
	*dwErrorSize = 0;

	if (FAILED(hr)) 
	{		
		*dwErrorSize = SetError(hr,ErrorMessage);
		return -1;
	}

	CComVariant VntUnwrapped;
	hr = spObjectHandle->Unwrap(&VntUnwrapped);
	if (FAILED(hr))
		return -1;


	CComPtr<IDispatch> pDisp;
	pDisp = VntUnwrapped.pdispVal;	
	
	return (DWORD) pDisp.p;
}

/// *** Creates an instance of a class from an assembly referenced through it's disk path
DWORD WINAPI ClrCreateInstanceFrom( char *AssemblyFileName, char *className, char *ErrorMessage, DWORD *dwErrorSize)
{
	CComPtr<_ObjectHandle> spObjectHandle;
	
	if (!spDefAppDomain)
	{		
		if (ClrLoad(ErrorMessage,dwErrorSize) != 1)
			return -1;
	}

	DWORD hr;

	//Creates an instance of the type specified in the Assembly
	hr = spDefAppDomain->CreateInstanceFrom(		
									_bstr_t(AssemblyFileName), 
									_bstr_t(className),
									&spObjectHandle
	);	

	*dwErrorSize = 0;

	if (FAILED(hr)) 
	{		
		*dwErrorSize = SetError(hr,ErrorMessage);
		return -1;
	}

	CComVariant VntUnwrapped;
	hr = spObjectHandle->Unwrap(&VntUnwrapped);
	if (FAILED(hr))
		return -1;

	CComPtr<IDispatch> pDisp;
	pDisp = VntUnwrapped.pdispVal;
	
	// *** pass the raw COM pointer back
	return (DWORD) pDisp.p;
}

