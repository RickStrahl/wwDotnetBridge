#include <atlbase.h>
#include <mscoree.h>
#include <metahost.h> 
#include <string.h>
#include <string>
#include <system_error>

#import "mscorlib.tlb" raw_interfaces_only	no_smart_pointers high_property_prefixes("_get","_put","_putref") 

using namespace mscorlib;
//using namespace System;
//using namespace System::Security;
//using namespace System::Security::Policy;


CComPtr<ICorRuntimeHost>	spRuntimeHost = NULL;
CComPtr<_AppDomain>			spDefAppDomain = NULL;
CComPtr<_Object>             spDefObject = NULL;

// .NET Framework 4
CComPtr<ICLRMetaHost> pMetaHost;
CComPtr<ICLRRuntimeInfo> pRuntimeInfo;

CComBSTR ClrVersion;

DWORD WINAPI SetClrVersion(char *version)
{
	ClrVersion = version;
	return 1;
}

class HresultException
{
	char const* const error;
	HRESULT const hr;	

public:
	HresultException(char const* error) : error(error), hr(0) {}
	HresultException(HRESULT hr) : error(nullptr), hr(hr) {}

	DWORD GetMessage(char* errorMessage)
	{
		auto outputSize = strlen(errorMessage);
		if (error) {
			strcpy_s(errorMessage, outputSize, error);
		}
		else {
			std::string message = std::system_category().message(hr);
			strcpy(errorMessage, message.c_str());
			sprintf_s((char *)errorMessage, outputSize, "%ws", (LPWSTR)errorMessage);
		}
		return strlen(errorMessage);
	}
};

void VerifyHresult(HRESULT hr)
{
	if (FAILED(hr))
		throw HresultException(hr);
}

/// Starts up the CLR and creates a Default AppDomain
void WINAPI ClrLoadLegacyVersion2()
{
#pragma warning(push)
#pragma warning(disable: 4996) // CorBindToRuntimeEx is deprecated, but we need to use it to start CLR version 2.
	//Retrieve a pointer to the ICorRuntimeHost interface
	VerifyHresult(CorBindToRuntimeEx(
		ClrVersion,	//Retrieve latest version by default
		L"wks",	//Request a WorkStation build of the CLR
		STARTUP_LOADER_OPTIMIZATION_MULTI_DOMAIN | STARTUP_CONCURRENT_GC,
		CLSID_CorRuntimeHost,
		IID_ICorRuntimeHost,
		(void**)&spRuntimeHost
	));
#pragma warning(pop)

	//Start the CLR
	VerifyHresult(spRuntimeHost->Start());

	CComPtr<IUnknown> pUnk;
	WCHAR domainId[50];
	swprintf(domainId, 50, L"%s_%i", L"wwDotNetBridge", GetTickCount());
	VerifyHresult(spRuntimeHost->CreateDomain(domainId, NULL, &pUnk));
	VerifyHresult(pUnk->QueryInterface(&spDefAppDomain.p));
}

void WINAPI ClrLoad()
{
	if (ClrVersion.Length() >= 2 && ClrVersion[0] == 'v' && ClrVersion[1] == '2') {
		ClrLoadLegacyVersion2();
		return;
	}

	// Tutorial on how this works: https://code.msdn.microsoft.com/windowsdesktop/CppHostCLR-e6581ee0
	VerifyHresult(CLRCreateInstance(CLSID_CLRMetaHost, IID_PPV_ARGS(&pMetaHost)));
	VerifyHresult(pMetaHost->GetRuntime(ClrVersion, IID_PPV_ARGS(&pRuntimeInfo)));
	BOOL fLoadable;
	VerifyHresult(pRuntimeInfo->IsLoadable(&fLoadable));
	if (!fLoadable) throw HresultException("CLR is not loadable.");
	VerifyHresult(pRuntimeInfo->GetInterface(CLSID_CorRuntimeHost, IID_PPV_ARGS(&spRuntimeHost)));
	VerifyHresult(spRuntimeHost->Start());
	CComPtr<IUnknown> spAppDomainThunk;
	VerifyHresult(spRuntimeHost->GetDefaultDomain(&spAppDomainThunk));
	VerifyHresult(spAppDomainThunk->QueryInterface(IID_PPV_ARGS(&spDefAppDomain)));
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
IDispatch* WINAPI ClrCreateInstance(char *AssemblyName, char *className, char *ErrorMessage, DWORD *dwErrorSize)
{
	try {
		if (!spDefAppDomain)
			ClrLoad();

		CComPtr<_ObjectHandle> spObjectHandle;
		VerifyHresult(spDefAppDomain->CreateInstance(_bstr_t(AssemblyName), _bstr_t(className), &spObjectHandle));

		CComVariant VntUnwrapped;
		
		VerifyHresult(spObjectHandle->Unwrap(&VntUnwrapped));
		return VntUnwrapped.pdispVal;
	}
	catch (HresultException ex) {
		*dwErrorSize = ex.GetMessage(ErrorMessage);
		return NULL;
	}
}

/// *** Creates an instance of a class from an assembly referenced through its disk path
IDispatch* WINAPI ClrCreateInstanceFrom(char *AssemblyFileName, char *className, char *ErrorMessage, DWORD *dwErrorSize)
{
	try {
		if (!spDefAppDomain)
			ClrLoad();

		CComPtr<_ObjectHandle> spObjectHandle;
		VerifyHresult(spDefAppDomain->CreateInstanceFrom(_bstr_t(AssemblyFileName), _bstr_t(className), &spObjectHandle));

		CComVariant VntUnwrapped;
		VerifyHresult(spObjectHandle->Unwrap(&VntUnwrapped));

		return VntUnwrapped.pdispVal;
	}
	catch (HresultException ex) {
		*dwErrorSize = ex.GetMessage(ErrorMessage);
		return NULL;
	}
}
