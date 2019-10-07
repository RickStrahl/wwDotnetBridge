/*
		.NET Core Runtime Loader Helpers
		---------------------------------
		(c) Rick Strahl, West Wind Technologies, 2019

		This library provides the abillity to hoist a .NET Core Runtime instance
		into another process. It works fine for loading, but this approach
		does not support unloading and reloading.

		Any attempt to unload .NET Core works, but reloading will fail.

		There are two other APIs that are still under construction that
		might prove more flexible, but it is unclear whether they will
		support COM marshaling in the same way this approach does.
*/
#include <stdio.h>
#include <stdlib.h>
#include <string.h>

#include <string>

#include <comdef.h>
#include <atlbase.h>
#include <atlcomcli.h>
#include <MSCorEE.h>

#define WINDOWS TRUE;

#include "coreclrhost.h"


#define CORECLR_FILE_NAME "coreclr.dll"
#define FS_SEPARATOR "\\"
#define PATH_DELIMITER ";"


//CComBSTR ClrVersion;
void* hostHandle;
unsigned int domainId = 0;
HMODULE coreClr;
int appDomainCounter = 0;

typedef HRESULT(__stdcall* createWwDotnetBridgeHandler)(CComPtr<IUnknown>* ptr);
createWwDotnetBridgeHandler createWwDotnetBridge;

class HResultException
{
	char const* const error;
	HRESULT const hr;

public:
	HResultException(char const* error) : error(error), hr(0) {}
	HResultException(HRESULT hr) : error(nullptr), hr(hr) {}

	DWORD GetMessage(char* errorMessage)
	{
		auto outputSize = strlen(errorMessage);
		if (error) {
			strcpy_s(errorMessage, outputSize, error);
		}
		else {
			std::string message = std::system_category().message(hr);
			strcpy(errorMessage, message.c_str());
			//sprintf_s((char *)errorMessage, outputSize, "%ws", (LPWSTR)errorMessage);
		}
		return strlen(errorMessage);
	}
};
void VerifyHResult(HRESULT hr)
{
	if (FAILED(hr))
		throw HResultException(hr);
}
void VerifyHResult(char *msg)
{	
	throw HResultException(msg);
}

// Win32 directory search for .dll files
void BuildTpaList(const char* directory, const char* extension, std::string& tpaList)
{
	// This will add all files with a .dll extension to the TPA list.
	// This will include unmanaged assemblies (coreclr.dll, for example) that don't
	// belong on the TPA list. In a real host, only managed assemblies that the host
	// expects to load should be included. Having extra unmanaged assemblies doesn't
	// cause anything to fail, though, so this function just enumerates all dll's in
	// order to keep this sample concise.
	std::string searchPath(directory);
	searchPath.append(FS_SEPARATOR);
	searchPath.append("*");
	searchPath.append(extension);

	WIN32_FIND_DATAA findData;
	HANDLE fileHandle = FindFirstFileA(searchPath.c_str(), &findData);

	if (fileHandle != INVALID_HANDLE_VALUE)
	{
		do
		{
			// Append the assembly to the list
			tpaList.append(directory);			
			tpaList.append(findData.cFileName);
			tpaList.append(PATH_DELIMITER);

			// Note that the CLR does not guarantee which assembly will be loaded if an assembly
			// is in the TPA list multiple times (perhaps from different paths or perhaps with different NI/NI.dll
			// extensions. Therefore, a real host should probably add items to the list in priority order and only
			// add a file if it's not already present on the list.
			//
			// For this simple sample, though, and because we're only loading TPA assemblies from a single path,
			// and have no native images, we can ignore that complication.
		} while (FindNextFileA(fileHandle, &findData));
		FindClose(fileHandle);
	}
}




BOOL CoreClrLoad(char* runtimePath, char* errorMessage, DWORD* size)
 {
	if (!runtimePath)
	{
		strcpy(errorMessage, "Please pass in a path to a .NET Runtime version.");
		return FALSE; // strcpy(runtimePath, "c:\\program files (x86)\\dotnet\\shared\\Microsoft.NETCore.App\\3.0.0");
	}

	try {

		// Construct the CoreCLR path
		// For this sample, we know CoreCLR's path. For other hosts,
		// it may be necessary to probe for coreclr.dll/libcoreclr.so
		std::string coreClrPath(runtimePath);
		if (coreClrPath.at(coreClrPath.length()-1) != '\\')
			coreClrPath.append(FS_SEPARATOR);
		coreClrPath.append(CORECLR_FILE_NAME);

		// Construct the managed library path
		std::string managedLibraryPath(runtimePath);

		coreClr = LoadLibraryExA(coreClrPath.c_str(), NULL, 0);
		if (coreClr == NULL) {
			// TODO: need error information here
			VerifyHResult("Couldn't load CoreClr assembly.");
			return FALSE;
		}

		coreclr_initialize_ptr initializeCoreClr = (coreclr_initialize_ptr)GetProcAddress(coreClr, "coreclr_initialize");

		std::string tpaList;
		BuildTpaList(runtimePath, ".dll", tpaList);

		// <Snippet3>
		// Define CoreCLR properties
		// Other properties related to assembly loading are common here,
		// but for this simple sample, TRUSTED_PLATFORM_ASSEMBLIES is all
		// that is needed. Check hosting documentation for other common properties.
		const char* propertyKeys[] = {
			"TRUSTED_PLATFORM_ASSEMBLIES", // Trusted assemblies
			"APP_PATHS"  // Bin Paths
		};

		char curDir[MAX_PATH];
		GetCurrentDirectory(MAX_PATH, (LPSTR)curDir);

		std::string appPaths(curDir);

		appPaths.append(PATH_DELIMITER);
		appPaths.append((const char *)curDir);
		appPaths.append(FS_SEPARATOR);
		appPaths.append("bin");


		const char* propertyValues[] = {
			tpaList.c_str(),
			appPaths.c_str()
		};

		appDomainCounter++;
		char appDomainName[50];
		sprintf(appDomainName, "%s_%i", "wwDotNetBridge",appDomainCounter);

		// This function both starts the .NET Core runtime and creates
		// the default (and only) AppDomain
		HRESULT hr = initializeCoreClr(
			runtimePath,        // App base path
			appDomainName,
			sizeof(propertyKeys) / sizeof(char*),   // Property count
			propertyKeys,       // Property names
			propertyValues,     // Property values
			&hostHandle,        // Host handle
			&domainId);         // AppDomain ID


		VerifyHResult(hr);

		return TRUE;
	}
	catch (HResultException ex) {
		ex.GetMessageA(errorMessage);
		*size = strlen(errorMessage);
		return FALSE;
	}
}

// *** Unloads the CLR from the process
DWORD WINAPI CoreClrUnload()
{
	coreclr_shutdown_ptr shutdownCoreClr = (coreclr_shutdown_ptr)GetProcAddress(coreClr, "coreclr_shutdown");	
	HRESULT hr = shutdownCoreClr(hostHandle, domainId);
	VerifyHResult(hr);

	createWwDotnetBridge = NULL;
	hostHandle = NULL;
	domainId = 0;
	coreClr = NULL;
	HMODULE coreClr = 0; 
	
	
	return hr;
}



/// *** Creates an instance of a class from an assembly referenced through its disk path
IDispatch* WINAPI CoreClrCreateInstanceFrom(char *runtimePath, char *version, char *ErrorMessage, DWORD *dwErrorSize)
{
	if (domainId == 0)
		CoreClrLoad(runtimePath, ErrorMessage, dwErrorSize);

	try {
		HRESULT hr;

		// Create the function pointer based on signature and assembly version
		if (createWwDotnetBridge == NULL)
		{
			std::string assemblyName("wwDotNetBridge, Version=");
			assemblyName.append(version);
			assemblyName.append(", Culture=neutral, PublicKeyToken=null");

			coreclr_create_delegate_ptr createManagedDelegate = (coreclr_create_delegate_ptr)GetProcAddress(coreClr, "coreclr_create_delegate");
			hr = createManagedDelegate(
				hostHandle,
				domainId,
				assemblyName.c_str(),
				"Westwind.WebConnection.wwDotnetBridgeFactory",
				"CreatewwDotnetBridgeByRef",
				(VOID * *)& createWwDotnetBridge);
			VerifyHResult(hr);
		}
		
		// Now call the factory .NET method 
		CComPtr<IUnknown> obj;		
		hr = createWwDotnetBridge((CComPtr<IUnknown>*) & obj);
		VerifyHResult(hr);

		// Convert to dispatch object we can pass to FoxPro
		CComPtr<IDispatch> disp;
		hr = obj->QueryInterface(&disp);
		VerifyHResult(hr);
		
		return disp;
	}
	catch (HResultException ex) {		
		ex.GetMessage(ErrorMessage);
		*dwErrorSize = strlen(ErrorMessage);
		return NULL;
	}

	return NULL;
}
