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
#include <windows.h>

#define WINDOWS TRUE;

#include "coreclrhost.h"
#include <system_error>


#define CORECLR_FILE_NAME "coreclr.dll"
#define FS_SEPARATOR "\\"
#define PATH_DELIMITER ";"


//CComBSTR ClrVersion;
void* hostHandle;
unsigned int domainId = 0;
HMODULE coreClr;
int appDomainCounter = 0;

typedef HRESULT(__stdcall* createWwDotnetBridgeHandler)(CComPtr<IDispatch>* ptr);
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


void module_handle_function() noexcept
{
}

static std::string ReplaceAll(std::string str, const std::string & from, const std::string & to) {
	size_t start_pos = 0;
	while ((start_pos = str.find(from, start_pos)) != std::string::npos) {
		str.replace(start_pos, from.length(), to);
		start_pos += to.length(); // Move past the replacement
	}
	return str;
}

std::string GetHostDirectory(HMODULE hModule)
{
	char path[MAX_PATH];
	GetModuleFileNameA(hModule, path, MAX_PATH);
	std::string dir(path);
	size_t pos = dir.rfind(FS_SEPARATOR);
	return (std::string::npos == pos) ? "" : dir.substr(0, pos);
}

BOOL CoreClrLoad(char* runtimePath, char* errorMessage, DWORD* size)
 {
	// The GET_MODULE_HANDLE_EX_FLAG_PIN flag seems to prevent VFP from unloading and clearing memory.	
	HMODULE hm = NULL;
	void* address = module_handle_function;
	GetModuleHandleEx(
		GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS | GET_MODULE_HANDLE_EX_FLAG_PIN,
		(LPCSTR)address,
		&hm
	);

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
		std::string desktopPath = ReplaceAll(runtimePath, "Microsoft.NETCore.App", "Microsoft.WindowsDesktop.App");
		BuildTpaList(desktopPath.c_str(), ".dll", tpaList); // Desktop must come first so that its assemblies take precedence over .NET Core assemblies when the same assembly is present in different forms in both runtimes (e.g. WindowsBase.dll).
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
		std::string hostDir = GetHostDirectory(hm);

		std::string appPaths(curDir);
		appPaths.append(PATH_DELIMITER);
		appPaths.append((const char *)curDir).append(FS_SEPARATOR).append("bin");
		if (hostDir != curDir)
		{
			appPaths.append(PATH_DELIMITER);
			appPaths.append(hostDir.c_str());
		}

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

			// this works!
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

		// this fails - HR Error: Not Supported
		CComPtr<IDispatch> obj;		
		hr = createWwDotnetBridge((CComPtr<IDispatch> *) &obj);
		VerifyHResult(hr);

		return obj;
	}
	catch (HResultException ex) {		
		ex.GetMessage(ErrorMessage);
		*dwErrorSize = strlen(ErrorMessage);
		return NULL;
	}
	catch(...) {		
		return NULL;
	}

	return NULL;
}
