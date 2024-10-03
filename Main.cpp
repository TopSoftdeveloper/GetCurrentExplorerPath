// Test.cpp : This file contains the 'main' function. Program execution begins and ends there.
//

#include <iostream>
#include <windows.h>

#include <Windows.h>
#include <shlobj.h>
#include <atlcomcli.h>  // for COM smart pointers
#include <atlbase.h>    // for COM smart pointers
#include <vector>
#include <system_error>
#include <memory>
#include <iostream>

// Throw a std::system_error if the HRESULT indicates failure.
template< typename T >
void ThrowIfFailed(HRESULT hr, T&& msg)
{
	if (FAILED(hr))
		throw std::system_error{ hr, std::system_category(), std::forward<T>(msg) };
}

// Deleter for a PIDL allocated by the shell.
struct CoTaskMemDeleter
{
	void operator()(ITEMIDLIST* pidl) const { ::CoTaskMemFree(pidl); }
};
// A smart pointer for PIDLs.
using UniquePidlPtr = std::unique_ptr< ITEMIDLIST, CoTaskMemDeleter >;

// Return value of GetCurrentExplorerFolders()
struct ExplorerFolderInfo
{
	HWND hwnd = nullptr;  // window handle of explorer
	UniquePidlPtr pidl;   // PIDL that points to current folder
};

// Get information about all currently open explorer windows.
// Throws std::system_error exception to report errors.
std::vector< ExplorerFolderInfo > GetCurrentExplorerFolders()
{
	CComPtr< IShellWindows > pshWindows;
	ThrowIfFailed(
		pshWindows.CoCreateInstance(CLSID_ShellWindows),
		"Could not create instance of IShellWindows");

	long count = 0;
	ThrowIfFailed(
		pshWindows->get_Count(&count),
		"Could not get number of shell windows");

	std::vector< ExplorerFolderInfo > result;
	result.reserve(count);

	for (long i = 0; i < count; ++i)
	{
		ExplorerFolderInfo info;

		CComVariant vi{ i };
		CComPtr< IDispatch > pDisp;
		ThrowIfFailed(
			pshWindows->Item(vi, &pDisp),
			"Could not get item from IShellWindows");

		if (!pDisp)
			// Skip - this shell window was registered with a NULL IDispatch
			continue;

		CComQIPtr< IWebBrowserApp > pApp{ pDisp };
		if (!pApp)
			// This window doesn't implement IWebBrowserApp 
			continue;

		// Get the window handle.
		pApp->get_HWND(reinterpret_cast<SHANDLE_PTR*>(&info.hwnd));

		CComQIPtr< IServiceProvider > psp{ pApp };
		if (!psp)
			// This window doesn't implement IServiceProvider
			continue;

		CComPtr< IShellBrowser > pBrowser;
		if (FAILED(psp->QueryService(SID_STopLevelBrowser, &pBrowser)))
			// This window doesn't provide IShellBrowser
			continue;

		CComPtr< IShellView > pShellView;
		if (FAILED(pBrowser->QueryActiveShellView(&pShellView)))
			// For some reason there is no active shell view
			continue;

		CComQIPtr< IFolderView > pFolderView{ pShellView };
		if (!pFolderView)
			// The shell view doesn't implement IFolderView
			continue;

		// Get the interface from which we can finally query the PIDL of
		// the current folder.
		CComPtr< IPersistFolder2 > pFolder;
		if (FAILED(pFolderView->GetFolder(IID_IPersistFolder2, (void**)& pFolder)))
			continue;

		LPITEMIDLIST pidl = nullptr;
		if (SUCCEEDED(pFolder->GetCurFolder(&pidl)))
		{
			// Take ownership of the PIDL via std::unique_ptr.
			info.pidl = UniquePidlPtr{ pidl };
			result.push_back(std::move(info));
		}
	}

	return result;
}

int main()
{
	::CoInitialize(nullptr);

	try
	{
		std::wcout << L"Currently open explorer windows:\n";
		for (const auto& info : GetCurrentExplorerFolders())
		{
			CComHeapPtr<wchar_t> pPath;
			if (SUCCEEDED(::SHGetNameFromIDList(info.pidl.get(), SIGDN_URL, &pPath)))
			{
				std::wcout << L"hwnd: 0x" << std::hex << info.hwnd
					<< L", path: " << static_cast<LPWSTR>(pPath) << L"\n";
			}
		}
	}
	catch (std::system_error& e)
	{
		std::cout << "ERROR: " << e.what() << "\nError code: " << e.code() << "\n";
	}

	::CoUninitialize();
}
