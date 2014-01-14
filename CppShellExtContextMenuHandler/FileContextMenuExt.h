
#pragma once

#include <windows.h>
#include <vector>
#include <shlobj.h>     // For IShellExtInit and IContextMenu
#include <time.h>
#include <fstream>  

using namespace std;

class FileContextMenuExt : public IShellExtInit, public IContextMenu
{
public:
    // IUnknown
    IFACEMETHODIMP QueryInterface(REFIID riid, void **ppv);
    IFACEMETHODIMP_(ULONG) AddRef();
    IFACEMETHODIMP_(ULONG) Release();

    // IShellExtInit
    IFACEMETHODIMP Initialize(LPCITEMIDLIST pidlFolder, LPDATAOBJECT pDataObj, HKEY hKeyProgID);

    // IContextMenu
    IFACEMETHODIMP QueryContextMenu(HMENU hMenu, UINT indexMenu, UINT idCmdFirst, UINT idCmdLast, UINT uFlags);
    IFACEMETHODIMP InvokeCommand(LPCMINVOKECOMMANDINFO pici);
    IFACEMETHODIMP GetCommandString(UINT_PTR idCommand, UINT uFlags, UINT *pwReserved, LPSTR pszName, UINT cchMax);
	
    FileContextMenuExt(void);

protected:
    ~FileContextMenuExt(void);

private:
    long m_cRef;


	struct FileAttributes {
		wchar_t fileName[MAX_PATH]; 
		DWORD	fileSize;
		time_t	fileDate;
		char temp[26];

	};

	vector<FileAttributes> m_szSelectedFiles;


    void OnVerbDisplayFileName(HWND hWnd);

    PWSTR m_pszMenuText;
    PCSTR m_pszVerb;
    PCWSTR m_pwszVerb;
    PCSTR m_pszVerbCanonicalName;
    PCWSTR m_pwszVerbCanonicalName;
    PCSTR m_pszVerbHelpText;
    PCWSTR m_pwszVerbHelpText;

	ofstream logFile;
};