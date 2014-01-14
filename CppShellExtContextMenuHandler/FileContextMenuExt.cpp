

#include "FileContextMenuExt.h"
#include "resource.h"
#include <strsafe.h>
#include <Shlwapi.h>
#include <time.h>
#pragma comment(lib, "shlwapi.lib")

#include <sys/stat.h>


extern HINSTANCE g_hInst;
extern long g_cDllRef;

#define IDM_DISPLAY             0  // The command's identifier offset

FileContextMenuExt::FileContextMenuExt(void) : m_cRef(1), 
    m_pszMenuText(L"&Write to log file"),
    m_pszVerb("writelog"),
    m_pwszVerb(L"writelog"),
    m_pszVerbCanonicalName("writelog"),
    m_pwszVerbCanonicalName(L"writelog"),
    m_pszVerbHelpText("Write log file"),
    m_pwszVerbHelpText(L"Write log file")
{
    InterlockedIncrement(&g_cDllRef);


	logFile.open("d:\\FileLog.log", ofstream::app );
}

FileContextMenuExt::~FileContextMenuExt(void)
{

    InterlockedDecrement(&g_cDllRef);
	logFile.close();
}


void FileContextMenuExt::OnVerbDisplayFileName(HWND hWnd)
{
    wchar_t szMessage[300];
	char date[26];

	for(UINT FileIterator = 0; 
		FileIterator < this->m_szSelectedFiles.size(); 
		FileIterator ++)
	{
		if(-1 == ctime_s(date, 26, 
			&m_szSelectedFiles[FileIterator].fileDate))
		{
			strcpy_s(date, "Date n/a");
		}

		char nameStr[MAX_PATH];
		wcstombs(nameStr, m_szSelectedFiles[FileIterator].fileName, MAX_PATH);
		logFile << nameStr << "\t" 
				<< date << "\t" 
				<< "Size " << m_szSelectedFiles[FileIterator].fileSize 
				<< " bytes.\n";
	}
    if (SUCCEEDED(StringCchPrintf(szMessage, ARRAYSIZE(szMessage), 
		L"%i records added to log file\r\n\r\n", this->m_szSelectedFiles.size() )))
    {
        MessageBox(hWnd, szMessage, L"CppShellExtContextMenuHandler", MB_OK);
    }
}


#pragma region IUnknown

// Query to the interface the component supported.
IFACEMETHODIMP FileContextMenuExt::QueryInterface(REFIID riid, void **ppv)
{
    static const QITAB qit[] = 
    {
        QITABENT(FileContextMenuExt, IContextMenu),
        QITABENT(FileContextMenuExt, IShellExtInit), 
        { 0 },
    };
    return QISearch(this, qit, riid, ppv);
}

// Increase the reference count for an interface on an object.
IFACEMETHODIMP_(ULONG) FileContextMenuExt::AddRef()
{
    return InterlockedIncrement(&m_cRef);
}

// Decrease the reference count for an interface on an object.
IFACEMETHODIMP_(ULONG) FileContextMenuExt::Release()
{
    ULONG cRef = InterlockedDecrement(&m_cRef);
    if (0 == cRef)
    {
        delete this;
    }

    return cRef;
}

#pragma endregion


#pragma region IShellExtInit

IFACEMETHODIMP FileContextMenuExt::Initialize(
    LPCITEMIDLIST pidlFolder, LPDATAOBJECT pDataObj, HKEY hKeyProgID)
{
    if (NULL == pDataObj)
    {
        return E_INVALIDARG;
    }

    HRESULT hr = E_FAIL;

    FORMATETC fe = { CF_HDROP, NULL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL };
    STGMEDIUM stm;

    if (SUCCEEDED(pDataObj->GetData(&fe, &stm)))
    {
        HDROP hDrop = static_cast<HDROP>(GlobalLock(stm.hGlobal));
        if (hDrop != NULL)
        {

            UINT nFiles = DragQueryFile(hDrop, 0xFFFFFFFF, NULL, 0);
			for(UINT FileIterator = 0; FileIterator < nFiles; FileIterator++)
            {
                // Get the path of the file.
				FileAttributes fileAttributes;
				wchar_t filePath[MAX_PATH];
				if (0 != DragQueryFile(hDrop, FileIterator, filePath, ARRAYSIZE(filePath)))
                {

					struct _stat buf;
					SHFILEINFOW sfi = {0};
					if( SUCCEEDED(
							SHGetFileInfo(filePath,
							-1,
							&sfi,
							sizeof(sfi),
							SHGFI_DISPLAYNAME)) 
						&& _wstat( filePath, &buf ) != -1
					)
					{
							wcscpy_s(fileAttributes.fileName, sfi.szDisplayName);
							fileAttributes.fileSize = buf.st_size;
							fileAttributes.fileDate = buf.st_ctime;
							m_szSelectedFiles.push_back(fileAttributes);
							hr = S_OK;
					}
                }
            }

            GlobalUnlock(stm.hGlobal);
        }

        ReleaseStgMedium(&stm);
    }

    return hr;
}

#pragma endregion


#pragma region IContextMenu

//
//   FUNCTION: FileContextMenuExt::QueryContextMenu
//
//   PURPOSE: The Shell calls IContextMenu::QueryContextMenu to allow the 
//            context menu handler to add its menu items to the menu. It 
//            passes in the HMENU handle in the hmenu parameter. The 
//            indexMenu parameter is set to the index to be used for the 
//            first menu item that is to be added.
//
IFACEMETHODIMP FileContextMenuExt::QueryContextMenu(
    HMENU hMenu, UINT indexMenu, UINT idCmdFirst, UINT idCmdLast, UINT uFlags)
{

    if (CMF_DEFAULTONLY & uFlags)
    {
        return MAKE_HRESULT(SEVERITY_SUCCESS, 0, USHORT(0));
    }


    MENUITEMINFO mii = { sizeof(mii) };
    mii.fMask = MIIM_BITMAP | MIIM_STRING | MIIM_FTYPE | MIIM_ID | MIIM_STATE;
    mii.wID = idCmdFirst + IDM_DISPLAY;
    mii.fType = MFT_STRING;
    mii.dwTypeData = m_pszMenuText;
    mii.fState = MFS_ENABLED;
    if (!InsertMenuItem(hMenu, indexMenu, TRUE, &mii))
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    // Add a separator.
    MENUITEMINFO sep = { sizeof(sep) };
    sep.fMask = MIIM_TYPE;
    sep.fType = MFT_SEPARATOR;
    if (!InsertMenuItem(hMenu, indexMenu + 1, TRUE, &sep))
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    return MAKE_HRESULT(SEVERITY_SUCCESS, 0, USHORT(IDM_DISPLAY + 1));
}


//
//   FUNCTION: FileContextMenuExt::InvokeCommand
//
//   PURPOSE: This method is called when a user clicks a menu item to tell 
//            the handler to run the associated command. The lpcmi parameter 
//            points to a structure that contains the needed information.
//
IFACEMETHODIMP FileContextMenuExt::InvokeCommand(LPCMINVOKECOMMANDINFO pici)
{
    BOOL fUnicode = FALSE;

    if (pici->cbSize == sizeof(CMINVOKECOMMANDINFOEX))
    {
        if (pici->fMask & CMIC_MASK_UNICODE)
        {
            fUnicode = TRUE;
        }
    }

 
    if (!fUnicode && HIWORD(pici->lpVerb))
    {
        // Is the verb supported by this context menu extension?
        if (StrCmpIA(pici->lpVerb, m_pszVerb) == 0)
        {
            OnVerbDisplayFileName(pici->hwnd);
        }
        else
        {

            return E_FAIL;
        }
    }

    else if (fUnicode && HIWORD(((CMINVOKECOMMANDINFOEX*)pici)->lpVerbW))
    {
        if (StrCmpIW(((CMINVOKECOMMANDINFOEX*)pici)->lpVerbW, m_pwszVerb) == 0)
        {
            OnVerbDisplayFileName(pici->hwnd);
        }
        else
        {
            return E_FAIL;
        }
    }

    else
    {
        if (LOWORD(pici->lpVerb) == IDM_DISPLAY)
        {
            OnVerbDisplayFileName(pici->hwnd);
        }
        else
        {
            return E_FAIL;
        }
    }

    return S_OK;
}


//
//   FUNCTION: CFileContextMenuExt::GetCommandString
//
//   PURPOSE: If a user highlights one of the items added by a context menu 
//            handler, the handler's IContextMenu::GetCommandString method is 
//            called to request a Help text string that will be displayed on 
//            the Windows Explorer status bar. This method can also be called 
//            to request the verb string that is assigned to a command. 
//            Either ANSI or Unicode verb strings can be requested. This 
//            example only implements support for the Unicode values of 
//            uFlags, because only those have been used in Windows Explorer 
//            since Windows 2000.
//
IFACEMETHODIMP FileContextMenuExt::GetCommandString(UINT_PTR idCommand, 
    UINT uFlags, UINT *pwReserved, LPSTR pszName, UINT cchMax)
{
    HRESULT hr = E_INVALIDARG;

    if (idCommand == IDM_DISPLAY)
    {
        switch (uFlags)
        {
        case GCS_HELPTEXTW:
            hr = StringCchCopy(reinterpret_cast<PWSTR>(pszName), cchMax, 
                m_pwszVerbHelpText);
            break;

        case GCS_VERBW:
            hr = StringCchCopy(reinterpret_cast<PWSTR>(pszName), cchMax, 
                m_pwszVerbCanonicalName);
            break;

        default:
            hr = S_OK;
        }
    }

    return hr;
}

#pragma endregion