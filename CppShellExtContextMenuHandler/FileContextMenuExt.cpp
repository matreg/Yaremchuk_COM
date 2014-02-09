#include "FileContextMenuExt.h"
#include "resource.h"
#include <strsafe.h>
#include <Shlwapi.h>
#include <fstream>
#include <thread>
#include <queue>

#pragma comment(lib, "shlwapi.lib")

extern HINSTANCE g_hInst;
extern long g_cDllRef;

#define IDM_DISPLAY             0  

std::string logFileAddress;

FileContextMenuExt::FileContextMenuExt(void) : m_cRef(1), 
    m_pszMenuText(L"&Calculate"),
    m_pszVerb("cppdisplay"),
    m_pwszVerb(L"cppdisplay"),
    m_pszVerbCanonicalName("CppDisplayFileName"),
    m_pwszVerbCanonicalName(L"CppDisplayFileName"),
    m_pszVerbHelpText("Calculate"),
    m_pwszVerbHelpText(L"Calculate")
{
    InterlockedIncrement(&g_cDllRef);

    m_hMenuBmp = LoadImage(g_hInst, MAKEINTRESOURCE(IDB_OK), 
        IMAGE_BITMAP, 0, 0, LR_DEFAULTSIZE | LR_LOADTRANSPARENT);
}

FileContextMenuExt::~FileContextMenuExt(void)
{
    if (m_hMenuBmp)
    {
        DeleteObject(m_hMenuBmp);
        m_hMenuBmp = NULL;
    }

    InterlockedDecrement(&g_cDllRef);
}

std::string FileNameWithoutPath(std::string fileName)
{
	std::string tempFileName;
	for ( size_t i = fileName.size() - 1 ; i >= 0 ; i-- )
		if (( fileName[i] != '\\' ) && ( fileName[i] != '/' ))
			tempFileName = fileName[i] + tempFileName;
		else break;
	return tempFileName;
}

void GetInformationAboutFileFromAddresses(std::multimap<std::string,FileInformation>::iterator it)
{
    WIN32_FILE_ATTRIBUTE_DATA wfad;
    SYSTEMTIME systime;
	GetFileAttributesExA(it->second.fileNameWithPath.c_str(), GetFileExInfoStandard, &wfad);
	FileTimeToSystemTime(&wfad.ftCreationTime, &systime);
	SystemTimeToTzSpecificLocalTime(NULL, &systime, &systime);
	it->second.fileDate = systime;
	it->second.fileSize = wfad.nFileSizeLow;

	DWORD sum = 0;
	std::ifstream file( it->second.fileNameWithPath.c_str(),  std::ios::binary );
	if ( file.is_open() )
	{
	   char byte;
	   while ( file.good() )
	   {
		  file.read(&byte, sizeof(char) );
		  sum += std::abs(byte);
	   }
	   file.close();
	}
	it->second.preByteSum = sum;
}

void UploadLogFile(std::multimap<std::string,FileInformation>::iterator it)
{
	std::ofstream logFile(logFileAddress.c_str(),std::ofstream::app);
	char outStrToFile[150];
	std::string dimensionFile = "B ";
	ULONG outFileSize = it->second.fileSize;
	if (outFileSize > 1000000)
	{
		outFileSize /= 1000000;
		dimensionFile = "MB";
	} else if (outFileSize > 1000)
	{
		outFileSize /= 1000;
		dimensionFile = "KB";
	}
	StringCchPrintfA(outStrToFile, ARRAYSIZE(outStrToFile), "%-70s %10d%2s %5d/%-2d/%-2d %2d:%2d:%-2d %15d",it->first.c_str(), outFileSize, dimensionFile.c_str(),
															it->second.fileDate.wYear, it->second.fileDate.wMonth, it->second.fileDate.wDay, 
															it->second.fileDate.wHour, it->second.fileDate.wMinute,it->second.fileDate.wSecond,
															it->second.preByteSum);
	logFile << outStrToFile << std::endl;
}

void StartThread(std::multimap<std::string,FileInformation>::iterator it)
{
	GetInformationAboutFileFromAddresses(it);
	UploadLogFile(it);
}
void FileContextMenuExt::OnVerbDisplayFileName()
{
	USHORT countThread = std::thread::hardware_concurrency();
	if (countThread > 1)
		countThread--;
	std::multimap<std::string,FileInformation>::iterator it = filesNames.begin();
	std::queue<std::thread> queuqeThread;
	while (it != filesNames.end())
	{
		if (!queuqeThread.empty())
		{
			if (!queuqeThread.front().joinable())
			{
				queuqeThread.pop();
				if (it != filesNames.end())
				{
					queuqeThread.push(std::thread(StartThread,it));
					queuqeThread.back().join();
					it++;
				}
			}
		} 
		else 
		{
			for (int i = 0; i < countThread; i++)
			{
				if (it != filesNames.end())
				{
					queuqeThread.push(std::thread(StartThread,it));
					queuqeThread.back().join();
					it++;
				}
				else break;
			}
		}
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

// Initialize the context menu handler.
IFACEMETHODIMP FileContextMenuExt::Initialize(LPCITEMIDLIST pidlFolder, LPDATAOBJECT pDataObj, HKEY hKeyProgID)
{
    if (NULL == pDataObj)
    {
        return E_INVALIDARG;
    }

    HRESULT hr = E_FAIL;

    FORMATETC fe = { CF_HDROP, NULL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL };
    STGMEDIUM stm;

	SYSTEMTIME lT;
	GetLocalTime(&lT);
	char logFileName[MAX_PATH];
	StringCchPrintfA(logFileName, ARRAYSIZE(logFileName), "C:\\LogFile_%d_%d_%d_%d_%d_%d.txt",lT.wYear, lT.wMonth,lT.wDay,lT.wHour,lT.wMinute,lT.wSecond);
	logFileAddress = logFileName;

    if (SUCCEEDED(pDataObj->GetData(&fe, &stm)))
    {
        HDROP hDrop = static_cast<HDROP>(GlobalLock(stm.hGlobal));
        if (hDrop != NULL)
        {
            UINT nFiles = DragQueryFileA(hDrop, 0xFFFFFFFF, NULL, 0);
			UINT countDirectory = 0;
			for (UINT i = 0; i < nFiles; i++)
			{
				char tempFileName[MAX_PATH];
				FileInformation tempFileInformation;

				if (0 != DragQueryFileA(hDrop, i, tempFileName, ARRAYSIZE(tempFileName)))
					if (FILE_ATTRIBUTE_DIRECTORY != GetFileAttributesA(tempFileName))
					{
						tempFileInformation.fileNameWithPath = tempFileName;
						filesNames.insert(std::pair<std::string,FileInformation>(FileNameWithoutPath(tempFileName),tempFileInformation));
					}
					else countDirectory++;
			}
		if (countDirectory != nFiles)
			hr = S_OK;
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
    // If uFlags include CMF_DEFAULTONLY then we should not do anything.
    if (CMF_DEFAULTONLY & uFlags)
    {
        return MAKE_HRESULT(SEVERITY_SUCCESS, 0, USHORT(0));
    }

    // Use either InsertMenu or InsertMenuItem to add menu items.
    // Learn how to add sub-menu from:
    // http://www.codeproject.com/KB/shell/ctxextsubmenu.aspx

    MENUITEMINFO mii = { sizeof(mii) };
    mii.fMask = MIIM_BITMAP | MIIM_STRING | MIIM_FTYPE | MIIM_ID | MIIM_STATE;
    mii.wID = idCmdFirst + IDM_DISPLAY;
    mii.fType = MFT_STRING;
    mii.dwTypeData = m_pszMenuText;
    mii.fState = MFS_ENABLED;
    mii.hbmpItem = static_cast<HBITMAP>(m_hMenuBmp);
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

    // Return an HRESULT value with the severity set to SEVERITY_SUCCESS. 
    // Set the code value to the offset of the largest command identifier 
    // that was assigned, plus one (1).
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

    // Determine which structure is being passed in, CMINVOKECOMMANDINFO or 
    // CMINVOKECOMMANDINFOEX based on the cbSize member of lpcmi. Although 
    // the lpcmi parameter is declared in Shlobj.h as a CMINVOKECOMMANDINFO 
    // structure, in practice it often points to a CMINVOKECOMMANDINFOEX 
    // structure. This struct is an extended version of CMINVOKECOMMANDINFO 
    // and has additional members that allow Unicode strings to be passed.
    if (pici->cbSize == sizeof(CMINVOKECOMMANDINFOEX))
    {
        if (pici->fMask & CMIC_MASK_UNICODE)
        {
            fUnicode = TRUE;
        }
    }

    // Determines whether the command is identified by its offset or verb.
    // There are two ways to identify commands:
    // 
    //   1) The command's verb string 
    //   2) The command's identifier offset
    // 
    // If the high-order word of lpcmi->lpVerb (for the ANSI case) or 
    // lpcmi->lpVerbW (for the Unicode case) is nonzero, lpVerb or lpVerbW 
    // holds a verb string. If the high-order word is zero, the command 
    // offset is in the low-order word of lpcmi->lpVerb.

    // For the ANSI case, if the high-order word is not zero, the command's 
    // verb string is in lpcmi->lpVerb. 
    if (!fUnicode && HIWORD(pici->lpVerb))
    {
        // Is the verb supported by this context menu extension?
        if (StrCmpIA(pici->lpVerb, m_pszVerb) == 0)
        {
            OnVerbDisplayFileName();
        }
        else
        {
            // If the verb is not recognized by the context menu handler, it 
            // must return E_FAIL to allow it to be passed on to the other 
            // context menu handlers that might implement that verb.
            return E_FAIL;
        }
    }

    // For the Unicode case, if the high-order word is not zero, the 
    // command's verb string is in lpcmi->lpVerbW. 
    else if (fUnicode && HIWORD(((CMINVOKECOMMANDINFOEX*)pici)->lpVerbW))
    {
        // Is the verb supported by this context menu extension?
        if (StrCmpIW(((CMINVOKECOMMANDINFOEX*)pici)->lpVerbW, m_pwszVerb) == 0)
        {
            OnVerbDisplayFileName();
        }
        else
        {
            // If the verb is not recognized by the context menu handler, it 
            // must return E_FAIL to allow it to be passed on to the other 
            // context menu handlers that might implement that verb.
            return E_FAIL;
        }
    }

    // If the command cannot be identified through the verb string, then 
    // check the identifier offset.
    else
    {
        // Is the command identifier offset supported by this context menu 
        // extension?
        if (LOWORD(pici->lpVerb) == IDM_DISPLAY)
        {
            OnVerbDisplayFileName();
        }
        else
        {
            // If the verb is not recognized by the context menu handler, it 
            // must return E_FAIL to allow it to be passed on to the other 
            // context menu handlers that might implement that verb.
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
            // Only useful for pre-Vista versions of Windows that have a 
            // Status bar.
            hr = StringCchCopy(reinterpret_cast<PWSTR>(pszName), cchMax, 
                m_pwszVerbHelpText);
            break;

        case GCS_VERBW:
            // GCS_VERBW is an optional feature that enables a caller to 
            // discover the canonical name for the verb passed in through 
            // idCommand.
            hr = StringCchCopy(reinterpret_cast<PWSTR>(pszName), cchMax, 
                m_pwszVerbCanonicalName);
            break;

        default:
            hr = S_OK;
        }
    }

    // If the command (idCommand) is not supported by this context menu 
    // extension handler, return E_INVALIDARG.

    return hr;
}

#pragma endregion