#pragma once

#include <windows.h>
#include <shlobj.h>   
#include <string>
#include <map>

struct FileInformation
{
	DWORD preByteSum;
	SYSTEMTIME fileDate;
	ULONG fileSize;
	std::string fileNameWithPath;
};

class FileContextMenuExt : public IShellExtInit, public IContextMenu
{
public:

    IFACEMETHODIMP QueryInterface(REFIID riid, void **ppv);
    IFACEMETHODIMP_(ULONG) AddRef();
    IFACEMETHODIMP_(ULONG) Release();

    IFACEMETHODIMP Initialize(LPCITEMIDLIST pidlFolder, LPDATAOBJECT pDataObj, HKEY hKeyProgID);

    IFACEMETHODIMP QueryContextMenu(HMENU hMenu, UINT indexMenu, UINT idCmdFirst, UINT idCmdLast, UINT uFlags);
    IFACEMETHODIMP InvokeCommand(LPCMINVOKECOMMANDINFO pici);
    IFACEMETHODIMP GetCommandString(UINT_PTR idCommand, UINT uFlags, UINT *pwReserved, LPSTR pszName, UINT cchMax);
	
	std::multimap<std::string,FileInformation> filesNames;

    FileContextMenuExt(void);

protected:
    ~FileContextMenuExt(void);

private:
    // Reference count of component.
    long m_cRef;

    // The method that handles the "display" verb.
    void OnVerbDisplayFileName();

    PWSTR m_pszMenuText;
    HANDLE m_hMenuBmp;
    PCSTR m_pszVerb;
    PCWSTR m_pwszVerb;
    PCSTR m_pszVerbCanonicalName;
    PCWSTR m_pwszVerbCanonicalName;
    PCSTR m_pszVerbHelpText;
    PCWSTR m_pwszVerbHelpText;
};