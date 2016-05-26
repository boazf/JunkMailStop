#include "stdafx.h"
#include "windows.h"
#include "atlbase.h"
#include "AddinMessageBox.h"
#include "Resource.h"

// Windows enumeration is done here in order to find outlook's main window in order to show
// the error dialig in front of it, in case an error dialog is neccessary.
struct EnumWindowsProcArgs {
    EnumWindowsProcArgs( DWORD p ) : pid( p ), hMainWnd(NULL) { }
    const DWORD pid;
    HWND hMainWnd;
};

static BOOL CALLBACK EnumWindowsProc(HWND hWnd, LPARAM lParam)
{
    EnumWindowsProcArgs *args = (EnumWindowsProcArgs *)lParam;
    DWORD windowPID;

    (void)::GetWindowThreadProcessId( hWnd, &windowPID );
    if ( windowPID == args->pid ) {
        if (GetWindow(hWnd, GW_OWNER) == (HWND)0 && IsWindowVisible(hWnd))
        {
            args->hMainWnd = hWnd;
            return FALSE;
        }
    }

    return TRUE;
}

static int AddinMessageBox(HWND hParentWnd, LPCWSTR szMessage, UINT uiType)
{
    CComBSTR bstrCaption;
    bstrCaption.LoadStringW(IDS_ADDIN_NAME);
    return MessageBox(hParentWnd, szMessage, bstrCaption.m_str, MB_OK | uiType);
}

int AddinMessageBox(LPCWSTR szMessage, UINT uiType)
{
    return AddinMessageBox(::GetActiveWindow(), szMessage, MB_OK | uiType);
}

int AddinMessageBox(UINT uiMessageId, UINT uiType)
{
    CComBSTR bstrMessage;
    bstrMessage.LoadStringW(uiMessageId);
    return AddinMessageBox(static_cast<LPCWSTR>(bstrMessage.m_str), uiType);
}

int AddinMessageBoxOnInit(LPCWSTR szMessage, UINT uiType)
{
    EnumWindowsProcArgs args(GetCurrentProcessId());
    EnumWindows(EnumWindowsProc, (LPARAM)&args);
    return AddinMessageBox(args.hMainWnd, szMessage, uiType);
}
