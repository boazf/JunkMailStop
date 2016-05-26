#include "StdAfx.h"
#include "ServersDetection.h"
#include "debug.h"

DWORD WINAPI ProgressDialogThreadProc(LPVOID lpParameter)
{
    DEBUG_TRACE(TRACE_VERBOSE_INFO, (L"Created thread for the servers detection progress dialog.\n"));
    ServersDetectionProgressDlg *pProgressDlg = (ServersDetectionProgressDlg *)lpParameter;
    pProgressDlg->DoModal();
    DEBUG_TRACE(TRACE_VERBOSE_INFO, (L"Exiting thread for the servers detection progress dialog.\n"));
    return 0;
}

