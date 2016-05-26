#include "StdAfx.h"
#include "ServersDetectionProgressDlg.h"
#include "debug.h"
#include "ServerNameParser.h"
#include "ItemsEvents.h"

LRESULT ServersDetectionProgressDlg::OnInitDialog(UINT, WPARAM, LPARAM, BOOL&)
{
    m_progress = GetDlgItem(IDC_DETECTION_PROGRESS);
    m_message = GetDlgItem(IDC_DETECTION_PART);
    SetEvent(m_hSyncEvent);
    return 0;
}

void ServersDetectionProgressDlg::SetSyncEvent(HANDLE hSyncEvent)
{
     m_hSyncEvent = hSyncEvent;
}

LRESULT ServersDetectionProgressDlg::OnBnClickedCancel(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
    m_bShouldCancel = true;

    return 0;
}

bool ServersDetectionProgressDlg::ShouldCancel(void)
{
    return m_bShouldCancel;
}

void ServersDetectionProgressDlg::SetPosition(int position)
{
    m_progress.SetPos(position);
}

void ServersDetectionProgressDlg::SetProgressRange(int range)
{
    m_progress.SetRange(0, range);
}

void ServersDetectionProgressDlg::SetMessage(LPCWSTR szMessage)
{
    m_message.SetWindowTextW(szMessage);
}

void ServersDetectionProgressDlg::Finish(int nRetCode)
{
    EndDialog(nRetCode);
}
