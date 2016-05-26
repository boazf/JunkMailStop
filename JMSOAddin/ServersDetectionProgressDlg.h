#pragma once
#include "atlcontrols.h"
#include "atlwin.h"

class ServersDetectionProgressDlg : public CDialogImpl<ServersDetectionProgressDlg>
{
public:
    ServersDetectionProgressDlg(void) : m_bShouldCancel(false)
    {
    }

    ~ServersDetectionProgressDlg(void)
    {
    }

    enum { IDD = IDD_DETECT_SERVERS_PROGRESS_DIALOG };

    BEGIN_MSG_MAP(ServersDetectionProgressDlg)
        MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog);
        COMMAND_HANDLER(IDCANCEL, BN_CLICKED, OnBnClickedCancel)
    END_MSG_MAP()

private:
    ATLControls::CProgressBarCtrl m_progress;
    ATLControls::CStatic m_message;

public:
    bool m_bShouldCancel;
    HANDLE m_hSyncEvent;

public:
    LRESULT OnInitDialog(UINT, WPARAM, LPARAM, BOOL&);
    LRESULT OnBnClickedCancel(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);

public:
    void SetPosition(int position);
    void SetProgressRange(int range);
    void SetMessage(LPCWSTR szMessage);
    bool ShouldCancel();
    void Finish(int nRetCode);
    void SetSyncEvent(HANDLE hSyncEvent);
};

