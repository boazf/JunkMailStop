#pragma once
#include "atlcontrols.h"
#include "atlwin.h"
#include "Config.h"

class SelectServersDlg : public CDialogImpl<SelectServersDlg>
{
public:
    SelectServersDlg(BlackList blackList);
    ~SelectServersDlg(void);

    enum { IDD = IDD_SELECT_SERVERS };

    BEGIN_MSG_MAP(SelectServersDlg)
        MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog);
        COMMAND_HANDLER(IDC_SERVERS_LIST, LBN_SELCHANGE, OnLbnSelchangeServersList)
        COMMAND_HANDLER(IDOK, BN_CLICKED, OnBnClickedOk)
        COMMAND_HANDLER(IDCANCEL, BN_CLICKED, OnBnClickedCancel)
    END_MSG_MAP()

private:
    ATLControls::CListBox m_serversNames;
private:
    LRESULT OnInitDialog(UINT, WPARAM, LPARAM, BOOL&);

public:
    BlackList m_BlackList;
    
public:
    LRESULT OnLbnSelchangeServersList(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
    LRESULT OnBnClickedOk(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
    LRESULT OnBnClickedCancel(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
};

