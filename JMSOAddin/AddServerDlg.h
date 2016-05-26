#pragma once
#include "string"
#include "atlcontrols.h"
#include "atlwin.h"

class AddServerDlg : public CDialogImpl<AddServerDlg>
{
public:
    AddServerDlg(void);
    ~AddServerDlg(void);

    enum { IDD = IDD_ADD_SERVER };

    BEGIN_MSG_MAP(AddServerDlg)
        MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog);
        MESSAGE_HANDLER(WM_COMMAND, OnCommand)
    END_MSG_MAP()

public:
    wstring m_serverName;
    
private:
    LRESULT OnCommand(UINT, WPARAM wParam, LPARAM lParam, BOOL&);
    LRESULT OnInitDialog(UINT, WPARAM, LPARAM, BOOL&);
};

