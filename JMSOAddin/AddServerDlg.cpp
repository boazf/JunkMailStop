#include "StdAfx.h"
#include "AddServerDlg.h"
#include "Config.h"

AddServerDlg::AddServerDlg(void)
{
}


AddServerDlg::~AddServerDlg(void)
{
}

LRESULT AddServerDlg::OnInitDialog(UINT, WPARAM, LPARAM, BOOL&)
{
    ATLControls::CEdit serverName = GetDlgItem(IDC_SERVER_NAME);

    serverName.SendMessageW(EM_SETLIMITTEXT, MAX_SERVER_NAME_LEN);

    return 0;
}

LRESULT AddServerDlg::OnCommand(UINT, WPARAM wParam, LPARAM, BOOL&)
{
    switch(wParam)
    {
    case IDOK:
        {
            ATLControls::CEdit serverName = GetDlgItem(IDC_SERVER_NAME);
            ATLASSERT(serverName != NULL);
            WCHAR szServerName[MAX_SERVER_NAME_LEN + 1];
            int ret = serverName.GetLine(0, szServerName, NELEMS(szServerName));
            ATLASSERT(ret <= MAX_SERVER_NAME_LEN);
            szServerName[ret] = L'\0';
            m_serverName = szServerName;
            EndDialog(IDOK);
        }
        break;
    case IDCANCEL:
        EndDialog(IDCANCEL);
        break;
    }

    return 0;
}

