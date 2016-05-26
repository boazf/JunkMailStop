#include "StdAfx.h"
#include "SelectServersDlg.h"


SelectServersDlg::SelectServersDlg(BlackList blackList)
{
    m_BlackList = blackList;
}


SelectServersDlg::~SelectServersDlg(void)
{
}

LRESULT SelectServersDlg::OnInitDialog(UINT, WPARAM, LPARAM, BOOL&)
{
    m_serversNames = GetDlgItem(IDC_SERVERS_LIST);
    ATLASSERT(m_serversNames != NULL);

    for (BlackList::const_iterator i = m_BlackList.cbegin();
         i != m_BlackList.cend();
         i++)
    {
        m_serversNames.AddString((*i).data());
    }

    GetDlgItem(IDOK).EnableWindow(m_serversNames.GetSelCount() > 1);

    return 0;
}

LRESULT SelectServersDlg::OnLbnSelchangeServersList(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
    ATLControls::CButton okButton = GetDlgItem(IDOK);
    okButton.EnableWindow(m_serversNames.GetSelCount() > 0);

    return 0;
}


LRESULT SelectServersDlg::OnBnClickedOk(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
    WCHAR szServerName[MAX_SERVER_NAME_LEN + 1];
    int selCount = m_serversNames.GetSelCount();
    ATLASSERT(selCount > 0);
    CAutoPtr<int> pSelItems(new int[selCount]);
    m_serversNames.GetSelItems(selCount, pSelItems);
    m_BlackList.clear();
    for (int i = 0; i < selCount; i++)
    {
        int ret = m_serversNames.GetText(pSelItems[i], szServerName);
        ATLASSERT(ret <= MAX_SERVER_NAME_LEN);
        szServerName[ret] = L'\0';
        m_BlackList.insert(wstring(szServerName));
    }

    EndDialog(IDOK);

    return 0;
}


LRESULT SelectServersDlg::OnBnClickedCancel(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
    EndDialog(IDCANCEL);

    return 0;
}
