// JMSOAddinPropPage.cpp : Implementation of CJMSOAddinPropPage
#include "stdafx.h"
#include "JMSOAddinPropPage.h"
#include "debug.h"
#include "Config.h"
#include "AddServerDlg.h"

// CJMSOAddinPropPage

STDMETHODIMP CJMSOAddinPropPage::SetClientSite(IOleClientSite *pClientSite)
{
    HRESULT hr;

    // Call default ATL implementation
    hr = IOleObjectImpl<CJMSOAddinPropPage>::SetClientSite(pClientSite);
    if (hr != S_OK)
    {
        DEBUG_TRACE(
            TRACE_ERRORS,
            (L"Failed in IOleObjectImpl::SetClientSite(), hr = %x\n", hr));
        return hr;
    }

    // pClientSite may be NULL when container has being destructed
    if (pClientSite != NULL)
    {
        CComQIPtr<Outlook::PropertyPageSite> pPropertyPageSite(pClientSite);
        hr = pPropertyPageSite.CopyTo(&m_pPropPageSite);
        if (FAILED(hr))
        {
            DEBUG_TRACE(
                TRACE_ERRORS,
                (L"Failed in pPropertyPageSite.CopyTo(), hr = %x\n", hr));
        }
    }

    m_fDirty = false;
    return hr;
}

LRESULT CJMSOAddinPropPage::OnInitDialog(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{
    UNREFERENCED_PARAMETER(uMsg); 
    UNREFERENCED_PARAMETER(wParam); 
    UNREFERENCED_PARAMETER(lParam); 
    UNREFERENCED_PARAMETER(bHandled);
    m_fDirty = false;
    m_bInitializing = TRUE;

    m_serversBlackList = GetDlgItem(IDC_SERVERS_BLACK_LIST);
    ATLASSERT(m_serversBlackList);
    m_removeBtn = GetDlgItem(IDC_REMOVE_BUTTON);
    ATLASSERT(m_removeBtn);
    m_addBtn = GetDlgItem(IDC_ADD_BUTTON);
    ATLASSERT(m_addBtn);

    BlackList blackList;
    g_pCConfig->GetBlackList(blackList);
    for (BlackList::iterator i = blackList.begin(); i != blackList.end(); i++)
    {
        m_serversBlackList.AddString((*i).data());
    }
    m_bInitializing = FALSE;

    return 0;
}

STDMETHODIMP CJMSOAddinPropPage::GetControlInfo(LPCONTROLINFO lpCI)
{
    UNREFERENCED_PARAMETER(lpCI);
    return S_OK;
}

STDMETHODIMP CJMSOAddinPropPage::GetPageInfo(BSTR * HelpFile, long * HelpContext)
{
    UNREFERENCED_PARAMETER(HelpFile); 
    UNREFERENCED_PARAMETER(HelpContext);
    return E_NOTIMPL;
}

STDMETHODIMP CJMSOAddinPropPage::get_Dirty(VARIANT_BOOL * Dirty)
{
    *Dirty = m_fDirty.boolVal;
    return S_OK;
}

HRESULT CJMSOAddinPropPage::MarkAsDirty(void)
{
    HRESULT hr;

    m_fDirty = true;
    hr = m_pPropPageSite->OnStatusChange();
    if (FAILED(hr))
    {
        DEBUG_TRACE(
            TRACE_ERRORS,
            (L"Failed in m_pPropPageSite->OnStatusChange(), hr = 0x%x\n", hr));
    }

    return hr;
}

STDMETHODIMP CJMSOAddinPropPage::Apply()
{
    m_bInitializing = TRUE;

    BlackList blackList;
    int nServers = m_serversBlackList.GetCount();

    for (int i = 0; i < nServers; i++)
    {
        WCHAR szServerName[MAX_SERVER_NAME_LEN + 1];
        int ret = m_serversBlackList.GetText(i, szServerName);
        ATLASSERT(ret <= MAX_SERVER_NAME_LEN);
        szServerName[ret] = L'\0';
        blackList.insert(wstring(szServerName));
    }
    g_pCConfig->SetBlackList(blackList);

    m_bInitializing = FALSE;
    m_fDirty = false;

    return S_OK;
}



LRESULT CJMSOAddinPropPage::OnLbnSelchangeServersBlackList(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
    int iSelCount = m_serversBlackList.GetSelCount();

    m_removeBtn.EnableWindow(iSelCount > 0);

    return 0;
}

static int comp_int(const void *pi1, const void *pi2)
{
    return *((int *)pi2) - *((int *)pi1);
}

LRESULT CJMSOAddinPropPage::OnBnClickedRemoveButton(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
    int iSelCount = m_serversBlackList.GetSelCount();
    ATLASSERT(iSelCount > 0);
    CAutoPtr<int> pSelItems(new int[iSelCount]);
    m_serversBlackList.GetSelItems(iSelCount, pSelItems);
    qsort(pSelItems, iSelCount, sizeof(int), comp_int);
    for (int i = 0; i < iSelCount; i++)
    {
        m_serversBlackList.DeleteString(pSelItems[i]);
    }

    m_removeBtn.EnableWindow(m_serversBlackList.GetSelCount() > 0);
    MarkAsDirty();

    return 0;
}


LRESULT CJMSOAddinPropPage::OnBnClickedAddButton(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
    AddServerDlg dlg;
    int ret = dlg.DoModal();
    if (ret == IDOK)
    {
        m_serversBlackList.AddString(dlg.m_serverName.data());
        MarkAsDirty();
    }

    return 0;
}
