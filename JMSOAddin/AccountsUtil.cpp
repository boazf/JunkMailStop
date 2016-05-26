#include "StdAfx.h"
#include "AccountsUtil.h"
#include "debug.h"


CAccountsUtil::CAccountsUtil(void)
{
}


CAccountsUtil::~CAccountsUtil(void)
{
}

CComPtr<Outlook::_Application> CAccountsUtil::m_spApp;


void CAccountsUtil::SetApplication(CComPtr<Outlook::_Application> spApp)
{
    ATLASSERT(m_spApp == NULL);
    m_spApp = spApp;
}

HRESULT CAccountsUtil::EnumAccounts(AccountEnumProc EnumProc, LPVOID lpvParam)
{
    HRESULT hr;

    ATLASSERT(m_spApp != NULL);
    // Get the inbox items collection in order to sink ItemAdd events.
    CComPtr<Outlook::_NameSpace> spSess;

    // Get the session object.
    hr = m_spApp->get_Session(&spSess);
    if (FAILED(hr))
    {
        DEBUG_TRACE(
            TRACE_ERRORS,
            (L"Failed in m_spApp->get_Session(), hr = 0x%x\n", hr));
        return hr;
    }

    CComPtr<Outlook::_Accounts> spAccounts;
    
    hr = spSess->get_Accounts(&spAccounts);
    if (FAILED(hr))
    {
        DEBUG_TRACE(
            TRACE_ERRORS,
            (L"Failed in spSess->get_Accounts(), hr = 0x%x\n", hr));
        return hr;
    }

    long dwAccounts;

    spAccounts->get_Count(&dwAccounts);
    if (FAILED(hr))
    {
        DEBUG_TRACE(
            TRACE_ERRORS,
            (L"Failed in spAccounts->get_Count(), hr = 0x%x\n", hr));
        return hr;
    }

    for (long nAccount = 0; nAccount < dwAccounts; nAccount++)
    {
        CComPtr<Outlook::_Account> spAccount;
        CComVariant index = nAccount + 1;

        hr = spAccounts->Item(index, &spAccount);
        if (FAILED(hr))
        {
            DEBUG_TRACE(
                TRACE_ERRORS,
                (L"Failed in spAccounts->Item(%d), hr = 0x%x\n", nAccount, hr));
            return hr;
        }

        BOOL bCont = true;

        hr = EnumProc(spAccount, &bCont, lpvParam);
        if (FAILED(hr))
        {
            DEBUG_TRACE(
                TRACE_ERRORS,
                (L"Failed in account enumeration procedure (%d), hr = 0x%x\n", nAccount, hr));
            return hr;
        }

        if (!bCont)
        {
            break;
        }
    }

    return S_OK;
}
