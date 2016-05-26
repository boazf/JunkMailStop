// ItemsEvents.cpp : Implementation of CItemsEvents

#include "stdafx.h"
#include "debug.h"
#include "ItemsEvents.h"
#include "AccountsUtil.h"
#include "ServerNameParser.h"
#include "Config.h"

_ATL_FUNC_INFO OnItemAddEventInfo ={
                    CC_STDCALL,
                    VT_EMPTY,
                    1,
                    {VT_DISPATCH}}; // (IDispatch *)

// CItemsEvents

HRESULT CItemsEvents::set_JunkFolder(Outlook::MAPIFolder *pFolder)
{
    m_spJunkFolder = pFolder;

    return S_OK;
}

void __stdcall CItemsEvents::OnItemAdd(IDispatch *pIDisp)
{
    HRESULT hr;
    CComQIPtr<Outlook::_MailItem> spMailItem(pIDisp); // QI for _MailItem object.

    if (spMailItem == NULL)
    {
        // Ignore any non mail items.
        DEBUG_TRACE(TRACE_VERBOSE_INFO, (L"Ignoring non email item."));
        return;
    }

    CComBSTR bstrSubject;

    hr = spMailItem->get_Subject(&bstrSubject);
    if (SUCCEEDED(hr))
    {
        DEBUG_TRACE(
            TRACE_VERBOSE_INFO,
            (L"OnItemAdd: Received new mail Item with subject: %s\n", bstrSubject));
    }
    else
    {
        DEBUG_TRACE(
            TRACE_ERRORS,
            (L"OnItemAdd: failed to botain mail subject\n"));
    }

    CAutoPtr<ServerNameParser> spMailReceivedFromServers;

    try
    {
        spMailReceivedFromServers.Attach(new ServerNameParser(spMailItem));
    }
    catch (...)
    {
        DEBUG_TRACE(
            TRACE_ERRORS,
            (L"Failed to instantiate ServerNameParser class\n"));
    }

    BlackList blackList; 
    g_pCConfig->GetBlackList(blackList);
    CComBSTR bstrServers(L"");
    DWORD nServers = 0;
    bool bMoved = false;

    do
    {
        CComBSTR bstrServerName;

        spMailReceivedFromServers->GetServer(bstrServerName);
        if (bstrServerName == NULL)
        {
            break;
        }

        DEBUG_TRACE(TRACE_VERBOSE_INFO, (L"From server: %s\n", bstrServerName));

        nServers++;
        if (nServers > 1)
        {
            bstrServers.Append("; ");
        }
        bstrServers.Append(bstrServerName);

        if (blackList.find(wstring(bstrServerName)) == blackList.end())
        {
            continue;
        }

        if (!bMoved)
        {
            DEBUG_TRACE(TRACE_INFO, ((L"Moving mail item to the junk folder\n")));

            ATLASSERT(m_spJunkFolder != NULL);

            CComPtr<IDispatch> spItem;

            hr = spMailItem->Move(m_spJunkFolder, &spItem);
            if (FAILED(hr))
            {
                DEBUG_TRACE(TRACE_ERRORS, (L"Failed to move mail item to the junk folder, hr=%x", hr));
            }
            else
            {
                spMailItem.Release();
                spItem.QueryInterface(&spMailItem);
                ATLASSERT(spMailItem != NULL);
                bMoved = true;
            }
        }
    } while (true);

    SetServersColumn(spMailItem, bstrServers);
}

HRESULT CItemsEvents::SetServersColumn(Outlook::_MailItem *pMailItem, BSTR bstrServers)
{
    HRESULT hr;
    // Retrieve the user properties collection.
    CComPtr<Outlook::UserProperties> spUserProps;

    hr = pMailItem->get_UserProperties(&spUserProps);
    if (FAILED(hr))
    {
        DEBUG_TRACE(
            TRACE_ERRORS,
            (L"Failed in pMailItem->get_UserProperties(), hr = 0x%x\n"));
        return hr;
    }

    // Get the name of the Servers column.
    CComBSTR bstrServersColName;

    if (!bstrServersColName.LoadString(IDS_SERVERS_COL_HEADING))
    {
        DEBUG_TRACE(
            TRACE_ERRORS,
            (L"Failed to load Servers column name string, error = %d\n", GetLastError()));
        return HRESULT_FROM_WIN32(GetLastError());
    }

    // Find the Servers property in the user properties collection.
    CComPtr<Outlook::UserProperty> spServersProp;
    
    hr = spUserProps->Find(bstrServersColName.m_str, CComVariant(DISP_E_PARAMNOTFOUND, VT_ERROR), &spServersProp);
    if (FAILED(hr))
    {
        DEBUG_TRACE(
            TRACE_ERRORS,
            (L"Failed in spUserProps->Find(), hr = 0x%x\n"));
        return hr;
    }

    if (spServersProp == NULL)
    {
        // The Servers property do not exist, add it.
        hr = spUserProps->Add(bstrServersColName.m_str, Outlook::olText, CComVariant(TRUE), CComVariant(1), &spServersProp);
        if (FAILED(hr))
        {
            DEBUG_TRACE(
                TRACE_ERRORS,
                (L"Failed in spUserProps->Add(\"Amount\"), hr = 0x%x, Trying to delete the property:\n", hr));
            return hr;
        }
        ATLASSERT(spServersProp != NULL);
    }

    // Set the value of the Servers property.
    hr = spServersProp->put_Value(CComVariant(bstrServers));
    if (FAILED(hr))
    {
        DEBUG_TRACE(
            TRACE_ERRORS,
            (L"Failed in spVCProp->put_Value(), hr = 0x%x\n"));
        return hr;
    }

    // Save the mail item, without it the changes will not take in effect.
    hr = pMailItem->Save();
    if (FAILED(hr))
    {
        DEBUG_TRACE(
            TRACE_ERRORS,
            (L"Failed in pMailItem->Save(), hr = 0x%x\n"));
        return hr;
    }

    return S_OK;
}