// CJMSOAddin.cpp : Implementation of CJMSOAddin

#include "stdafx.h"
#include <mapix.h>
#include "CJMSOAddin.h"
#include "ItemsEvents.h"
#include "AccountsUtil.h"
#include "msxml2.h"
#include "ServerNameParser.h"
#include "Config.h"
#include "SelectServersDlg.h"
#include "ServersDetection.h"
#include "SelectionUtils.h"
#include "AddinMessageBox.h"
#include "RibbonManifest.h"

// Descriptors of parameters of event sink functions.
_ATL_FUNC_INFO OnSimpleEventInfo ={
                    CC_STDCALL,
                    VT_EMPTY,
                    0}; // No Parameters
_ATL_FUNC_INFO OnOptionsAddPagesInfo={
                    CC_STDCALL,
                    VT_EMPTY,
                    1,
                    {VT_DISPATCH}}; // (IDispatch *)

// CJMSOAddin

STDMETHODIMP CJMSOAddin::InterfaceSupportsErrorInfo(REFIID riid)
{
    static const IID* arr[] = 
    {
        &IID_IJMSOAddin
    };

    for (int i=0; i < sizeof(arr) / sizeof(arr[0]); i++)
    {
        if (InlineIsEqualGUID(*arr[i],riid))
            return S_OK;
    }
    return S_FALSE;
}

HRESULT STDMETHODCALLTYPE CJMSOAddin::raw_GetCustomUI(BSTR RibbonID, BSTR *bstrRibbonXml)
{
    UNREFERENCED_PARAMETER(RibbonID);
    if(!bstrRibbonXml)
    {
        DEBUG_TRACE(TRACE_ERRORS, (L"Received NULL parameter in bstrRibbonXml\n"));
        return E_POINTER;
    }

    HRESULT hr = RibbonManifest::GetRibbonManifest(bstrRibbonXml);
    if (FAILED(hr))
    {
        DEBUG_TRACE(TRACE_ERRORS, (L"Failed in RibbonManifest::GetRibbonManifest, error=0x%x\n", hr));
    }

    return hr;
}

typedef struct _AddServersToBlackListParams
{
    BlackList blackList;
    BlackList configBlackList;
} AddServersToBlackListParams;

static HRESULT AddServersToBlackListEnumProc(IDispatch *pDisp, LPVOID lpParams)
{
    AddServersToBlackListParams *params = static_cast<AddServersToBlackListParams *>(lpParams);

    CComQIPtr<Outlook::_MailItem> spMailItem(pDisp);
    if (spMailItem == NULL)
    {
        DEBUG_TRACE(TRACE_INFO, (L"Ignoring non-mail item.\n"));
        return S_OK; 
    }

    ServerNameParser parser(spMailItem);
    CComBSTR server;
    do
    {
        parser.GetServer(server);
        if (server != NULL)
        {
            wstring serverName(wstring(server.m_str));
            if (params->blackList.find(serverName) == params->blackList.cend() &&
                params->configBlackList.find(serverName) == params->configBlackList.cend())
            {
                params->blackList.insert(serverName);
            }
        }
    } while (server != NULL);

    return S_OK;
}

static HRESULT AddServersToBlackList(IDispatch* pDisp)
{
    CComQIPtr<Outlook::Selection> spSelection(pDisp);
    if (spSelection == NULL)
    {
        DEBUG_TRACE(TRACE_ERRORS, (L"Failed to query Selection interface!\n"));
        return E_NOINTERFACE;
    }

    AddServersToBlackListParams params;
    g_pCConfig->GetBlackList(params.configBlackList);
    SelectionUtils selectionUtils(spSelection);
    HRESULT hr = selectionUtils.EnumSelectionItems(AddServersToBlackListEnumProc, &params);
    if (FAILED(hr))
    {
        DEBUG_TRACE(TRACE_ERRORS, (L"Failed in selectionUtils.EnumSelectionItems, error=0x%x\n", hr));
        return hr;
    }

    if (params.blackList.size() > 0)
    {
        SelectServersDlg selectServersDlg(params.blackList);
        int ret = selectServersDlg.DoModal();
        if (ret == IDOK)
        {
            g_pCConfig->AddServersToBlackList(selectServersDlg.m_BlackList);
        }
    }
    else
    {
        AddinMessageBox(IDS_NO_SERVERS_TO_ADD, MB_ICONINFORMATION);
    }

    return S_OK;
}

static HRESULT GetRibbonInfo(IDispatch *pDisp, CComPtr<Office::IRibbonControl> &spRibbon, CComBSTR &bstrMenuId, int &iMenuId)
{
    HRESULT hr;

    pDisp->QueryInterface<Office::IRibbonControl>(&spRibbon);
    if (spRibbon == NULL)
    {
        DEBUG_TRACE(TRACE_ERRORS, (L"Failed to query IRibbonControl\n"));
        return E_NOINTERFACE;
    }

    hr = spRibbon->get_Id(&bstrMenuId);
    if (spRibbon == NULL)
    {
        DEBUG_TRACE(TRACE_ERRORS, (L"Failed in spRibbon->get_Id(), error=0x%x\n", hr));
        return hr;
    }

    DEBUG_TRACE(TRACE_VERBOSE_INFO, (L"Clicked menu: %s\n", bstrMenuId.m_str));

    iMenuId = RibbonManifest::GetMenuId(bstrMenuId.m_str);

    return S_OK;
}

HRESULT STDMETHODCALLTYPE CJMSOAddin::GetVisible(IDispatch *pDisp, VARIANT_BOOL *pvarReturnedVal)
{
    CComQIPtr<Office::IRibbonControl> spRibbon;
    int menuId;
    CComBSTR bstrMenuId;
    HRESULT hr = GetRibbonInfo(pDisp, spRibbon, bstrMenuId, menuId);
    if (FAILED(hr))
    {
        DEBUG_TRACE(TRACE_ERRORS, (L"Failed to query GetRibbonInfo, error = 0x%x\n", hr));
        return E_NOINTERFACE;
    }

    switch(menuId)
    {
    case ID_CONTEXT_MENU_MAIL_ITEM_ADD_SERVERS:
    case ID_CONTEXT_MENU_MULTIPLE_ITEMS_ADD_SERVERS:
    case ID_CONTEXT_MENU_MAIL_ITEM_DETECT_SERVER:
    case ID_CONTEXT_MENU_MULTIPLE_ITEMS_DETECT_SERVERS:
        break;

    case ID_CONTEXT_MENU_FOLDER_DETECT_SERVERS:
        {
            CComQIPtr<Outlook::MAPIFolder> spFolder(spRibbon->Context);

            if (spFolder == NULL)
            {
                DEBUG_TRACE(TRACE_ERRORS, (L"Failed to query MAPIFolder interface!\n"));
                return E_NOINTERFACE;
            }

            Outlook::OlItemType olDefaultItemType;

            hr = spFolder->get_DefaultItemType(&olDefaultItemType);
            if (FAILED(hr))
            {
                DEBUG_TRACE(TRACE_ERRORS, (L"Failed in spFolder->get_DefaultItemType(), error=0x%x\n", hr));
                return hr;
            }
            *pvarReturnedVal = olDefaultItemType == Outlook::olMailItem ? VARIANT_TRUE : VARIANT_FALSE;
        }
        break;

    default:
        DEBUG_TRACE(TRACE_INFO, (L"Clicked unknown control, id=%s\n", bstrMenuId.m_str));
    }

    return S_OK;
}

HRESULT STDMETHODCALLTYPE CJMSOAddin::ButtonClicked(IDispatch* pDisp)
{
    CComQIPtr<Office::IRibbonControl> spRibbon;
    int menuId;
    CComBSTR bstrMenuId;
    HRESULT hr = GetRibbonInfo(pDisp, spRibbon, bstrMenuId, menuId);
    if (FAILED(hr))
    {
        DEBUG_TRACE(TRACE_ERRORS, (L"Failed in GetRibbonInfo(), error = 0x%x\n", hr));
        return hr;
    }

    switch(menuId)
    {
    case ID_CONTEXT_MENU_MAIL_ITEM_ADD_SERVERS:
    case ID_CONTEXT_MENU_MULTIPLE_ITEMS_ADD_SERVERS:
        {
            hr = AddServersToBlackList(spRibbon->Context);
            if (FAILED(hr))
            {
                DEBUG_TRACE(TRACE_ERRORS, (L"Failed in AddServersToBlackList, error=0x%x\n"));
                if (menuId == ID_CONTEXT_MENU_MAIL_ITEM_ADD_SERVERS)
                {
                    AddinMessageBox(IDS_FAILED_TO_PROCESS_MAIL_ITEM, MB_ICONERROR);
                }
                else
                {
                    AddinMessageBox(IDS_FAILED_TO_PROCESS_MAIL_ITEMS, MB_ICONERROR);
                }
            }
        }
        break;

    case ID_CONTEXT_MENU_MAIL_ITEM_DETECT_SERVER:
    case ID_CONTEXT_MENU_MULTIPLE_ITEMS_DETECT_SERVERS:
        {
            CComQIPtr<Outlook::Selection> spSelection(spRibbon->Context);
            if (spSelection == NULL)
            {
                DEBUG_TRACE(TRACE_ERRORS, (L"Failed to query Selection interface!\n"));
                ServersDetection<Outlook::Selection>::ShowErrorMessageBox();
            }

            ServersDetection<Outlook::Selection> serversDetection;
            serversDetection.DetectServers(spSelection);
        }
        break;

    case ID_CONTEXT_MENU_FOLDER_DETECT_SERVERS:
        {
            CComQIPtr<Outlook::MAPIFolder> spFolder(spRibbon->Context);

            if (spFolder == NULL)
            {
                DEBUG_TRACE(TRACE_ERRORS, (L"Failed to query MAPIFolder interface!\n"));
                ServersDetection<Outlook::_Items>::ShowErrorMessageBox();
            }

            CComPtr<Outlook::_Items> spItems;

            hr = spFolder->get_Items(&spItems);
            if (FAILED(hr))
            {
                DEBUG_TRACE(TRACE_ERRORS, (L"Failed in m_spFolder->get_Items(), error=0x%x\n", hr));
                ServersDetection<Outlook::_Items>::ShowErrorMessageBox();
                return hr;
            }

            ServersDetection<Outlook::_Items> serversDetection;
            serversDetection.DetectServers(spItems);
        }
        break;

    default:
        DEBUG_TRACE(TRACE_INFO, (L"Clicked unknown control, id=%s\n", bstrMenuId.m_str));
    }

    return hr;
}

//
// _IDTExtensibility2 Methods Implementation
//
HRESULT STDMETHODCALLTYPE CJMSOAddin::OnConnection(
    LPDISPATCH Application, 
    ext_ConnectMode ConnectMode, 
    LPDISPATCH AddInInst, 
    SAFEARRAY * * custom)
{
    UNREFERENCED_PARAMETER(ConnectMode);
    UNREFERENCED_PARAMETER(AddInInst); 
    UNREFERENCED_PARAMETER(custom);

    g_pCConfig = new CJMSOAddinConfig();
    if (g_pCConfig == NULL)
    {
        DEBUG_TRACE(TRACE_ERRORS, (L"Failed to allocate configuration object!\n"));
        return HRESULT_FROM_WIN32(ERROR_OUTOFMEMORY);
    }

    HRESULT hr;

    DEBUG_TRACE(TRACE_INFO, (L"OnConnection.\n"));

    CComQIPtr <Outlook::_Application> spApp(Application); 
    ATLASSERT(spApp);
    m_spApp = spApp; //Store the application object
    CAccountsUtil::SetApplication(spApp);
    
    //Sink Application Events.
    hr = AppEvents::DispEventAdvise((IDispatch*)m_spApp,
                                    &__uuidof(Outlook::ApplicationEvents_10));
    if (FAILED(hr))
    {
        DEBUG_TRACE(
            TRACE_ERRORS, 
            (L"Failed to set a sink for aplication events, hr = 0x%x\n", hr));
        return hr;
    }

    hr = CAccountsUtil::EnumAccounts(CJMSOAddin::SinkAccountInboxAddItemEvents, &m_AccountInfoList);
    if (FAILED(hr))
    {
        DEBUG_TRACE(
            TRACE_ERRORS, 
            (L"Failed in EnumAccounts(), hr = 0x%x\n", hr));
        return hr;
    }

    hr = AddAddinColumnToOutlookView();
    if (FAILED(hr))
    {
        DEBUG_TRACE(
            TRACE_ERRORS, 
            (L"Failed in AddAddinColumnToOutlookView(), hr = 0x%x\n", hr));
        return hr;
    }

    DEBUG_TRACE(TRACE_INFO, (L"Completed successfully\n"));

    return S_OK;
}

HRESULT CJMSOAddin::SinkAccountInboxAddItemEvents(Outlook::_Account *lpAccount, BOOL *lpbCont, LPVOID lpvParam)
{
    UNREFERENCED_PARAMETER(lpbCont);
    HRESULT hr;
    CComBSTR bstrAccountName;
    
    CAtlList<AccountInfo *> *pAccountInfoList = reinterpret_cast<CAtlList<AccountInfo *> *>(lpvParam);

    hr = lpAccount->get_DisplayName(&bstrAccountName);
    if (SUCCEEDED(hr))
    {
        DEBUG_TRACE(
            TRACE_VERBOSE_INFO,
            (L"Initializing account %s\n", bstrAccountName));
    }
    else
    {
        DEBUG_TRACE(
            TRACE_INFO,
            (L"Failed in spAccount->get_DisplayName(), hr = 0x%x\n", hr));
    }

    CComPtr<Outlook::_Store> spStore;

    hr = lpAccount->get_DeliveryStore(&spStore);
    if (FAILED(hr))
    {
        DEBUG_TRACE(
            TRACE_ERRORS,
            (L"Failed in spAccount->get_DeliveryStore(%s), hr = 0x%x\n", bstrAccountName, hr));
        return hr;
    }

    CComPtr<Outlook::MAPIFolder> spFolder;

    hr = spStore->GetDefaultFolder(Outlook::olFolderInbox, &spFolder);
    if (FAILED(hr))
    {
        DEBUG_TRACE(
            TRACE_ERRORS,
            (L"Failed in spStore->GetDefaultFolder(%s), hr = 0x%x\n", bstrAccountName, hr));
        return hr;
    }

    CComPtr<Outlook::_Items> spItems;

    hr = spFolder->get_Items(&spItems);
    if (FAILED(hr))
    {
        DEBUG_TRACE(
            TRACE_ERRORS,
            (L"Failed in spFolder->get_Items(%s), hr = 0x%x\n", bstrAccountName, hr));
        return hr;
    }

    //Sink Items Events.

    CComPtr<CComObject<CItemsEvents>> spItemsEvents;

    hr = CComObject<CItemsEvents>::CreateInstance(&spItemsEvents);
    if (FAILED(hr))
    {
        DEBUG_TRACE(
            TRACE_ERRORS,
            (L"Failed in CComObject<CItemsEvents>::CreateInstance(%s), hr = 0x%x\n", bstrAccountName, hr));
        return hr;
    }

    hr = spItemsEvents->DispEventAdvise((IDispatch*)spItems, &__uuidof(Outlook::ItemsEvents));
    if (FAILED(hr))
    {
        DEBUG_TRACE(
            TRACE_ERRORS,
            (L"Failed in spItemsEvents->DispEventAdvise((%s), hr = 0x%x\n", bstrAccountName, hr));
        return hr;
    }

    CComPtr<Outlook::MAPIFolder> spJunkFolder;

    hr = spStore->GetDefaultFolder(Outlook::olFolderJunk, &spJunkFolder);
    if (FAILED(hr))
    {
        DEBUG_TRACE(
            TRACE_ERRORS,
            (L"Failed in spStore->GetDefaultFolder(%s), hr = 0x%x\n", bstrAccountName, hr));
        return hr;
    }

    hr = spItemsEvents->set_JunkFolder(spJunkFolder);
    if (FAILED(hr))
    {
        DEBUG_TRACE(
            TRACE_ERRORS,
            (L"Failed in spItemsEvents->set_JunkFolder(%s), hr = 0x%x\n", bstrAccountName, hr));
        return hr;
    }

    AccountInfo *pAccountInfo = new AccountInfo();
    pAccountInfo->spItems.Attach(spItems.Detach());
    pAccountInfo->spItemsEvents.Attach(spItemsEvents.Detach());
    pAccountInfoList->AddTail(pAccountInfo);

    return S_OK;
}

HRESULT STDMETHODCALLTYPE CJMSOAddin::OnDisconnection(
    ext_DisconnectMode RemoveMode, 
    SAFEARRAY * * custom)
{
    UNREFERENCED_PARAMETER(RemoveMode); 
    UNREFERENCED_PARAMETER(custom);
    HRESULT hr1, hr = S_OK;

    DEBUG_TRACE(TRACE_INFO, (L"Disconnecting.\n"));

    // Unadvise COM to stop notification of application events.
    hr1 = AppEvents::DispEventUnadvise((IDispatch*)m_spApp);
    if (FAILED(hr1))
    {
        DEBUG_TRACE(
            TRACE_ERRORS,
            (L"Failed to Unadvise application events, hr = %x\n", hr1));
        hr = hr1;
    }

    while (m_AccountInfoList.GetCount() > 0)
    {
        AccountInfo *lpAccountInfo = m_AccountInfoList.RemoveHead();
        hr1 = lpAccountInfo->spItemsEvents.Detach()->DispEventUnadvise((IDispatch *)lpAccountInfo->spItems);
        if (FAILED(hr1))
        {
            DEBUG_TRACE(
                TRACE_ERRORS,
                (L"Failed to Unadvise items events (%d), hr = %x\n", hr1));
            hr = hr1;
        }
        delete lpAccountInfo;
    }

    delete g_pCConfig;

    return hr;
}

HRESULT STDMETHODCALLTYPE CJMSOAddin::OnAddInsUpdate(SAFEARRAY * * custom)
{
    UNREFERENCED_PARAMETER(custom);
    return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE CJMSOAddin::OnStartupComplete(SAFEARRAY * * custom)
{
    UNREFERENCED_PARAMETER(custom);
    DEBUG_TRACE(TRACE_INFO, (L"OnStartupComplete.\n"));

    return S_OK;
}

HRESULT STDMETHODCALLTYPE CJMSOAddin::OnBeginShutdown(SAFEARRAY * * custom)
{
    UNREFERENCED_PARAMETER(custom);
    return S_OK;;
}

// Options Add Pages application event handler.
void __stdcall CJMSOAddin::OnOptionsAddPages(IDispatch *Ctrl)
{
    HRESULT hr;

    // QI PropertyPages interface pointer
    CComQIPtr<Outlook::PropertyPages> spPages(Ctrl);

    if (spPages == NULL)
    {
        DEBUG_TRACE(TRACE_ERRORS, (L"Failed to QI PropertyPages interfacepointer.\n"));
        return;
    }

    // Load the property page title string.
    CComBSTR bstrPageTitle;

    if (!bstrPageTitle.LoadString(IDS_PROPPAGETITLE))
    {
        DEBUG_TRACE(TRACE_ERRORS, (L"Failed to load string, error = %d\n", GetLastError()));
        return;
    }

    // Add the Junk Mail Stop property paegs of the options dialog box.
    CComVariant varProgId(OLESTR("JMSOAddin.JMSOAddinPropPage"));
    CComBSTR bstrTitle(bstrPageTitle.m_str);

    hr = spPages->Add((_variant_t)varProgId,(_bstr_t)bstrTitle);
    if (FAILED(hr))
    {
        DEBUG_TRACE(
            TRACE_ERRORS,
            (L"Failed to add Junk Mail Stop property page, hr = %x\n"));
    }
    else
    {
        DEBUG_TRACE(
            TRACE_VERBOSE_INFO, 
            (L"Junk Mail Stop property page was added to Outlook options dialog.\n"));
    }
}

LPCWSTR g_wszServersSchema;  // A global string that holds the Servers coumn schema

// A structure that describes an XML node (node name and node value).
struct NAME_VALUE_PAIR {
    LPCWSTR lpwszName;      // Node name
    LPCWSTR lpwszValue;     // Node value.
};

//
// This function creates and builds a new Servers XML column node according to the 
// infomation found in the paNameValuePairs array.
//
static HRESULT BuildColumnNode(
    IXMLDOMDocument2 *pXMLDoc,              // The XML document in which the column is to be added
    NAME_VALUE_PAIR *paNameValuePairs,      // An array of string pairs the dfefines the sub nodes
    ULONG nArrayEmlements,                  // The number of entries in the paNameValuePairs array
    IXMLDOMNode **ppColNode)                // A pointer to a pointer that points to the newly create node.
{
    HRESULT hr;

    // Create a node for the column
    hr = pXMLDoc->createNode(CComVariant(NODE_ELEMENT), L"column", NULL, ppColNode);
    if (FAILED(hr))
    {
        DEBUG_TRACE(
            TRACE_ERRORS, 
            (L"Failed in spXMLDoc->createNode(\"column\"), hr = %x\n", hr));
        return hr;
    }

    DEBUG_TRACE(TRACE_INFO, (L"Created a column node.\n"));

    // Build the sub-nodes.
    for (ULONG i = 0; i < nArrayEmlements; i++)
    {
        // Create a sub-node.
        CComPtr<IXMLDOMNode> spSubNode;

        hr = pXMLDoc->createNode(CComVariant(NODE_ELEMENT), 
                                 CComBSTR(paNameValuePairs[i].lpwszName), 
                                 NULL, 
                                 &spSubNode);
        if (FAILED(hr))
        {
            DEBUG_TRACE(
                TRACE_ERRORS, 
                (L"Failed in spXMLDoc->createNode(\"%s\"), hr = %x\n", 
                paNameValuePairs[i].lpwszName));
            return hr;
        }

        DEBUG_TRACE(
            TRACE_INFO, 
            (L"Created %s column sub-node\n", paNameValuePairs[i].lpwszName));

        // Put the value of the sub-node
        hr = spSubNode->put_nodeTypedValue(CComVariant(paNameValuePairs[i].lpwszValue));
        if (FAILED(hr))
        {
            DEBUG_TRACE(
                TRACE_ERRORS, 
                (L"Failed in spSubNode->put_nodeTypedValue(\"%s\"), hr = %x\n", 
                paNameValuePairs[i].lpwszValue, hr));
            return hr;
        }

        DEBUG_TRACE(
            TRACE_INFO,
            (L"Added %s value to sub-node.\n", paNameValuePairs[i].lpwszValue));

        // Append the sub-node to the children of the column node.
        hr = (*ppColNode)->appendChild(spSubNode, NULL);
        if (FAILED(hr))
        {
            DEBUG_TRACE(
                TRACE_ERRORS, 
                (L"Failed in (*ppColNode)->appendChild(), hr = %x\n", hr));
            return hr;
        }

        DEBUG_TRACE(TRACE_INFO, (L"sub-node appended.\n"));
    }

    return S_OK;
}

#define SCHEMA_PREFIX L"http://schemas.microsoft.com/mapi/string/{00020329-0000-0000-C000-000000000046}/"

//
// This function looks whether the Verdicach Outlook columns are shown in the current
// view of the inbox folder. If one or more of the columns are missing, it adds them
// to the view.
//
HRESULT CJMSOAddin::AddAddinColumnToOutlookView(void)
{
    HRESULT hr;
    CComBSTR bstrServers;
    CComBSTR bstrReceived(L"Received");

    // Load the column names.
    if (!bstrServers.LoadString(IDS_SERVERS_COL_HEADING))
    {
        DEBUG_TRACE(TRACE_ERRORS, (L"Failed to load servers column heading string, error = \n", GetLastError()));
        return HRESULT_FROM_WIN32(GetLastError());
    }

    // Get the session object.
    CComPtr<Outlook::_NameSpace> spSess;

    hr = m_spApp->get_Session(&spSess);
    if (FAILED(hr))
    {
        DEBUG_TRACE(TRACE_ERRORS, (L"Failed in pApp->get_Session(), hr = %x\n", hr));
        return hr;
    }

    // Get the inbox folder.
    CComPtr<Outlook::MAPIFolder> spFolder;

    hr = spSess->GetDefaultFolder(Outlook::olFolderInbox, &spFolder);
    if (FAILED(hr))
    {
        DEBUG_TRACE(TRACE_ERRORS, (L"Failed in spSess->GetDefaultFolder(), hr = %x\n", hr));
        return hr;
    }

    // Following code is a workaround for an apparent bug that causes the columns XML to be partial.
    // with the following code, for some unknown reason, the columns XML appears to be complete.
    CComPtr<Outlook::_Views> spViews;
    hr = spFolder->get_Views(&spViews);
    if (FAILED(hr))
    {
        DEBUG_TRACE(TRACE_ERRORS, (L"Failed in spFolder->get_Views(), hr = %x\n", hr));
    }
    else
    {
        LONG nViews;
        hr = spViews->get_Count(&nViews);
        if (FAILED(hr))
        {
            DEBUG_TRACE(TRACE_ERRORS, (L"Failed in spViews->get_Count(), hr = %x\n", hr));
        }
        else
        {
            for (LONG i = 0; i < nViews; i++)
            {
                CComPtr<Outlook::View> spView;
                hr = spViews->Item(CComVariant(i + 1), &spView);
                if (FAILED(hr))
                {
                    DEBUG_TRACE(TRACE_ERRORS, (L"Failed in spViews->Item(), hr = %x\n", hr));
                }
                else
                {
                    CComBSTR bstrViewName;
                    hr = spView->get_Name(&bstrViewName);
                    if (FAILED(hr))
                    {
                        DEBUG_TRACE(TRACE_ERRORS, (L"Failed in spView->get_Name(), hr = %x\n", hr));
                    }
                }
            }
        }
    }
    // End of workaround

    // Get the view object of the inbox folder
    CComPtr<Outlook::View> spView;

    hr = spFolder->get_CurrentView(&spView);
    if (FAILED(hr))
    {
        DEBUG_TRACE(TRACE_ERRORS, (L"Failed in spFolder->get_CurrentView(), hr = %x\n", hr));
        return hr;
    }

    // Get the XML of the view. The view content and representation is defined by this XML
    CComBSTR bstrXML;

    hr = spView->get_XML(&bstrXML);
    if (FAILED(hr))
    {
        DEBUG_TRACE(TRACE_ERRORS, (L"Failed in spView->get_XML(), hr = %x\n", hr));
        return hr;
    }

    // Create an XML DOM document object.
    CComPtr<IXMLDOMDocument2> spXMLDoc;
    
    hr = CoCreateInstance(CLSID_DOMDocument30, NULL, CLSCTX_INPROC_SERVER, 
           IID_IXMLDOMDocument2, (void**)&spXMLDoc);
    if (FAILED(hr))
    {
        DEBUG_TRACE(TRACE_ERRORS, (L"Failed in CoCreateInstance(), hr = %x\n", hr));
        return hr;
    }
    
    // Load the XML into the newly created XML DOM document object.
    VARIANT_BOOL bSuccess;

    hr = spXMLDoc->loadXML(bstrXML,&bSuccess);
    if (FAILED(hr) || !bSuccess)
    {
        DEBUG_TRACE(
            TRACE_ERRORS, 
            (L"Failed in spXMLDoc->loadXML(), hr = %x, bSuccess = %d\n", 
            hr, bSuccess));
        return hr;
    }
    
    // Get the column nodes.
    CComPtr<IXMLDOMNodeList> spColNodes;  

    hr = spXMLDoc->getElementsByTagName(L"column",&spColNodes);
    if (FAILED(hr))
    {
        DEBUG_TRACE(
            TRACE_ERRORS, 
            (L"Failed in spXMLDoc->getElementsByTagName(\"column\"), hr = %x\n", hr));
        return hr;
    }

    // Get the number of column nodes
    long nColNodes;

    hr = spColNodes->get_length(&nColNodes);
    if (FAILED(hr))
    {
        DEBUG_TRACE(
            TRACE_ERRORS, 
            (L"Failed in spColNodes->get_length(), hr = %x\n", hr));
        return hr;
    }

    DEBUG_TRACE(TRACE_VERBOSE_INFO, (L"There are %d columns in outlook view\n", nColNodes));

    BOOL bServersColExist = FALSE;
    BOOL bReceivedNodeExist = FALSE;
    long iReceivedNode = 0;

    // Scan the column nodes. Find the Servers column and the columns before which the
    // Servers column shold be positioned.
    for (long i = 0; i < nColNodes && !bServersColExist && !bReceivedNodeExist; i++)
    {
        // Get the i-th node.
        CComPtr<IXMLDOMNode> spXDN;

        hr = spColNodes->get_item(i, &spXDN);
        if (FAILED(hr))
        {
            DEBUG_TRACE(
                TRACE_ERRORS, 
                (L"Failed in spColNodes->get_item(%d), hr = %x\n", i, hr));
            return hr;
        }

        // Get the child nodes collection.
        CComPtr<IXMLDOMNodeList> spNodeList;

        hr = spXDN->get_childNodes(&spNodeList);
        if (FAILED(hr))
        {
            DEBUG_TRACE(
                TRACE_ERRORS, 
                (L"Failed in spXDN->get_childNodes() of a column node, hr = %x\n", hr));
            return hr;
        }

        // Get the number of child nodes.
        long nChildNodes;

        hr = spNodeList->get_length(&nChildNodes);
        if (FAILED(hr))
        {
            DEBUG_TRACE(
                TRACE_ERRORS, 
                (L"Failed in spNodeList->get_length(), hr = %x\n", hr));
            return hr;
        }

        DEBUG_TRACE(TRACE_VERBOSE_INFO, (L"Column %d has %d child nodes.\n", i, nChildNodes));

        // Scan the child nodes.
        for (long iNode = 0; iNode < nChildNodes; iNode++)
        {
            // Get the i-th child node
            CComPtr<IXMLDOMNode> spXDN2;
        
            hr = spNodeList->get_item(iNode, &spXDN2);
            if (FAILED(hr))
            {
                DEBUG_TRACE(
                    TRACE_ERRORS, 
                    (L"Failed in spNodeList->get_item(%d), hr = %x\n", iNode, hr));
                return hr;
            }
        
            // Get the name of the child node.
            CComBSTR bstrNodeName;

            hr = spXDN2->get_nodeName(&bstrNodeName);
            if (FAILED(hr))
            {
                DEBUG_TRACE(
                    TRACE_ERRORS, 
                    (L"Failed in spXDN2->get_nodeName(), hr = %x\n", hr));
                return hr;
            }

            DEBUG_TRACE(
                TRACE_VERBOSE_INFO, 
                (L"Column: %i, child node: %d, child node name: %s\n", i, iNode, bstrNodeName.m_str));

            // The heading child node, defines the name of the column.
            if (wcscmp(bstrNodeName, L"heading") == 0)
            {
                // Get the value of the child node
                CComVariant vtValue;

                hr = spXDN2->get_nodeTypedValue(&vtValue);
                if (FAILED(hr))
                {
                    DEBUG_TRACE(
                        TRACE_ERRORS, 
                        (L"Failed in spXDN2->get_nodeTypedValue(), hr = %x\n", hr));
                    return hr;
                }


                DEBUG_TRACE(TRACE_VERBOSE_INFO, (L"Heading value: %s\n", vtValue.bstrVal));

                // See if this is the Servers node
                if (wcscmp(vtValue.bstrVal, bstrServers.m_str) == 0)
                {
                    // Mark that the Servers node exist.
                    bServersColExist = TRUE;

                    DEBUG_TRACE(TRACE_VERBOSE_INFO, (L"Servers column was found.\n"));
                }
                else
                {
                    // See if this is the Received node
                    if ((wcscmp(vtValue.bstrVal, bstrReceived.m_str) == 0))
                    {
                        // Mark that the From node exist and store its index.
                        bReceivedNodeExist = TRUE;
                        iReceivedNode = i;
                        DEBUG_TRACE(
                            TRACE_VERBOSE_INFO, 
                            (L"From column was found was column no. %d\n", iReceivedNode));
                    }
                }
                // After finding the heading child node, we do not have any intrest to 
                // scan the other child nodes.
                break;
            }
        }
    }

    // If the Servers column does not exist, build a new node for it
    if (!bServersColExist)
    {
        DEBUG_TRACE(TRACE_INFO, (L"Building the Servers column.\n"));

        CComPtr<IXMLDOMNode> spServersNode;

        size_t serversSchemaBuffSize = NELEMS(SCHEMA_PREFIX) + wcslen(bstrServers) + 1;
        g_wszServersSchema = new WCHAR[serversSchemaBuffSize];
        errno_t err = wcscpy_s(const_cast<LPWSTR>(g_wszServersSchema), serversSchemaBuffSize, SCHEMA_PREFIX);
        ATLASSERT(err == 0);
        err = wcscat_s(const_cast<LPWSTR>(g_wszServersSchema), serversSchemaBuffSize, bstrServers);
        ATLASSERT(err == 0);

        // this array defines the sub-nodes of the Amount node.
        NAME_VALUE_PAIR aNameValuePairs[] = {
            {L"heading", bstrServers.m_str},
            {L"prop", g_wszServersSchema},
            {L"type", L"string"},
            {L"width", L"68"},
            {L"style", L"padding-left:3px;text-align:left"}};

        // Build the Amount column node.
        hr = BuildColumnNode(spXMLDoc, aNameValuePairs, NELEMS(aNameValuePairs), &spServersNode);
        if (FAILED(hr))
        {
            DEBUG_TRACE(TRACE_ERRORS, (L"Failed in BuildColumnNode(\"Amount\"), hr = %x\n", hr));
            return hr;
        }

        // If we have found the "From" column, insert the Servers column right before it,
        // else insert it as the right most column.
        CComPtr<IXMLDOMNode> spReceivedColNode;

        if (bReceivedNodeExist)
        {
            hr = spColNodes->get_item(iReceivedNode, &spReceivedColNode);   
            if (FAILED(hr))
            {
                // If we failed to retrieve the "From" node, though it is very strage because we
                // crossed overv it while scanning the node, but nevertheless, we do not give up
                // and we still continue to insert the servers node as the left most column and hope 
                // for the best...
                DEBUG_TRACE(
                    TRACE_ERRORS, 
                    (L"Failed in spColNodes->get_item(iFromNode=%d), hr = %x\n", 
                    iReceivedNode, hr));
                spReceivedColNode = NULL;
            }
        }

        // Get the node of the last column. We will use it to insert the right most
        // columns.
        CComPtr<IXMLDOMNode> spLastColNode;

        hr = spColNodes->get_item(nColNodes - 1, &spLastColNode);   // -1 because it is 0 based.
        if (FAILED(hr))
        {
            DEBUG_TRACE(
                TRACE_ERRORS, 
                (L"Failed in spColNodes->get_item(nColNodes - 1=%d), hr = %x\n", 
                nColNodes - 1, hr));
            return hr;
        }

        // Get the parent node of all columns. It is required for inserting the new columns
        CComPtr<IXMLDOMNode> spParentNode;

        hr = spLastColNode->get_parentNode(&spParentNode);
        if (FAILED(hr))
        {
            DEBUG_TRACE(
                TRACE_ERRORS, 
                (L"Failed in spLastColNode->get_parentNode(), hr = %x\n", hr));
            return hr;
        }

        if (spReceivedColNode)
        {
            // Insert the Servers column before the From column.
            hr = spParentNode->insertBefore(spServersNode, CComVariant(spReceivedColNode), NULL);
            if (FAILED(hr))
            {
                DEBUG_TRACE(
                    TRACE_ERRORS, 
                    (L"Failed in spParentNode->insertBefore(spVCNode, "
                        L"CComVariant(spFromColNode), NULL), hr = %x\n", hr));
                return hr;
            }
            DEBUG_TRACE(TRACE_INFO, (L"Servers column was added before the from column.\n"));
        }
        else
        {
            // Get the next sibling of the node of the last column. We will use it to insert
            // the right most columns before it (we use the insertBefore method).
            CComPtr<IXMLDOMNode> spRefNode;

            hr = spLastColNode->get_nextSibling(&spRefNode);
            if (FAILED(hr))
            {
                DEBUG_TRACE(
                    TRACE_ERRORS, 
                    (L"Failed in spLastColNode->get_nextSibling(), hr = %x\n", hr));
                return hr;
            }

            // From column was not found, so we insert the Servers column at the right most end.
            if (spRefNode)
            {
                // Insert the new Servers column right before the sibling of the last column.
                hr = spParentNode->insertBefore(spServersNode, CComVariant(spRefNode), NULL);
                if (FAILED(hr))
                {
                    DEBUG_TRACE(
                        TRACE_ERRORS, 
                        (L"Failed in spParentNode->insertBefore(spVCNode, "
                            L"CComVariant(spRefNode), NULL), hr = %x\n", hr));
                    return hr;
                }
            }
            else
            {
                // In case the right most column have no sibling, we should call appendChild.
                hr = spParentNode->appendChild(spServersNode, NULL);
                if (FAILED(hr))
                {
                    DEBUG_TRACE(
                        TRACE_ERRORS, 
                        (L"Failed in spParentNode->appendChild(spVCNode, NULL), "
                            L"hr = %x\n", hr));
                    return hr;
                }
            }
            DEBUG_TRACE(TRACE_INFO, (L"Servers column was added as the right most column.\n"));
        }

        // At last, we can now modify the view.
        CComBSTR bstrXML;

        // Get the XML of the new view.
        hr = spXMLDoc->get_xml(&bstrXML);
        if (FAILED(hr))
        {
            DEBUG_TRACE(
                TRACE_ERRORS, 
                (L"Failed in spXMLDoc->get_xml(&bstrXML), hr = %x\n", hr));
            return hr;
        }
        // Store the new XML in the view object.
        hr = spView->put_XML(bstrXML);
        if (FAILED(hr))
        {
            DEBUG_TRACE(
                TRACE_ERRORS, 
                (L"Failed in spView->put_XML(bstrXML), hr = %x\n", hr));
            return hr;
        }
        // Commit the changes into the view object.
        hr = spView->Save();
        if (FAILED(hr))
        {
            DEBUG_TRACE(TRACE_ERRORS, (L"Failed in spView->Save(), hr = %x\n", hr));
            return hr;
        }
        DEBUG_TRACE(TRACE_INFO, (L"XML of the modified view was saved in Outlook view.\n"));
    }

    return S_OK;
}

