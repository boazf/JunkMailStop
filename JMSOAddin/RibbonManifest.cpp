#include "StdAfx.h"
#include "RibbonManifest.h"
#include "debug.h"

#define CONTEXT_MENU_MAIL_ITEM_ADD_SERVERS          L"ContextMenuMailItemAddServers"
#define CONTEXT_MENU_MULTIPLE_ITEMS_ADD_SERVERS     L"ContextMenuMultipleItemsAddServers"
#define CONTEXT_MENU_MAIL_ITEM_DETECT_SERVER        L"ContextMenuMailItemDetectServers"
#define CONTEXT_MENU_MULTIPLE_ITEMS_DETECT_SERVERS  L"ContextMenuMultipleItemsDetectServers"
#define CONTEXT_MENU_FOLDER_DETECT_SERVERS          L"ContextMenuFolderDetectServers"


typedef struct _StringIntPair
{
    LPCWSTR str;
    int i;
} StringIntPair;

static StringIntPair stringIntPairs[] =
{
    {CONTEXT_MENU_MAIL_ITEM_ADD_SERVERS,            ID_CONTEXT_MENU_MAIL_ITEM_ADD_SERVERS},
    {CONTEXT_MENU_MULTIPLE_ITEMS_ADD_SERVERS,       ID_CONTEXT_MENU_MULTIPLE_ITEMS_ADD_SERVERS},
    {CONTEXT_MENU_MAIL_ITEM_DETECT_SERVER,          ID_CONTEXT_MENU_MAIL_ITEM_DETECT_SERVER},
    {CONTEXT_MENU_MULTIPLE_ITEMS_DETECT_SERVERS,    ID_CONTEXT_MENU_MULTIPLE_ITEMS_DETECT_SERVERS},
    {CONTEXT_MENU_FOLDER_DETECT_SERVERS,            ID_CONTEXT_MENU_FOLDER_DETECT_SERVERS}
};

RibbonManifest::RibbonManifest(void)
{
}


RibbonManifest::~RibbonManifest(void)
{
}

RibbonManifest::StringToIntMap RibbonManifest::m_stringToIntMap;

static HRESULT GetResource(int nId, LPCWSTR lpType, LPVOID* ppvResourceData, DWORD* pdwSizeInBytes)
{
    HMODULE hModule = _AtlBaseModule.GetModuleInstance();
    if (!hModule)
    {
        DEBUG_TRACE(TRACE_ERRORS, (L"_AtlBaseModule.GetModuleInstance() returned NNULL\n"));
        return E_UNEXPECTED;
    }

    HRSRC hRsrc = FindResource(hModule, MAKEINTRESOURCE(nId), lpType);
    if (!hRsrc)
    {
        DEBUG_TRACE(TRACE_ERRORS, (L"Failed in FindResourc()e, error=%d\n", GetLastError()));
        return HRESULT_FROM_WIN32(GetLastError());
    }

    HGLOBAL hGlobal = LoadResource(hModule, hRsrc);
    if (!hGlobal)
    {
        DEBUG_TRACE(TRACE_ERRORS, (L"Failed in LoadResource(), error=%d\n", GetLastError()));
        return HRESULT_FROM_WIN32(GetLastError());
    }

    *pdwSizeInBytes = SizeofResource(hModule, hRsrc);
    *ppvResourceData = LockResource(hGlobal);

    return S_OK;
}

static BSTR GetXMLResource(int nId)
{
    LPVOID pResourceData = NULL;
    DWORD dwSizeInBytes = 0;
    HRESULT hr = GetResource(nId, L"XML", &pResourceData, &dwSizeInBytes);
    if (FAILED(hr))
    {
        DEBUG_TRACE(TRACE_ERRORS, (L"Failed in GetResource(), error=0x%x\n", hr));
        return NULL;
    }

    // Assumes that the data is not stored in Unicode.
    CComBSTR cbstr(dwSizeInBytes, reinterpret_cast<LPCSTR>(pResourceData));
    return cbstr.Detach();
}

HRESULT RibbonManifest::GetRibbonManifest(BSTR * bstrRibbonXml)
{
    if(!bstrRibbonXml)
    {
        DEBUG_TRACE(TRACE_ERRORS, (L"Received NULL parameter in bstrRibbonXml\n"));
        return E_POINTER;
    }

    // Localize the XML
    CComBSTR bstrRibbonResource(GetXMLResource(IDR_XML_RIBBON_MANIFEST));
    CComBSTR bstrMailItemLabel, bstrFolderLabel;
    bstrMailItemLabel.LoadStringW(IDS_ADD_BLACK_SERVERS);
    bstrFolderLabel.LoadStringW(IDS_DETECT_SERVERS);
    CStringW strRibbonXml;
    strRibbonXml.FormatMessageW(
        bstrRibbonResource.m_str, 
        bstrMailItemLabel.m_str, 
        bstrFolderLabel.m_str,
        CONTEXT_MENU_MAIL_ITEM_ADD_SERVERS,
        CONTEXT_MENU_MULTIPLE_ITEMS_ADD_SERVERS,
        CONTEXT_MENU_MAIL_ITEM_DETECT_SERVER,
        CONTEXT_MENU_MULTIPLE_ITEMS_DETECT_SERVERS,
        CONTEXT_MENU_FOLDER_DETECT_SERVERS);
    *bstrRibbonXml = CComBSTR(strRibbonXml.GetString()).Detach();

    return S_OK;
}

int RibbonManifest::GetMenuId(LPCWSTR id)
{
    if (m_stringToIntMap.size() == 0)
    {
        for (int i = 0; i < NELEMS(stringIntPairs); i++)
        {
            m_stringToIntMap[stringIntPairs[i].str] = stringIntPairs[i].i;
        }
    }
    return m_stringToIntMap[wstring(id)];
}


