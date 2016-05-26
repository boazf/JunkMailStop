#pragma once
#include "string"
extern bool operator<(const wstring &s1, const wstring &s2);
#include "map"

#define ID_CONTEXT_MENU_MAIL_ITEM_ADD_SERVERS           1
#define ID_CONTEXT_MENU_MULTIPLE_ITEMS_ADD_SERVERS      2
#define ID_CONTEXT_MENU_MAIL_ITEM_DETECT_SERVER         3
#define ID_CONTEXT_MENU_MULTIPLE_ITEMS_DETECT_SERVERS   4
#define ID_CONTEXT_MENU_FOLDER_DETECT_SERVERS           5

class RibbonManifest
{
public:
    RibbonManifest(void);
    ~RibbonManifest(void);

public:
    static HRESULT GetRibbonManifest(BSTR *bstrRibbonXml);
    static int GetMenuId(LPCWSTR id);

private:
    typedef map<wstring, int> StringToIntMap;

private:
    static StringToIntMap m_stringToIntMap;
};

