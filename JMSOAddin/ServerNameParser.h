#pragma once

#include "string"

class ServerNameParser
{
public:
    ServerNameParser(Outlook::_MailItem *pMailItem);
    ~ServerNameParser(void);

public:
    HRESULT GetServer(CComBSTR &bstrServerName);
    HRESULT GetServers(CComBSTR &bstrServers);

private:
    HRESULT Init(void);

private:
    CComPtr<Outlook::_MailItem> m_spMailItem;
    BOOL m_bIsInitialized;
    CAtlList<wstring> m_ServersList;
};

