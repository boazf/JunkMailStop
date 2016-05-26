#pragma once
#include "string"
extern bool operator<(const wstring &s1, const wstring &s2);
#include "hash_set"
#include "CritSect.h"

typedef hash_set<wstring> BlackList;

#define MAX_SERVER_NAME_LEN 127

// This class is used to store and retrieve the configuration parameters of
// Junk Mail Stop Outlook Addin.
class CJMSOAddinConfig
{
public:
    CJMSOAddinConfig();
    ~CJMSOAddinConfig();
    void GetBlackList(BlackList &blacklist);
    ULONG SetBlackList(const BlackList &blackList);
    ULONG AddServersToBlackList(const BlackList &blackList);
    ULONG RemoveServersFromBlackList(const BlackList &blackList);
    bool IsInitError(void) const;

private:
    CRegKey m_RegKey;               // Registry key for the configuration store
    BlackList m_blackList;
    CHandle m_hExitMonitor;
    bool m_isInitError;
    CCritSect m_cs;

private:
    static DWORD WINAPI ConfigMonitor(LPVOID lpParam);
    ULONG GetBlackListRegKey(CRegKey &blackListRegKey);
    void InitializeBlackList(void);
    ULONG SetBlackServerValue(CRegKey &regKey, LPCWSTR szServerName, DWORD dwValue=0);
    ULONG DeleteBlackServerValue(CRegKey &regKey, LPCWSTR szServerName);
};

// A global configuration class instance.
extern CJMSOAddinConfig *g_pCConfig;