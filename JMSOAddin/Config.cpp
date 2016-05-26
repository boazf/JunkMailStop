// This file contains the implementation of the configuration repository of the
// Junk Mail Stop Outlook  Addin. The repository is implemented on top of the registry.

#include "stdafx.h"
#include "Config.h"
#include "debug.h"
#include "resource.h"
#include "ExceptionWithMessage.h"
#include "AddinMessageBox.h"

// the registry path where the configuration resides.
#define VCOA_CONFIG_REGISTRY_PATH       L"Software\\JunkStop\\Outlook Addin"
#define BLACK_LIST_KEY_NAME             L"BlackList"

DWORD WINAPI CJMSOAddinConfig::ConfigMonitor(LPVOID lpParam)
{
    DEBUG_TRACE(TRACE_VERBOSE_INFO, (L"Started configuration monitor thread.\n"));

    CJMSOAddinConfig *pConfig = static_cast<CJMSOAddinConfig *>(lpParam);
    CRegKey regBlackList;
    ULONG ulRes = pConfig->GetBlackListRegKey(regBlackList);
    if (ulRes != ERROR_SUCCESS)
    {
        DEBUG_TRACE(TRACE_ERRORS, (L"Failed to retrieve the black list registry key, error=%d\n", GetLastError()));
        return ulRes;
    }
    CHandle hRegChangeEvent(CreateEvent(NULL, FALSE, FALSE, NULL));
    HANDLE handles[] = {hRegChangeEvent.m_h, pConfig->m_hExitMonitor.m_h};
    bool bExit = false;
    do
    {
        LSTATUS lStatus = RegNotifyChangeKeyValue(regBlackList.m_hKey, TRUE, REG_NOTIFY_CHANGE_LAST_SET, hRegChangeEvent.m_h, TRUE);
        if (lStatus != ERROR_SUCCESS)
        {
            DEBUG_TRACE(TRACE_ERRORS, (L"Failed in RegNotifyChangeKeyValue(), error=%d\n", GetLastError()));
            return 0;
        }
        DWORD triggeredEvent = WaitForMultipleObjects(2, handles, FALSE, INFINITE);
        switch (triggeredEvent)
        {
        case WAIT_OBJECT_0:
            try
            {
                pConfig->InitializeBlackList();
            }
            catch (CExceptionWithMessage *e)
            {
                DEBUG_TRACE(TRACE_ERRORS, (e->GetErrorMessage()));
                delete e;
                pConfig->m_blackList.clear();
                pConfig->m_isInitError = true;
            }
            break;

        case WAIT_OBJECT_0 + 1:
            bExit = true;
            break;

        default:
            DEBUG_TRACE(TRACE_ERRORS, (L"Failed in WaitForMultipleObjects, reCode=%d, error=%d\n", triggeredEvent, GetLastError()));
            bExit = true;
            break;
        }
    } while (!bExit);

    return 0;
}

// Configuration erpository class constructor. It sets all values to default values
// and creates/opens the configuration registry key.
CJMSOAddinConfig::CJMSOAddinConfig()
{
    try
    {
        // Create/open the configuration registry key.
        ULONG ulRes;

        ulRes = m_RegKey.Create(HKEY_CURRENT_USER, VCOA_CONFIG_REGISTRY_PATH);
        if (ulRes != ERROR_SUCCESS)
        {
            DEBUG_TRACE(
                TRACE_ERRORS,
                (L"Failed in m_RegKey.Create(HKEY_CURRENT_USER, %s), Error = %d\n", VCOA_CONFIG_REGISTRY_PATH , ulRes));
            throw new CExceptionWithMessage(IDS_REG_OPEN_ROOT_ERROR, VCOA_CONFIG_REGISTRY_PATH, ulRes);
        }

        InitializeBlackList();

        DWORD dwThreadId;
        m_hExitMonitor.Attach(CreateEvent(NULL, FALSE, FALSE, NULL));
        CHandle hThread(CreateThread(NULL, 0, ConfigMonitor, this, 0, &dwThreadId));
        if (hThread == NULL)
        {
            DEBUG_TRACE(TRACE_ERRORS, (L"Failed to create registroy monitor thread, error=%d\n", GetLastError()));
            throw new CExceptionWithMessage(IDS_MONITOR_THREAD_ERROR, ulRes);
        }
    } 
    catch (CExceptionWithMessage *e)
    {
        // :-(
        // Show the error dialog box.
        AddinMessageBoxOnInit(e->GetErrorMessage(), MB_ICONERROR);
        delete e;
        m_blackList.clear();
        m_isInitError = true;
    }
}

CJMSOAddinConfig::~CJMSOAddinConfig()
{
    // Stop the monitor thread.
    SetEvent(m_hExitMonitor.m_h);
}

void CJMSOAddinConfig::InitializeBlackList(void)
{
    CLock lock(&m_cs);

    // Get the black list's registry key
    CRegKey blackListRegKey;
    ULONG ulRes = GetBlackListRegKey(blackListRegKey);
    if (ulRes != ERROR_SUCCESS)
    {
        DEBUG_TRACE(
            TRACE_ERRORS,
            (L"Failed in GetBlackListRegKey(), Error = %d\n", ulRes));
        throw new CExceptionWithMessage(IDS_REG_OPEN_BLACK_LIST_ERROR, VCOA_CONFIG_REGISTRY_PATH, BLACK_LIST_KEY_NAME, ulRes);
    }

    // Query information about the black list data
    DWORD dwValues, dwMaxValueNameLen, dwMaxValueLen;

    ulRes = RegQueryInfoKey(blackListRegKey.m_hKey, NULL, NULL, NULL, NULL, NULL, NULL, &dwValues, &dwMaxValueNameLen, &dwMaxValueLen, NULL, NULL);
    if (ulRes != ERROR_SUCCESS)
    {
        DEBUG_TRACE(
            TRACE_ERRORS,
            (L"Failed in RegQueryInfoKey, Error = %d\n", ulRes));
        throw new CExceptionWithMessage(IDS_REG_QUERY_BLACK_LIST_ERROR, VCOA_CONFIG_REGISTRY_PATH, BLACK_LIST_KEY_NAME, ulRes);
    }

    // Enumerate the black list's servers
    ATLASSERT(dwValues == 0 || dwMaxValueLen == sizeof(DWORD));
    CAutoPtr<WCHAR> szValueName;
    szValueName.Attach(new WCHAR[dwMaxValueNameLen + 1]);
    m_blackList.clear();
    for (DWORD dwIndex = 0; dwIndex < dwValues; dwIndex++)
    {
        DWORD dwValueNameLen = dwMaxValueNameLen + 1;
        DWORD dwType;
        DWORD dwValue;
        DWORD dwValueLen = sizeof(DWORD);

        ulRes = RegEnumValue(blackListRegKey.m_hKey, dwIndex, szValueName, &dwValueNameLen, NULL, &dwType, (LPBYTE)&dwValue, &dwValueLen);
        if (ulRes != ERROR_SUCCESS)
        {
            DEBUG_TRACE(
                TRACE_ERRORS,
                (L"Failed in RegEnumValue, Error = %d\n", ulRes));
            throw new CExceptionWithMessage(IDS_REG_GET_BLACK_LST_VALUE_ERROR, VCOA_CONFIG_REGISTRY_PATH, BLACK_LIST_KEY_NAME, dwIndex, ulRes);
        }
        ATLASSERT(dwType == REG_DWORD);
        ATLASSERT(dwValueNameLen <= dwMaxValueNameLen);
        ATLASSERT(dwValueLen == dwMaxValueLen);
        m_blackList.insert(wstring(szValueName));
    }

    m_isInitError = false;
}

// Retrieve the black list's registry key
ULONG CJMSOAddinConfig::GetBlackListRegKey(CRegKey &blackListRegKey)
{
    ULONG ulRes = blackListRegKey.Create(m_RegKey.m_hKey, BLACK_LIST_KEY_NAME);
    if (ulRes != ERROR_SUCCESS)
    {
        DEBUG_TRACE(
            TRACE_ERRORS,
            (L"Failed in m_blackListRegKey.Create(%s), Error = %d\n", BLACK_LIST_KEY_NAME, ulRes));
    }

    return ulRes;
}

void CJMSOAddinConfig::GetBlackList(BlackList &blacklist)
{
    CLock lock(&m_cs);
    blacklist = m_blackList;
}

ULONG CJMSOAddinConfig::DeleteBlackServerValue(CRegKey &blackListRegKey, LPCWSTR szServerName)
{
    ULONG ulRes = blackListRegKey.DeleteValue(szServerName);
    if (ulRes != ERROR_SUCCESS)
    {
        DEBUG_TRACE(TRACE_ERRORS, (L"Failed in blackListRegKey.DeleteValue(%s), error=%d\n", szServerName, ulRes));
    }
    return ulRes;
}

ULONG CJMSOAddinConfig::SetBlackServerValue(CRegKey &blackListRegKey, LPCWSTR szServerName, DWORD dwValue)
{
    ULONG ulRes = blackListRegKey.SetDWORDValue(szServerName, dwValue);
    if (ulRes != ERROR_SUCCESS)
    {
        DEBUG_TRACE(TRACE_ERRORS, (L"Failed in blackListRegKey.SetDWORDValue(%s), error=%d\n", szServerName, ulRes));
    }
    return ulRes;
}

static void SubtractBlackLists(BlackList &res, const BlackList l1, const BlackList l2)
{
    res = l1;
    for (BlackList::iterator i = l2.cbegin(); i != l2.cend(); i++)
    {
        res.erase(*i);
    }
}

// Set a new black list and store it in the registry.
ULONG CJMSOAddinConfig::SetBlackList(const BlackList &blackList)
{
    CLock lock(&m_cs);
    BlackList serversToRemove;
    SubtractBlackLists(serversToRemove, m_blackList, blackList);
    ULONG ulRes = RemoveServersFromBlackList(serversToRemove);
    if (ulRes != ERROR_SUCCESS)
    {
        DEBUG_TRACE(TRACE_ERRORS, (L"Failed in RemoveServersFromBlackList(), error=%d\n", ulRes));
        return ulRes;
    }

    BlackList serversToAdd;
    SubtractBlackLists(serversToAdd, blackList, m_blackList);
    ulRes = AddServersToBlackList(serversToAdd);
    if (ulRes != ERROR_SUCCESS)
    {
        DEBUG_TRACE(TRACE_ERRORS, (L"Failed in AddServersToBlackList(), error=%d\n", ulRes));
        return ulRes;
    }

    return ERROR_SUCCESS;
}

ULONG CJMSOAddinConfig::AddServersToBlackList(const BlackList &blackList)
{
    CLock lock(&m_cs);

    if (blackList.size() == 0)
    {
        return ERROR_SUCCESS;
    }

    CRegKey blackListRegKey;
    ULONG ulRes = GetBlackListRegKey(blackListRegKey);
    if (ulRes != ERROR_SUCCESS)
    {
        DEBUG_TRACE(TRACE_ERRORS, (L"Failed in GetBlackListRegKey(), error=%d\n", ulRes));
        return ulRes;
    }
    for (BlackList::const_iterator i = blackList.cbegin();
         i != blackList.cend();
         i++)
    {
        if (m_blackList.find(*i) == m_blackList.end())
        {
            ulRes = SetBlackServerValue(blackListRegKey, (*i).data());
            if (ulRes != ERROR_SUCCESS)
            {
                DEBUG_TRACE(TRACE_ERRORS, (L"Failed in SetBlackServerValue(), error=%d\n", ulRes));
                return ulRes;
            }
        }
    }

    return ERROR_SUCCESS;
}

ULONG CJMSOAddinConfig::RemoveServersFromBlackList(const BlackList &blackList)
{
    CLock lock(&m_cs);

    if (blackList.size() == 0)
    {
        return ERROR_SUCCESS;
    }

    CRegKey blackListRegKey;
    ULONG ulRes = GetBlackListRegKey(blackListRegKey);
    if (ulRes != ERROR_SUCCESS)
    {
        DEBUG_TRACE(TRACE_ERRORS, (L"Failed in GetBlackListRegKey(), error=%d\n", ulRes));
        return ulRes;
    }
    for (BlackList::const_iterator i = blackList.cbegin();
         i != blackList.cend();
         i++)
    {
        if (m_blackList.find(*i) != m_blackList.end())
        {
            DeleteBlackServerValue(blackListRegKey, (*i).data());
            if (ulRes != ERROR_SUCCESS)
            {
                DEBUG_TRACE(TRACE_ERRORS, (L"Failed in DeleteBlackServerValue(), error=%d\n", ulRes));
                return ulRes;
            }
        }
    }

    return ERROR_SUCCESS;
}

// Indicates whether the configuration was initialized successfuly.
bool CJMSOAddinConfig::IsInitError(void) const
{
    return m_isInitError;
}

// A global instance of the configuration repository class.
CJMSOAddinConfig *g_pCConfig;

// This operator is required by the hash_set data structure that stores the black list's server names
bool operator<(const wstring &s1, const wstring &s2)
{
    return s1.compare(s2) < 0;
}
