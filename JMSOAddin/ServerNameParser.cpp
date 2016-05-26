#include "StdAfx.h"
#include "ServerNameParser.h"
#include "debug.h"
#include "string"


ServerNameParser::ServerNameParser(Outlook::_MailItem *pMailItem) : 
    m_bIsInitialized(false),
    m_spMailItem(pMailItem)
{
}

#define RECEIVED_HEADER L"\r\nReceived: from "

#ifdef _DEBUG
#define CHECKERRNO(x) { errno_t err = x; ATLASSERT(err == 0); }
#else
#define CHECKERRNO(x) x;
#endif

HRESULT ServerNameParser::Init(void)
{
    ATLASSERT(!m_bIsInitialized);

    HRESULT hr;
    CComPtr<Outlook::_PropertyAccessor> spPropertyAccessor;

    hr = m_spMailItem->get_PropertyAccessor(&spPropertyAccessor);
    if (FAILED(hr))
    {
        DEBUG_TRACE(TRACE_ERRORS, (L"Failed in spMailItem->get_PropertyAccessor(), hr=%x", hr));
        return hr;
    }

    CComVariant headers;

    hr = spPropertyAccessor->GetProperty(L"http://schemas.microsoft.com/mapi/proptag/0x007D001E", &headers);
    if (FAILED(hr))
    {
        DEBUG_TRACE(TRACE_ERRORS, (L"Failed in spPropertyAccessor->GetProperty(), hr=%x", hr));
        return hr;
    }

    DEBUG_TRACE(TRACE_VERBOSE_INFO, (L"Headers: %s\n", headers.bstrVal));

    LPWSTR lpszHeaders = headers.bstrVal;

    do
    {
        lpszHeaders = wcsstr(lpszHeaders, RECEIVED_HEADER);
        if (lpszHeaders == NULL)
        {
            break;
        }
        lpszHeaders += NELEMS(RECEIVED_HEADER) - 1;
        LPWSTR lpEol = wcsstr(lpszHeaders, L"\r\n");
        CAutoPtr<WCHAR> lpReceivedLine;
        if (lpEol != NULL)
        {
            DWORD dwLineLen = lpEol - lpszHeaders;
            lpReceivedLine.Attach(new WCHAR[dwLineLen + 1]);
            CHECKERRNO(wcsncpy_s(lpReceivedLine, dwLineLen + 1, lpszHeaders, dwLineLen));
            lpReceivedLine[dwLineLen] = L'\0';
            lpszHeaders = lpEol;
        }
        else
        {
            DWORD dwLineLen = wcslen(lpszHeaders);
            lpReceivedLine.Attach(new WCHAR[dwLineLen + 1]);
            CHECKERRNO(wcscpy_s(lpReceivedLine, dwLineLen + 1, lpszHeaders));
            lpszHeaders = NULL;
        }

        LPWSTR lpServerName = lpReceivedLine;

        DEBUG_TRACE(TRACE_VERBOSE_INFO, (L"Receved from line: %s\n", lpReceivedLine));
        LPCWSTR szOpeningBrackets = L"([";
        LPCWSTR szClosingBrackets = L")]";
        if (wcschr(szOpeningBrackets, *lpServerName) != NULL)
        {
            int i = 0;
            while (szOpeningBrackets[i] != '\0')
            {
                if (*lpServerName == szOpeningBrackets[i])
                {
                    lpServerName += 1;
                    LPWSTR lpSep = wcschr(lpServerName, szClosingBrackets[i]);
                    if (lpSep == NULL)
                    {
                        continue;
                    }
                    *lpSep = L'\0';
                }
                i++;
            }
            LPCWSTR aszLocalHosts[] = { L"10.", L"127.", L"192,168." };
            bool bLocalhost = false;
            for (i = 0; i < NELEMS(aszLocalHosts) && !bLocalhost; i++)
            {
                bLocalhost = wcsncmp(lpServerName, aszLocalHosts[i], wcslen(aszLocalHosts[i])) == 0;
            }
            if (bLocalhost)
            {
                continue;
            }
        }
        else
        {
            LPWSTR lpSep = wcspbrk(lpServerName, L" \t");
            if (lpSep != NULL)
            {
                *lpSep = L'\0';
            }
        }
        wstring strServerName(lpServerName);
        for (wstring::iterator i = strServerName.begin();
                i != strServerName.end();
                i++)
        {
            *i = static_cast<WCHAR>(tolower(*i));
        }
        if (strServerName == L"localhost")
        {
            continue;
        }
        m_ServersList.AddHead(strServerName);
    } while(lpszHeaders != NULL);

    m_spMailItem.Release();
    m_bIsInitialized = true;

    return S_OK;
}


ServerNameParser::~ServerNameParser(void)
{
}

HRESULT ServerNameParser::GetServer(CComBSTR &bstrServerName)
{
    HRESULT hr;

    if (!m_bIsInitialized)
    {
        hr = Init();
        if (FAILED (hr))
        {
            DEBUG_TRACE(TRACE_ERRORS, (L"Failed in Init(), hr = 0x%x\n", hr));
            return hr;
        }
    }

    if (m_ServersList.IsEmpty())
    {
        bstrServerName.Attach(NULL);
    }
    else
    {
        bstrServerName = CComBSTR(m_ServersList.RemoveHead().data());
    }

    return S_OK;
}

HRESULT ServerNameParser::GetServers(CComBSTR &bstrServers)
{
    bstrServers = L"";
    DWORD nServers = 0;

    do
    {
        CComBSTR bstrServerName;

        GetServer(bstrServerName);
        if (bstrServerName == NULL)
        {
            break;
        }

        nServers++;
        if (nServers > 1)
        {
            bstrServers.Append("; ");
        }
        bstrServers.Append(bstrServerName);
    } while (true);

    return S_OK;
}