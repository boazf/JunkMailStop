#pragma once
#include "AddinMessageBox.h"
#include "ServersDetectionProgressDlg.h"

DWORD WINAPI ProgressDialogThreadProc(LPVOID lpParameter);

template<class Collection>
class ServersDetection
{
public:
    ServersDetection(void)
    {
    }

    ~ServersDetection(void)
    {
    }

public:
    void DetectServers(Collection *pCollection)
    {
        HRESULT hr;
        DWORD dwThreadId;
        CHandle hSyncEvent(CreateEvent(NULL, TRUE, FALSE, NULL));
        m_progressDlg.SetSyncEvent(hSyncEvent);
        HANDLE hThread = CreateThread(NULL, 0, ProgressDialogThreadProc, &m_progressDlg, 0, &dwThreadId);
        if (hThread == NULL)
        {
            DEBUG_TRACE(TRACE_ERRORS, (L"Failed to create thread for the progress dialog, error=%d\n", GetLastError()));
        }
        else
        {
            CloseHandle(hThread);
        }

        long nItems;

        hr = pCollection->get_Count(&nItems);
        if (FAILED(hr))
        {
            DEBUG_TRACE(TRACE_ERRORS, (L"Failed in spItems->get_Count(), error=0x%x\n", hr));
            ShowErrorMessageBox();
            return;
        }

        WaitForSingleObject(hSyncEvent, INFINITE);
        m_progressDlg.SetProgressRange(nItems);
        for (int iItem = 0; iItem < nItems && !m_progressDlg.ShouldCancel(); iItem++)
        {
            CComPtr<IDispatch> spItem;
            HRESULT hr = pCollection->Item(CComVariant(iItem + 1), &spItem);
            if (FAILED(hr))
            {
                DEBUG_TRACE(TRACE_ERRORS, (L"Failed in spItems->Item(%d), error=0x%x\n", iItem, hr));
                ShowErrorMessageBox();
                break;
            }

            CComQIPtr<Outlook::_MailItem> spMailItem(spItem);
            if (spMailItem != NULL)
            {
                ServerNameParser parser(spMailItem);

                CComBSTR bstrServers;
                hr = parser.GetServers(bstrServers);
                if (FAILED(hr))
                {
                    DEBUG_TRACE(TRACE_ERRORS, (L"Failed in parser.GetServers, error=0x%x\n", hr));
                    ShowErrorMessageBox();
                    break;
                }

                hr = CItemsEvents::SetServersColumn(spMailItem, bstrServers);
                if (FAILED(hr))
                {
                    DEBUG_TRACE(TRACE_ERRORS, (L"Failed in CItemsEvents::SetServersColumn, error=0x%x\n", hr));
                    ShowErrorMessageBox();
                    break;
                }
            }
            else
            {
                DEBUG_TRACE(TRACE_INFO, (L"Skipping non-mail item\n"));
            }

            m_progressDlg.SetPosition(iItem);
            CStringW strPart;
            strPart.FormatMessageW(L"%1!u! / %2!u!        ", iItem, nItems);
            m_progressDlg.SetMessage(strPart.GetString());
        }
        m_progressDlg.EndDialog(IDOK);
    }

    static void ShowErrorMessageBox(void)
    {
        AddinMessageBox(IDS_DETECT_SERVERS_ERROR, MB_ICONERROR);
    }


private:
    ServersDetectionProgressDlg m_progressDlg;
};

