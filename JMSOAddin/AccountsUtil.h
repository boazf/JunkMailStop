#pragma once
class CAccountsUtil
{
public:
    CAccountsUtil(void);
    ~CAccountsUtil(void);

public: 
    typedef HRESULT (*AccountEnumProc)(Outlook::_Account *lpAccount, BOOL *lpbCont, LPVOID lpvParam);

public:
    static HRESULT EnumAccounts(AccountEnumProc EnumProc, LPVOID lpvParam);
    static void SetApplication(CComPtr<Outlook::_Application> spApp);

private:
    static CComPtr<Outlook::_Application> m_spApp;
};

