#pragma once
class CExceptionWithMessage
{
public:
    CExceptionWithMessage(UINT uiId, ...);
    ~CExceptionWithMessage(void);
    LPCWSTR GetErrorMessage();

private:
    CStringW m_errorMessage;
};

