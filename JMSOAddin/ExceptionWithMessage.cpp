#include "StdAfx.h"
#include "ExceptionWithMessage.h"


CExceptionWithMessage::CExceptionWithMessage(UINT uiId, ...)
{
    va_list args;

    // Format the trace string.
    va_start(args, uiId);
    CStringW szFormat;
    szFormat.LoadString(uiId);
    m_errorMessage.FormatMessageV(szFormat, &args);
}


CExceptionWithMessage::~CExceptionWithMessage(void)
{
}

LPCWSTR CExceptionWithMessage::GetErrorMessage(void)
{
    return m_errorMessage.GetString();
}