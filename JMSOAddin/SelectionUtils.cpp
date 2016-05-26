#include "StdAfx.h"
#include "SelectionUtils.h"
#include "debug.h"

SelectionUtils::SelectionUtils(Outlook::Selection *pSelection)
{
    m_spSelection = pSelection;
}


SelectionUtils::~SelectionUtils(void)
{
}

HRESULT SelectionUtils::EnumSelectionItems(EnumSelectionItemsProc enumProc, LPVOID lpParams)
{
    HRESULT hr = S_OK;
    long count;

    hr = m_spSelection->get_Count(&count);
    if (FAILED(hr))
    {
        DEBUG_TRACE(TRACE_ERRORS, (L"Failed in spSelection->get_Count(), error=0x%x\n", hr));
        return hr;
    }

    for (long i = 0; i < count; i++)
    {
        CComPtr<IDispatch> spDisp;

        hr = m_spSelection->Item(CComVariant(i + 1), &spDisp);
        if (FAILED(hr))
        {
            DEBUG_TRACE(TRACE_ERRORS, (L"Failed in spSelection->Item(%d), error=0x%x\n", i, hr));
            return hr;
        }

        hr = enumProc(spDisp, lpParams);
        if (FAILED(hr))
        {
            DEBUG_TRACE(TRACE_ERRORS, (L"Failed in enumProc(%d), error=0x%x\n", i, hr));
            return hr;
        }
    }

    return S_OK;
}
