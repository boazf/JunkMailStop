#pragma once
class SelectionUtils
{
public:
    SelectionUtils(Outlook::Selection *pSelection);
    ~SelectionUtils(void);

public:
    typedef HRESULT (*EnumSelectionItemsProc)(IDispatch *pDisp, LPVOID lpParams);

public:
    HRESULT EnumSelectionItems(EnumSelectionItemsProc enumProc, LPVOID lpParams);

private:
    CComPtr<Outlook::Selection> m_spSelection;
};

