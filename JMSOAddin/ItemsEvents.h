// ItemsEvents.h : Declaration of the CItemsEvents

#pragma once
#include "resource.h"       // main symbols
#include "JMSOAddin.h"



#if defined(_WIN32_WCE) && !defined(_CE_DCOM) && !defined(_CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA)
#error "Single-threaded COM objects are not properly supported on Windows CE platform, such as the Windows Mobile platforms that do not include full DCOM support. Define _CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA to force ATL to support creating single-thread COM object's and allow use of it's single-threaded COM object implementations. The threading model in your rgs file was set to 'Free' as that is the only threading model supported in non DCOM Windows CE platforms."
#endif

using namespace ATL;

extern _ATL_FUNC_INFO OnItemAddEventInfo;

#define ITEMS_EVENTS_ID         3

// CItemsEvents

class ATL_NO_VTABLE CItemsEvents :
    public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CItemsEvents, &CLSID_ItemsEvents>,
	public IDispatchImpl<IItemsEvents, &IID_IItemsEvents, &LIBID_JMSOAddinLib, /*wMajor =*/ 1, /*wMinor =*/ 0>,
    public IDispEventSimpleImpl<ITEMS_EVENTS_ID,CItemsEvents,&__uuidof(Outlook::ItemsEvents)>
{
public:
	CItemsEvents()
	{
	}

	~CItemsEvents()
	{
	}

    DECLARE_REGISTRY_RESOURCEID(IDR_ITEMSEVENTS)


    BEGIN_COM_MAP(CItemsEvents)
	    COM_INTERFACE_ENTRY(IItemsEvents)
	    COM_INTERFACE_ENTRY(IDispatch)
    END_COM_MAP()

    //ATL Sink Message MAP
    BEGIN_SINK_MAP(CItemsEvents)
        SINK_ENTRY_INFO(
            ITEMS_EVENTS_ID,
            __uuidof(Outlook::ItemsEvents),
            /*ItemAdd*/0xf001,
            OnItemAdd,
            &OnItemAddEventInfo)
    END_SINK_MAP()

    DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct()
	{
		return S_OK;
	}

	void FinalRelease()
	{
	}

public:
    HRESULT set_JunkFolder(Outlook::MAPIFolder *pFolder);
    static HRESULT SetServersColumn(Outlook::_MailItem *pMailItem, BSTR bstrServers);

protected:
    void __stdcall OnItemAdd(IDispatch *pIDisp);

private:
    CComPtr<Outlook::MAPIFolder> m_spJunkFolder;
};

OBJECT_ENTRY_AUTO(__uuidof(ItemsEvents), CItemsEvents)
