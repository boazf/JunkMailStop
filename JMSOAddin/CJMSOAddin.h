// CJMSOAddin.h : Declaration of the CJMSOAddin

#pragma once
#include "resource.h"       // main symbols

#include "JMSOAddin.h"
#include "debug.h"
#include "ItemsEvents.h"

#if defined(_WIN32_WCE) && !defined(_CE_DCOM) && !defined(_CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA)
#error "Single-threaded COM objects are not properly supported on Windows CE platform, such as the Windows Mobile platforms that do not include full DCOM support. Define _CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA to force ATL to support creating single-thread COM object's and allow use of it's single-threaded COM object implementations. The threading model in your rgs file was set to 'Free' as that is the only threading model supported in non DCOM Windows CE platforms."
#endif

extern _ATL_FUNC_INFO OnSimpleEventInfo;
extern _ATL_FUNC_INFO OnOptionsAddPagesInfo;

#define EXPLORER_EVENTS_ID      1
#define APPLICATION_EVENTS_ID   2
#define ITEMS_EVENTS_ID         3

// CJMSOAddin

class ATL_NO_VTABLE CJMSOAddin :
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<CJMSOAddin, &CLSID_JMSOAddin>,
    public ISupportErrorInfo,
    public IDispatchImpl<IJMSOAddin, &IID_IJMSOAddin, &LIBID_JMSOAddinLib, /*wMajor =*/ 1, /*wMinor =*/ 0>,
    public IDispatchImpl<_IDTExtensibility2, &__uuidof(_IDTExtensibility2), &LIBID_AddInDesignerObjects, /* wMajor = */ 1>,
    public IDispatchImpl<Office::IRibbonExtensibility, &__uuidof(Office::IRibbonExtensibility), &Office::LIBID_Office, /* wMajor = */ 2, /* wMinor = */ 5>,
    public IDispatchImpl<IRibbonCallback, &__uuidof(IRibbonCallback), &LIBID_JMSOAddinLib, /*wMajor =*/ 1, /*wMinor =*/ 0>,
    public IDispEventSimpleImpl<APPLICATION_EVENTS_ID,CJMSOAddin,&__uuidof(Outlook::ApplicationEvents_10)>
{
public:
    CJMSOAddin()
    {
    }

    ~CJMSOAddin()
    {
    }

    DECLARE_REGISTRY_RESOURCEID(IDR_JMSOADDIN1)


    BEGIN_COM_MAP(CJMSOAddin)
        COM_INTERFACE_ENTRY2(IDispatch, IRibbonCallback)
        COM_INTERFACE_ENTRY(IJMSOAddin)
        COM_INTERFACE_ENTRY2(IDispatch, _IDTExtensibility2)
        COM_INTERFACE_ENTRY(ISupportErrorInfo)
        COM_INTERFACE_ENTRY(_IDTExtensibility2)
        COM_INTERFACE_ENTRY(Office::IRibbonExtensibility)
        COM_INTERFACE_ENTRY(IRibbonCallback)
    END_COM_MAP()

    //ATL Sink Message MAP
    BEGIN_SINK_MAP(CJMSOAddin)
        SINK_ENTRY_INFO(
            APPLICATION_EVENTS_ID,
            __uuidof(Outlook::ApplicationEvents_10),
            /*PropPageAdd*/0xf005,
            OnOptionsAddPages,
            &OnOptionsAddPagesInfo)
    END_SINK_MAP()

    // ISupportsErrorInfo
    STDMETHOD(InterfaceSupportsErrorInfo)(REFIID riid);


    DECLARE_PROTECT_FINAL_CONSTRUCT()

    HRESULT FinalConstruct()
    {
        return S_OK;
    }

    void FinalRelease()
    {
    }

//IJMSOAddin
protected:
    // Event handler functions
    void __stdcall OnItemAdd(IDispatch *pIDisp);
    void __stdcall OnOptionsAddPages(IDispatch* Ctrl);

// IJMSOAddin
public:
    // Type defination to avoid compilation 'ambigous call' errors
    typedef IDispEventSimpleImpl<APPLICATION_EVENTS_ID, CJMSOAddin, &__uuidof(Outlook::ApplicationEvents_10)> AppEvents;

    // _IDTExtensibility2 Methods
public:
    STDMETHOD(OnBeginShutdown)(SAFEARRAY * * custom);
    STDMETHOD(OnStartupComplete)(SAFEARRAY * * custom);
    STDMETHOD(OnAddInsUpdate)(SAFEARRAY * * custom);
    STDMETHOD(OnDisconnection)(ext_DisconnectMode RemoveMode, SAFEARRAY * * custom);
    STDMETHOD(OnConnection)(LPDISPATCH Application, ext_ConnectMode ConnectMode, LPDISPATCH AddInInst, SAFEARRAY * * custom);
    STDMETHOD(raw_GetCustomUI)(BSTR RibbonID, BSTR * RibbonXml);
    STDMETHOD(ButtonClicked)(IDispatch* ribbon);
    STDMETHOD(GetVisible)(IDispatch *pControl, VARIANT_BOOL *pvarReturnedVal);

public:
    static HRESULT SinkAccountInboxAddItemEvents(Outlook::_Account *lpAccount, BOOL *lpbCont, LPVOID lpvParam);

private:
    typedef struct _AccountInfo
    {
        CComPtr<Outlook::_Items> spItems;
        CComPtr<CItemsEvents> spItemsEvents;
    } AccountInfo;

private:
    // Member Variables 
    CComPtr<Outlook::_Application> m_spApp;     // Application
    CAtlList<AccountInfo *> m_AccountInfoList;

private: 
    typedef HRESULT (CJMSOAddin::*AccountEnumProc)(Outlook::_Account *lpAccount, BOOL *lpbCont, LPVOID lpvParam);

private:
    HRESULT EnumAccounts(AccountEnumProc EnumProc, LPVOID lpvParam);
    HRESULT AddAddinColumnToOutlookView(void);
};

OBJECT_ENTRY_AUTO(__uuidof(JMSOAddin), CJMSOAddin)
