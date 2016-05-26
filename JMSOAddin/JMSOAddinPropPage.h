// JMSOAddinPropPage.h : Declaration of the CJMSOAddinPropPage
#pragma once
#include "resource.h"       // main symbols
#include <atlctl.h>
#include "JMSOAddin.h"
#include "atlcontrols.h"
using namespace ATLControls;

#if defined(_WIN32_WCE) && !defined(_CE_DCOM) && !defined(_CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA)
#error "Single-threaded COM objects are not properly supported on Windows CE platform, such as the Windows Mobile platforms that do not include full DCOM support. Define _CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA to force ATL to support creating single-thread COM object's and allow use of it's single-threaded COM object implementations. The threading model in your rgs file was set to 'Free' as that is the only threading model supported in non DCOM Windows CE platforms."
#endif


// CJMSOAddinPropPage
class ATL_NO_VTABLE CJMSOAddinPropPage :
    public CComObjectRootEx<CComSingleThreadModel>,
    public IDispatchImpl<IJMSOAddinPropPage, &IID_IJMSOAddinPropPage, &LIBID_JMSOAddinLib, /*wMajor =*/ 1, /*wMinor =*/ 0>,
    public IPersistStreamInitImpl<CJMSOAddinPropPage>,
    public IOleControlImpl<CJMSOAddinPropPage>,
    public IOleObjectImpl<CJMSOAddinPropPage>,
    public IOleInPlaceActiveObjectImpl<CJMSOAddinPropPage>,
    public IViewObjectExImpl<CJMSOAddinPropPage>,
    public IOleInPlaceObjectWindowlessImpl<CJMSOAddinPropPage>,
    public ISupportErrorInfo,
    public IPersistStorageImpl<CJMSOAddinPropPage>,
    public ISpecifyPropertyPagesImpl<CJMSOAddinPropPage>,
    public IQuickActivateImpl<CJMSOAddinPropPage>,
#ifndef _WIN32_WCE
    public IDataObjectImpl<CJMSOAddinPropPage>,
#endif
    public IProvideClassInfo2Impl<&CLSID_JMSOAddinPropPage, NULL, &LIBID_JMSOAddinLib>,
#ifdef _WIN32_WCE // IObjectSafety is required on Windows CE for the control to be loaded correctly
    public IObjectSafetyImpl<CJMSOAddinPropPage, INTERFACESAFE_FOR_UNTRUSTED_CALLER>,
#endif
    public CComCoClass<CJMSOAddinPropPage, &CLSID_JMSOAddinPropPage>,
    public CComCompositeControl<CJMSOAddinPropPage>,
    public IDispatchImpl<Outlook::PropertyPage,&__uuidof(Outlook::PropertyPage),&LIBID_JMSOAddinLib>
{
public:


    CJMSOAddinPropPage()
    {
        m_bWindowOnly = TRUE;
        CalcExtent(m_sizeExtent);
    }

DECLARE_OLEMISC_STATUS(OLEMISC_RECOMPOSEONRESIZE |
    OLEMISC_CANTLINKINSIDE |
    OLEMISC_INSIDEOUT |
    OLEMISC_ACTIVATEWHENVISIBLE |
    OLEMISC_SETCLIENTSITEFIRST
)

DECLARE_REGISTRY_RESOURCEID(IDR_JMSOADDINPROPPAGE)


BEGIN_COM_MAP(CJMSOAddinPropPage)
    COM_INTERFACE_ENTRY(IJMSOAddinPropPage)
    COM_INTERFACE_ENTRY2(IDispatch,IJMSOAddinPropPage)
    COM_INTERFACE_ENTRY(IViewObjectEx)
    COM_INTERFACE_ENTRY(IViewObject2)
    COM_INTERFACE_ENTRY(IViewObject)
    COM_INTERFACE_ENTRY(IOleInPlaceObjectWindowless)
    COM_INTERFACE_ENTRY(IOleInPlaceObject)
    COM_INTERFACE_ENTRY2(IOleWindow, IOleInPlaceObjectWindowless)
    COM_INTERFACE_ENTRY(IOleInPlaceActiveObject)
    COM_INTERFACE_ENTRY(IOleControl)
    COM_INTERFACE_ENTRY(IOleObject)
    COM_INTERFACE_ENTRY(IPersistStreamInit)
    COM_INTERFACE_ENTRY2(IPersist, IPersistStreamInit)
    COM_INTERFACE_ENTRY(ISupportErrorInfo)
    COM_INTERFACE_ENTRY(ISpecifyPropertyPages)
    COM_INTERFACE_ENTRY(IQuickActivate)
    COM_INTERFACE_ENTRY(IPersistStorage)
#ifndef _WIN32_WCE
    COM_INTERFACE_ENTRY(IDataObject)
#endif
    COM_INTERFACE_ENTRY(IProvideClassInfo)
    COM_INTERFACE_ENTRY(IProvideClassInfo2)
#ifdef _WIN32_WCE // IObjectSafety is required on Windows CE for the control to be loaded correctly
    COM_INTERFACE_ENTRY_IID(IID_IObjectSafety, IObjectSafety)
#endif
    COM_INTERFACE_ENTRY(Outlook::PropertyPage) 
END_COM_MAP()

BEGIN_PROP_MAP(CJMSOAddinPropPage)
    PROP_DATA_ENTRY("_cx", m_sizeExtent.cx, VT_UI4)
    PROP_DATA_ENTRY("_cy", m_sizeExtent.cy, VT_UI4)
    // Example entries
    // PROP_ENTRY("Property Description", dispid, clsid)
    // PROP_PAGE(CLSID_StockColorPage)
END_PROP_MAP()


BEGIN_MSG_MAP(CJMSOAddinPropPage)
    MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
    CHAIN_MSG_MAP(CComCompositeControl<CJMSOAddinPropPage>)
    COMMAND_HANDLER(IDC_SERVERS_BLACK_LIST, LBN_SELCHANGE, OnLbnSelchangeServersBlackList)
    COMMAND_HANDLER(IDC_REMOVE_BUTTON, BN_CLICKED, OnBnClickedRemoveButton)
    COMMAND_HANDLER(IDC_ADD_BUTTON, BN_CLICKED, OnBnClickedAddButton)
END_MSG_MAP()
// Handler prototypes:
//  LRESULT MessageHandler(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
//  LRESULT CommandHandler(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled);
//  LRESULT NotifyHandler(int idCtrl, LPNMHDR pnmh, BOOL& bHandled);

BEGIN_SINK_MAP(CJMSOAddinPropPage)
    //Make sure the Event Handlers have __stdcall calling convention
END_SINK_MAP()

    STDMETHOD(OnAmbientPropertyChange)(DISPID dispid)
    {
        if (dispid == DISPID_AMBIENT_BACKCOLOR)
        {
            SetBackgroundColorFromAmbient();
            FireViewChange();
        }
        return IOleControlImpl<CJMSOAddinPropPage>::OnAmbientPropertyChange(dispid);
    }

    // ISupportsErrorInfo
    STDMETHOD(InterfaceSupportsErrorInfo)(REFIID riid)
    {
        static const IID* arr[] =
        {
            &IID_IJMSOAddinPropPage,
        };

        for (int i=0; i<sizeof(arr)/sizeof(arr[0]); i++)
        {
            if (InlineIsEqualGUID(*arr[i], riid))
                return S_OK;
        }
        return S_FALSE;
    }

// IViewObjectEx
    DECLARE_VIEW_STATUS(VIEWSTATUS_SOLIDBKGND | VIEWSTATUS_OPAQUE)

// IJMSOAddinPropPage

    enum { IDD = IDD_JMSOADDINPROPPAGE };

    DECLARE_PROTECT_FINAL_CONSTRUCT()

    HRESULT FinalConstruct()
    {
        return S_OK;
    }

    void FinalRelease()
    {
    }

private:
    CRegKey m_RegKey;
    CComVariant m_fDirty;
    BOOL m_bInitializing;
    CComPtr<Outlook::PropertyPageSite> m_pPropPageSite; 

private:
    ATLControls::CListBox m_serversBlackList;
    ATLControls::CButton m_removeBtn;
    ATLControls::CButton m_addBtn;

private:
    HRESULT MarkAsDirty(void);

public:
    LRESULT OnInitDialog(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
    STDMETHOD(GetControlInfo)(LPCONTROLINFO lpCI);
    STDMETHOD(SetClientSite)(IOleClientSite *pSite);
    STDMETHOD(get_Dirty)(VARIANT_BOOL * Dirty);
    STDMETHOD(GetPageInfo)(BSTR * HelpFile, long * HelpContext);
    STDMETHOD(Apply)();

public:
    LRESULT OnLbnSelchangeServersBlackList(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
    LRESULT OnBnClickedRemoveButton(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
    LRESULT OnBnClickedAddButton(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
};

OBJECT_ENTRY_AUTO(__uuidof(JMSOAddinPropPage), CJMSOAddinPropPage)
