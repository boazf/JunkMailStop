

/* this ALWAYS GENERATED file contains the definitions for the interfaces */


 /* File created by MIDL compiler version 7.00.0555 */
/* at Sun May 29 17:58:12 2016
 */
/* Compiler settings for JMSOAddin.idl:
    Oicf, W1, Zp8, env=Win32 (32b run), target_arch=X86 7.00.0555 
    protocol : dce , ms_ext, c_ext
    error checks: allocation ref bounds_check enum stub_data 
    VC __declspec() decoration level: 
         __declspec(uuid()), __declspec(selectany), __declspec(novtable)
         DECLSPEC_UUID(), MIDL_INTERFACE()
*/
/* @@MIDL_FILE_HEADING(  ) */

#pragma warning( disable: 4049 )  /* more than 64k source lines */


/* verify that the <rpcndr.h> version is high enough to compile this file*/
#ifndef __REQUIRED_RPCNDR_H_VERSION__
#define __REQUIRED_RPCNDR_H_VERSION__ 440
#endif

#include "rpc.h"
#include "rpcndr.h"

#ifndef __RPCNDR_H_VERSION__
#error this stub requires an updated version of <rpcndr.h>
#endif // __RPCNDR_H_VERSION__

#ifndef COM_NO_WINDOWS_H
#include "windows.h"
#include "ole2.h"
#endif /*COM_NO_WINDOWS_H*/

#ifndef __JMSOAddin_h__
#define __JMSOAddin_h__

#if defined(_MSC_VER) && (_MSC_VER >= 1020)
#pragma once
#endif

/* Forward Declarations */ 

#ifndef __IJMSOAddin_FWD_DEFINED__
#define __IJMSOAddin_FWD_DEFINED__
typedef interface IJMSOAddin IJMSOAddin;
#endif 	/* __IJMSOAddin_FWD_DEFINED__ */


#ifndef __IJMSOAddinPropPage_FWD_DEFINED__
#define __IJMSOAddinPropPage_FWD_DEFINED__
typedef interface IJMSOAddinPropPage IJMSOAddinPropPage;
#endif 	/* __IJMSOAddinPropPage_FWD_DEFINED__ */


#ifndef __IItemsEvents_FWD_DEFINED__
#define __IItemsEvents_FWD_DEFINED__
typedef interface IItemsEvents IItemsEvents;
#endif 	/* __IItemsEvents_FWD_DEFINED__ */


#ifndef __IRibbonCallback_FWD_DEFINED__
#define __IRibbonCallback_FWD_DEFINED__
typedef interface IRibbonCallback IRibbonCallback;
#endif 	/* __IRibbonCallback_FWD_DEFINED__ */


#ifndef __JMSOAddin_FWD_DEFINED__
#define __JMSOAddin_FWD_DEFINED__

#ifdef __cplusplus
typedef class JMSOAddin JMSOAddin;
#else
typedef struct JMSOAddin JMSOAddin;
#endif /* __cplusplus */

#endif 	/* __JMSOAddin_FWD_DEFINED__ */


#ifndef __JMSOAddinPropPage_FWD_DEFINED__
#define __JMSOAddinPropPage_FWD_DEFINED__

#ifdef __cplusplus
typedef class JMSOAddinPropPage JMSOAddinPropPage;
#else
typedef struct JMSOAddinPropPage JMSOAddinPropPage;
#endif /* __cplusplus */

#endif 	/* __JMSOAddinPropPage_FWD_DEFINED__ */


#ifndef __ItemsEvents_FWD_DEFINED__
#define __ItemsEvents_FWD_DEFINED__

#ifdef __cplusplus
typedef class ItemsEvents ItemsEvents;
#else
typedef struct ItemsEvents ItemsEvents;
#endif /* __cplusplus */

#endif 	/* __ItemsEvents_FWD_DEFINED__ */


/* header files for imported files */
#include "oaidl.h"
#include "ocidl.h"

#ifdef __cplusplus
extern "C"{
#endif 


#ifndef __IJMSOAddin_INTERFACE_DEFINED__
#define __IJMSOAddin_INTERFACE_DEFINED__

/* interface IJMSOAddin */
/* [unique][helpstring][nonextensible][dual][uuid][object] */ 


EXTERN_C const IID IID_IJMSOAddin;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("535B0F64-0382-46FE-9235-7E5CE5838EC5")
    IJMSOAddin : public IDispatch
    {
    public:
    };
    
#else 	/* C style interface */

    typedef struct IJMSOAddinVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IJMSOAddin * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            __RPC__deref_out  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IJMSOAddin * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IJMSOAddin * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            IJMSOAddin * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            IJMSOAddin * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            IJMSOAddin * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            IJMSOAddin * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS *pDispParams,
            /* [out] */ VARIANT *pVarResult,
            /* [out] */ EXCEPINFO *pExcepInfo,
            /* [out] */ UINT *puArgErr);
        
        END_INTERFACE
    } IJMSOAddinVtbl;

    interface IJMSOAddin
    {
        CONST_VTBL struct IJMSOAddinVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IJMSOAddin_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define IJMSOAddin_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define IJMSOAddin_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define IJMSOAddin_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define IJMSOAddin_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define IJMSOAddin_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define IJMSOAddin_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __IJMSOAddin_INTERFACE_DEFINED__ */


#ifndef __IJMSOAddinPropPage_INTERFACE_DEFINED__
#define __IJMSOAddinPropPage_INTERFACE_DEFINED__

/* interface IJMSOAddinPropPage */
/* [unique][helpstring][nonextensible][dual][uuid][object] */ 


EXTERN_C const IID IID_IJMSOAddinPropPage;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("52F8E47F-CFB2-43E4-9E88-6A2A33E519AA")
    IJMSOAddinPropPage : public IDispatch
    {
    public:
    };
    
#else 	/* C style interface */

    typedef struct IJMSOAddinPropPageVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IJMSOAddinPropPage * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            __RPC__deref_out  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IJMSOAddinPropPage * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IJMSOAddinPropPage * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            IJMSOAddinPropPage * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            IJMSOAddinPropPage * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            IJMSOAddinPropPage * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            IJMSOAddinPropPage * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS *pDispParams,
            /* [out] */ VARIANT *pVarResult,
            /* [out] */ EXCEPINFO *pExcepInfo,
            /* [out] */ UINT *puArgErr);
        
        END_INTERFACE
    } IJMSOAddinPropPageVtbl;

    interface IJMSOAddinPropPage
    {
        CONST_VTBL struct IJMSOAddinPropPageVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IJMSOAddinPropPage_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define IJMSOAddinPropPage_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define IJMSOAddinPropPage_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define IJMSOAddinPropPage_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define IJMSOAddinPropPage_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define IJMSOAddinPropPage_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define IJMSOAddinPropPage_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __IJMSOAddinPropPage_INTERFACE_DEFINED__ */


#ifndef __IItemsEvents_INTERFACE_DEFINED__
#define __IItemsEvents_INTERFACE_DEFINED__

/* interface IItemsEvents */
/* [unique][nonextensible][dual][uuid][object] */ 


EXTERN_C const IID IID_IItemsEvents;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("62F95F25-2C5E-4CBB-B243-28573FAFBC84")
    IItemsEvents : public IDispatch
    {
    public:
    };
    
#else 	/* C style interface */

    typedef struct IItemsEventsVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IItemsEvents * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            __RPC__deref_out  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IItemsEvents * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IItemsEvents * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            IItemsEvents * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            IItemsEvents * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            IItemsEvents * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            IItemsEvents * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS *pDispParams,
            /* [out] */ VARIANT *pVarResult,
            /* [out] */ EXCEPINFO *pExcepInfo,
            /* [out] */ UINT *puArgErr);
        
        END_INTERFACE
    } IItemsEventsVtbl;

    interface IItemsEvents
    {
        CONST_VTBL struct IItemsEventsVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IItemsEvents_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define IItemsEvents_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define IItemsEvents_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define IItemsEvents_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define IItemsEvents_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define IItemsEvents_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define IItemsEvents_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __IItemsEvents_INTERFACE_DEFINED__ */


#ifndef __IRibbonCallback_INTERFACE_DEFINED__
#define __IRibbonCallback_INTERFACE_DEFINED__

/* interface IRibbonCallback */
/* [unique][helpstring][nonextensible][dual][uuid][object] */ 


EXTERN_C const IID IID_IRibbonCallback;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("CE895442-9981-4315-AA85-4B9A5C7739D8")
    IRibbonCallback : public IDispatch
    {
    public:
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ButtonClicked( 
            /* [in] */ IDispatch *ribbonControl) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE GetVisible( 
            /* [in] */ IDispatch *pControl,
            /* [retval][out] */ VARIANT_BOOL *pvarReturnedVal) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct IRibbonCallbackVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IRibbonCallback * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            __RPC__deref_out  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IRibbonCallback * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IRibbonCallback * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            IRibbonCallback * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            IRibbonCallback * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            IRibbonCallback * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            IRibbonCallback * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS *pDispParams,
            /* [out] */ VARIANT *pVarResult,
            /* [out] */ EXCEPINFO *pExcepInfo,
            /* [out] */ UINT *puArgErr);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *ButtonClicked )( 
            IRibbonCallback * This,
            /* [in] */ IDispatch *ribbonControl);
        
        HRESULT ( STDMETHODCALLTYPE *GetVisible )( 
            IRibbonCallback * This,
            /* [in] */ IDispatch *pControl,
            /* [retval][out] */ VARIANT_BOOL *pvarReturnedVal);
        
        END_INTERFACE
    } IRibbonCallbackVtbl;

    interface IRibbonCallback
    {
        CONST_VTBL struct IRibbonCallbackVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IRibbonCallback_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define IRibbonCallback_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define IRibbonCallback_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define IRibbonCallback_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define IRibbonCallback_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define IRibbonCallback_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define IRibbonCallback_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#define IRibbonCallback_ButtonClicked(This,ribbonControl)	\
    ( (This)->lpVtbl -> ButtonClicked(This,ribbonControl) ) 

#define IRibbonCallback_GetVisible(This,pControl,pvarReturnedVal)	\
    ( (This)->lpVtbl -> GetVisible(This,pControl,pvarReturnedVal) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __IRibbonCallback_INTERFACE_DEFINED__ */



#ifndef __JMSOAddinLib_LIBRARY_DEFINED__
#define __JMSOAddinLib_LIBRARY_DEFINED__

/* library JMSOAddinLib */
/* [helpstring][version][uuid] */ 


EXTERN_C const IID LIBID_JMSOAddinLib;

EXTERN_C const CLSID CLSID_JMSOAddin;

#ifdef __cplusplus

class DECLSPEC_UUID("39C059C5-2959-4440-B687-E19D12EC250E")
JMSOAddin;
#endif

EXTERN_C const CLSID CLSID_JMSOAddinPropPage;

#ifdef __cplusplus

class DECLSPEC_UUID("E84495E9-5698-4A6F-9474-E9202EDEB68B")
JMSOAddinPropPage;
#endif

EXTERN_C const CLSID CLSID_ItemsEvents;

#ifdef __cplusplus

class DECLSPEC_UUID("FAA9B6DA-051E-43AF-8143-D08D44594FFA")
ItemsEvents;
#endif
#endif /* __JMSOAddinLib_LIBRARY_DEFINED__ */

/* Additional Prototypes for ALL interfaces */

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif


