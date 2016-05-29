

/* this ALWAYS GENERATED file contains the proxy stub code */


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

#if !defined(_M_IA64) && !defined(_M_AMD64)


#pragma warning( disable: 4049 )  /* more than 64k source lines */
#if _MSC_VER >= 1200
#pragma warning(push)
#endif

#pragma warning( disable: 4211 )  /* redefine extern to static */
#pragma warning( disable: 4232 )  /* dllimport identity*/
#pragma warning( disable: 4024 )  /* array to pointer mapping*/
#pragma warning( disable: 4152 )  /* function/data pointer conversion in expression */
#pragma warning( disable: 4100 ) /* unreferenced arguments in x86 call */

#pragma optimize("", off ) 

#define USE_STUBLESS_PROXY


/* verify that the <rpcproxy.h> version is high enough to compile this file*/
#ifndef __REDQ_RPCPROXY_H_VERSION__
#define __REQUIRED_RPCPROXY_H_VERSION__ 440
#endif


#include "rpcproxy.h"
#ifndef __RPCPROXY_H_VERSION__
#error this stub requires an updated version of <rpcproxy.h>
#endif /* __RPCPROXY_H_VERSION__ */


#include "JMSOAddin.h"

#define TYPE_FORMAT_STRING_SIZE   25                                
#define PROC_FORMAT_STRING_SIZE   63                                
#define EXPR_FORMAT_STRING_SIZE   1                                 
#define TRANSMIT_AS_TABLE_SIZE    0            
#define WIRE_MARSHAL_TABLE_SIZE   0            

typedef struct _JMSOAddin_MIDL_TYPE_FORMAT_STRING
    {
    short          Pad;
    unsigned char  Format[ TYPE_FORMAT_STRING_SIZE ];
    } JMSOAddin_MIDL_TYPE_FORMAT_STRING;

typedef struct _JMSOAddin_MIDL_PROC_FORMAT_STRING
    {
    short          Pad;
    unsigned char  Format[ PROC_FORMAT_STRING_SIZE ];
    } JMSOAddin_MIDL_PROC_FORMAT_STRING;

typedef struct _JMSOAddin_MIDL_EXPR_FORMAT_STRING
    {
    long          Pad;
    unsigned char  Format[ EXPR_FORMAT_STRING_SIZE ];
    } JMSOAddin_MIDL_EXPR_FORMAT_STRING;


static const RPC_SYNTAX_IDENTIFIER  _RpcTransferSyntax = 
{{0x8A885D04,0x1CEB,0x11C9,{0x9F,0xE8,0x08,0x00,0x2B,0x10,0x48,0x60}},{2,0}};


extern const JMSOAddin_MIDL_TYPE_FORMAT_STRING JMSOAddin__MIDL_TypeFormatString;
extern const JMSOAddin_MIDL_PROC_FORMAT_STRING JMSOAddin__MIDL_ProcFormatString;
extern const JMSOAddin_MIDL_EXPR_FORMAT_STRING JMSOAddin__MIDL_ExprFormatString;


extern const MIDL_STUB_DESC Object_StubDesc;


extern const MIDL_SERVER_INFO IJMSOAddin_ServerInfo;
extern const MIDL_STUBLESS_PROXY_INFO IJMSOAddin_ProxyInfo;


extern const MIDL_STUB_DESC Object_StubDesc;


extern const MIDL_SERVER_INFO IJMSOAddinPropPage_ServerInfo;
extern const MIDL_STUBLESS_PROXY_INFO IJMSOAddinPropPage_ProxyInfo;


extern const MIDL_STUB_DESC Object_StubDesc;


extern const MIDL_SERVER_INFO IItemsEvents_ServerInfo;
extern const MIDL_STUBLESS_PROXY_INFO IItemsEvents_ProxyInfo;


extern const MIDL_STUB_DESC Object_StubDesc;


extern const MIDL_SERVER_INFO IRibbonCallback_ServerInfo;
extern const MIDL_STUBLESS_PROXY_INFO IRibbonCallback_ProxyInfo;



#if !defined(__RPC_WIN32__)
#error  Invalid build platform for this stub.
#endif

#if !(TARGET_IS_NT40_OR_LATER)
#error You need Windows NT 4.0 or later to run this stub because it uses these features:
#error   -Oif or -Oicf.
#error However, your C/C++ compilation flags indicate you intend to run this app on earlier systems.
#error This app will fail with the RPC_X_WRONG_STUB_VERSION error.
#endif


static const JMSOAddin_MIDL_PROC_FORMAT_STRING JMSOAddin__MIDL_ProcFormatString =
    {
        0,
        {

	/* Procedure ButtonClicked */

			0x33,		/* FC_AUTO_HANDLE */
			0x6c,		/* Old Flags:  object, Oi2 */
/*  2 */	NdrFcLong( 0x0 ),	/* 0 */
/*  6 */	NdrFcShort( 0x7 ),	/* 7 */
/*  8 */	NdrFcShort( 0xc ),	/* x86 Stack size/offset = 12 */
/* 10 */	NdrFcShort( 0x0 ),	/* 0 */
/* 12 */	NdrFcShort( 0x8 ),	/* 8 */
/* 14 */	0x6,		/* Oi2 Flags:  clt must size, has return, */
			0x2,		/* 2 */

	/* Parameter ribbonControl */

/* 16 */	NdrFcShort( 0xb ),	/* Flags:  must size, must free, in, */
/* 18 */	NdrFcShort( 0x4 ),	/* x86 Stack size/offset = 4 */
/* 20 */	NdrFcShort( 0x2 ),	/* Type Offset=2 */

	/* Return value */

/* 22 */	NdrFcShort( 0x70 ),	/* Flags:  out, return, base type, */
/* 24 */	NdrFcShort( 0x8 ),	/* x86 Stack size/offset = 8 */
/* 26 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

	/* Procedure GetVisible */

/* 28 */	0x33,		/* FC_AUTO_HANDLE */
			0x6c,		/* Old Flags:  object, Oi2 */
/* 30 */	NdrFcLong( 0x0 ),	/* 0 */
/* 34 */	NdrFcShort( 0x8 ),	/* 8 */
/* 36 */	NdrFcShort( 0x10 ),	/* x86 Stack size/offset = 16 */
/* 38 */	NdrFcShort( 0x0 ),	/* 0 */
/* 40 */	NdrFcShort( 0x22 ),	/* 34 */
/* 42 */	0x6,		/* Oi2 Flags:  clt must size, has return, */
			0x3,		/* 3 */

	/* Parameter pControl */

/* 44 */	NdrFcShort( 0xb ),	/* Flags:  must size, must free, in, */
/* 46 */	NdrFcShort( 0x4 ),	/* x86 Stack size/offset = 4 */
/* 48 */	NdrFcShort( 0x2 ),	/* Type Offset=2 */

	/* Parameter pvarReturnedVal */

/* 50 */	NdrFcShort( 0x2150 ),	/* Flags:  out, base type, simple ref, srv alloc size=8 */
/* 52 */	NdrFcShort( 0x8 ),	/* x86 Stack size/offset = 8 */
/* 54 */	0x6,		/* FC_SHORT */
			0x0,		/* 0 */

	/* Return value */

/* 56 */	NdrFcShort( 0x70 ),	/* Flags:  out, return, base type, */
/* 58 */	NdrFcShort( 0xc ),	/* x86 Stack size/offset = 12 */
/* 60 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

			0x0
        }
    };

static const JMSOAddin_MIDL_TYPE_FORMAT_STRING JMSOAddin__MIDL_TypeFormatString =
    {
        0,
        {
			NdrFcShort( 0x0 ),	/* 0 */
/*  2 */	
			0x2f,		/* FC_IP */
			0x5a,		/* FC_CONSTANT_IID */
/*  4 */	NdrFcLong( 0x20400 ),	/* 132096 */
/*  8 */	NdrFcShort( 0x0 ),	/* 0 */
/* 10 */	NdrFcShort( 0x0 ),	/* 0 */
/* 12 */	0xc0,		/* 192 */
			0x0,		/* 0 */
/* 14 */	0x0,		/* 0 */
			0x0,		/* 0 */
/* 16 */	0x0,		/* 0 */
			0x0,		/* 0 */
/* 18 */	0x0,		/* 0 */
			0x46,		/* 70 */
/* 20 */	
			0x11, 0xc,	/* FC_RP [alloced_on_stack] [simple_pointer] */
/* 22 */	0x6,		/* FC_SHORT */
			0x5c,		/* FC_PAD */

			0x0
        }
    };


/* Object interface: IUnknown, ver. 0.0,
   GUID={0x00000000,0x0000,0x0000,{0xC0,0x00,0x00,0x00,0x00,0x00,0x00,0x46}} */


/* Object interface: IDispatch, ver. 0.0,
   GUID={0x00020400,0x0000,0x0000,{0xC0,0x00,0x00,0x00,0x00,0x00,0x00,0x46}} */


/* Object interface: IJMSOAddin, ver. 0.0,
   GUID={0x535B0F64,0x0382,0x46FE,{0x92,0x35,0x7E,0x5C,0xE5,0x83,0x8E,0xC5}} */

#pragma code_seg(".orpc")
static const unsigned short IJMSOAddin_FormatStringOffsetTable[] =
    {
    (unsigned short) -1,
    (unsigned short) -1,
    (unsigned short) -1,
    (unsigned short) -1,
    0
    };

static const MIDL_STUBLESS_PROXY_INFO IJMSOAddin_ProxyInfo =
    {
    &Object_StubDesc,
    JMSOAddin__MIDL_ProcFormatString.Format,
    &IJMSOAddin_FormatStringOffsetTable[-3],
    0,
    0,
    0
    };


static const MIDL_SERVER_INFO IJMSOAddin_ServerInfo = 
    {
    &Object_StubDesc,
    0,
    JMSOAddin__MIDL_ProcFormatString.Format,
    &IJMSOAddin_FormatStringOffsetTable[-3],
    0,
    0,
    0,
    0};
CINTERFACE_PROXY_VTABLE(7) _IJMSOAddinProxyVtbl = 
{
    0,
    &IID_IJMSOAddin,
    IUnknown_QueryInterface_Proxy,
    IUnknown_AddRef_Proxy,
    IUnknown_Release_Proxy ,
    0 /* IDispatch::GetTypeInfoCount */ ,
    0 /* IDispatch::GetTypeInfo */ ,
    0 /* IDispatch::GetIDsOfNames */ ,
    0 /* IDispatch_Invoke_Proxy */
};


static const PRPC_STUB_FUNCTION IJMSOAddin_table[] =
{
    STUB_FORWARDING_FUNCTION,
    STUB_FORWARDING_FUNCTION,
    STUB_FORWARDING_FUNCTION,
    STUB_FORWARDING_FUNCTION
};

CInterfaceStubVtbl _IJMSOAddinStubVtbl =
{
    &IID_IJMSOAddin,
    &IJMSOAddin_ServerInfo,
    7,
    &IJMSOAddin_table[-3],
    CStdStubBuffer_DELEGATING_METHODS
};


/* Object interface: IJMSOAddinPropPage, ver. 0.0,
   GUID={0x52F8E47F,0xCFB2,0x43E4,{0x9E,0x88,0x6A,0x2A,0x33,0xE5,0x19,0xAA}} */

#pragma code_seg(".orpc")
static const unsigned short IJMSOAddinPropPage_FormatStringOffsetTable[] =
    {
    (unsigned short) -1,
    (unsigned short) -1,
    (unsigned short) -1,
    (unsigned short) -1,
    0
    };

static const MIDL_STUBLESS_PROXY_INFO IJMSOAddinPropPage_ProxyInfo =
    {
    &Object_StubDesc,
    JMSOAddin__MIDL_ProcFormatString.Format,
    &IJMSOAddinPropPage_FormatStringOffsetTable[-3],
    0,
    0,
    0
    };


static const MIDL_SERVER_INFO IJMSOAddinPropPage_ServerInfo = 
    {
    &Object_StubDesc,
    0,
    JMSOAddin__MIDL_ProcFormatString.Format,
    &IJMSOAddinPropPage_FormatStringOffsetTable[-3],
    0,
    0,
    0,
    0};
CINTERFACE_PROXY_VTABLE(7) _IJMSOAddinPropPageProxyVtbl = 
{
    0,
    &IID_IJMSOAddinPropPage,
    IUnknown_QueryInterface_Proxy,
    IUnknown_AddRef_Proxy,
    IUnknown_Release_Proxy ,
    0 /* IDispatch::GetTypeInfoCount */ ,
    0 /* IDispatch::GetTypeInfo */ ,
    0 /* IDispatch::GetIDsOfNames */ ,
    0 /* IDispatch_Invoke_Proxy */
};


static const PRPC_STUB_FUNCTION IJMSOAddinPropPage_table[] =
{
    STUB_FORWARDING_FUNCTION,
    STUB_FORWARDING_FUNCTION,
    STUB_FORWARDING_FUNCTION,
    STUB_FORWARDING_FUNCTION
};

CInterfaceStubVtbl _IJMSOAddinPropPageStubVtbl =
{
    &IID_IJMSOAddinPropPage,
    &IJMSOAddinPropPage_ServerInfo,
    7,
    &IJMSOAddinPropPage_table[-3],
    CStdStubBuffer_DELEGATING_METHODS
};


/* Object interface: IItemsEvents, ver. 0.0,
   GUID={0x62F95F25,0x2C5E,0x4CBB,{0xB2,0x43,0x28,0x57,0x3F,0xAF,0xBC,0x84}} */

#pragma code_seg(".orpc")
static const unsigned short IItemsEvents_FormatStringOffsetTable[] =
    {
    (unsigned short) -1,
    (unsigned short) -1,
    (unsigned short) -1,
    (unsigned short) -1,
    0
    };

static const MIDL_STUBLESS_PROXY_INFO IItemsEvents_ProxyInfo =
    {
    &Object_StubDesc,
    JMSOAddin__MIDL_ProcFormatString.Format,
    &IItemsEvents_FormatStringOffsetTable[-3],
    0,
    0,
    0
    };


static const MIDL_SERVER_INFO IItemsEvents_ServerInfo = 
    {
    &Object_StubDesc,
    0,
    JMSOAddin__MIDL_ProcFormatString.Format,
    &IItemsEvents_FormatStringOffsetTable[-3],
    0,
    0,
    0,
    0};
CINTERFACE_PROXY_VTABLE(7) _IItemsEventsProxyVtbl = 
{
    0,
    &IID_IItemsEvents,
    IUnknown_QueryInterface_Proxy,
    IUnknown_AddRef_Proxy,
    IUnknown_Release_Proxy ,
    0 /* IDispatch::GetTypeInfoCount */ ,
    0 /* IDispatch::GetTypeInfo */ ,
    0 /* IDispatch::GetIDsOfNames */ ,
    0 /* IDispatch_Invoke_Proxy */
};


static const PRPC_STUB_FUNCTION IItemsEvents_table[] =
{
    STUB_FORWARDING_FUNCTION,
    STUB_FORWARDING_FUNCTION,
    STUB_FORWARDING_FUNCTION,
    STUB_FORWARDING_FUNCTION
};

CInterfaceStubVtbl _IItemsEventsStubVtbl =
{
    &IID_IItemsEvents,
    &IItemsEvents_ServerInfo,
    7,
    &IItemsEvents_table[-3],
    CStdStubBuffer_DELEGATING_METHODS
};


/* Object interface: IRibbonCallback, ver. 0.0,
   GUID={0xCE895442,0x9981,0x4315,{0xAA,0x85,0x4B,0x9A,0x5C,0x77,0x39,0xD8}} */

#pragma code_seg(".orpc")
static const unsigned short IRibbonCallback_FormatStringOffsetTable[] =
    {
    (unsigned short) -1,
    (unsigned short) -1,
    (unsigned short) -1,
    (unsigned short) -1,
    0,
    28
    };

static const MIDL_STUBLESS_PROXY_INFO IRibbonCallback_ProxyInfo =
    {
    &Object_StubDesc,
    JMSOAddin__MIDL_ProcFormatString.Format,
    &IRibbonCallback_FormatStringOffsetTable[-3],
    0,
    0,
    0
    };


static const MIDL_SERVER_INFO IRibbonCallback_ServerInfo = 
    {
    &Object_StubDesc,
    0,
    JMSOAddin__MIDL_ProcFormatString.Format,
    &IRibbonCallback_FormatStringOffsetTable[-3],
    0,
    0,
    0,
    0};
CINTERFACE_PROXY_VTABLE(9) _IRibbonCallbackProxyVtbl = 
{
    &IRibbonCallback_ProxyInfo,
    &IID_IRibbonCallback,
    IUnknown_QueryInterface_Proxy,
    IUnknown_AddRef_Proxy,
    IUnknown_Release_Proxy ,
    0 /* IDispatch::GetTypeInfoCount */ ,
    0 /* IDispatch::GetTypeInfo */ ,
    0 /* IDispatch::GetIDsOfNames */ ,
    0 /* IDispatch_Invoke_Proxy */ ,
    (void *) (INT_PTR) -1 /* IRibbonCallback::ButtonClicked */ ,
    (void *) (INT_PTR) -1 /* IRibbonCallback::GetVisible */
};


static const PRPC_STUB_FUNCTION IRibbonCallback_table[] =
{
    STUB_FORWARDING_FUNCTION,
    STUB_FORWARDING_FUNCTION,
    STUB_FORWARDING_FUNCTION,
    STUB_FORWARDING_FUNCTION,
    NdrStubCall2,
    NdrStubCall2
};

CInterfaceStubVtbl _IRibbonCallbackStubVtbl =
{
    &IID_IRibbonCallback,
    &IRibbonCallback_ServerInfo,
    9,
    &IRibbonCallback_table[-3],
    CStdStubBuffer_DELEGATING_METHODS
};

static const MIDL_STUB_DESC Object_StubDesc = 
    {
    0,
    NdrOleAllocate,
    NdrOleFree,
    0,
    0,
    0,
    0,
    0,
    JMSOAddin__MIDL_TypeFormatString.Format,
    1, /* -error bounds_check flag */
    0x20000, /* Ndr library version */
    0,
    0x700022b, /* MIDL Version 7.0.555 */
    0,
    0,
    0,  /* notify & notify_flag routine table */
    0x1, /* MIDL flag */
    0, /* cs routines */
    0,   /* proxy/server info */
    0
    };

const CInterfaceProxyVtbl * const _JMSOAddin_ProxyVtblList[] = 
{
    ( CInterfaceProxyVtbl *) &_IItemsEventsProxyVtbl,
    ( CInterfaceProxyVtbl *) &_IRibbonCallbackProxyVtbl,
    ( CInterfaceProxyVtbl *) &_IJMSOAddinProxyVtbl,
    ( CInterfaceProxyVtbl *) &_IJMSOAddinPropPageProxyVtbl,
    0
};

const CInterfaceStubVtbl * const _JMSOAddin_StubVtblList[] = 
{
    ( CInterfaceStubVtbl *) &_IItemsEventsStubVtbl,
    ( CInterfaceStubVtbl *) &_IRibbonCallbackStubVtbl,
    ( CInterfaceStubVtbl *) &_IJMSOAddinStubVtbl,
    ( CInterfaceStubVtbl *) &_IJMSOAddinPropPageStubVtbl,
    0
};

PCInterfaceName const _JMSOAddin_InterfaceNamesList[] = 
{
    "IItemsEvents",
    "IRibbonCallback",
    "IJMSOAddin",
    "IJMSOAddinPropPage",
    0
};

const IID *  const _JMSOAddin_BaseIIDList[] = 
{
    &IID_IDispatch,
    &IID_IDispatch,
    &IID_IDispatch,
    &IID_IDispatch,
    0
};


#define _JMSOAddin_CHECK_IID(n)	IID_GENERIC_CHECK_IID( _JMSOAddin, pIID, n)

int __stdcall _JMSOAddin_IID_Lookup( const IID * pIID, int * pIndex )
{
    IID_BS_LOOKUP_SETUP

    IID_BS_LOOKUP_INITIAL_TEST( _JMSOAddin, 4, 2 )
    IID_BS_LOOKUP_NEXT_TEST( _JMSOAddin, 1 )
    IID_BS_LOOKUP_RETURN_RESULT( _JMSOAddin, 4, *pIndex )
    
}

const ExtendedProxyFileInfo JMSOAddin_ProxyFileInfo = 
{
    (PCInterfaceProxyVtblList *) & _JMSOAddin_ProxyVtblList,
    (PCInterfaceStubVtblList *) & _JMSOAddin_StubVtblList,
    (const PCInterfaceName * ) & _JMSOAddin_InterfaceNamesList,
    (const IID ** ) & _JMSOAddin_BaseIIDList,
    & _JMSOAddin_IID_Lookup, 
    4,
    2,
    0, /* table of [async_uuid] interfaces */
    0, /* Filler1 */
    0, /* Filler2 */
    0  /* Filler3 */
};
#pragma optimize("", on )
#if _MSC_VER >= 1200
#pragma warning(pop)
#endif


#endif /* !defined(_M_IA64) && !defined(_M_AMD64)*/

