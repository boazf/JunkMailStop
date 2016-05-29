

/* this ALWAYS GENERATED file contains the IIDs and CLSIDs */

/* link this file in with the server and any clients */


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


#ifdef __cplusplus
extern "C"{
#endif 


#include <rpc.h>
#include <rpcndr.h>

#ifdef _MIDL_USE_GUIDDEF_

#ifndef INITGUID
#define INITGUID
#include <guiddef.h>
#undef INITGUID
#else
#include <guiddef.h>
#endif

#define MIDL_DEFINE_GUID(type,name,l,w1,w2,b1,b2,b3,b4,b5,b6,b7,b8) \
        DEFINE_GUID(name,l,w1,w2,b1,b2,b3,b4,b5,b6,b7,b8)

#else // !_MIDL_USE_GUIDDEF_

#ifndef __IID_DEFINED__
#define __IID_DEFINED__

typedef struct _IID
{
    unsigned long x;
    unsigned short s1;
    unsigned short s2;
    unsigned char  c[8];
} IID;

#endif // __IID_DEFINED__

#ifndef CLSID_DEFINED
#define CLSID_DEFINED
typedef IID CLSID;
#endif // CLSID_DEFINED

#define MIDL_DEFINE_GUID(type,name,l,w1,w2,b1,b2,b3,b4,b5,b6,b7,b8) \
        const type name = {l,w1,w2,{b1,b2,b3,b4,b5,b6,b7,b8}}

#endif !_MIDL_USE_GUIDDEF_

MIDL_DEFINE_GUID(IID, IID_IJMSOAddin,0x535B0F64,0x0382,0x46FE,0x92,0x35,0x7E,0x5C,0xE5,0x83,0x8E,0xC5);


MIDL_DEFINE_GUID(IID, IID_IJMSOAddinPropPage,0x52F8E47F,0xCFB2,0x43E4,0x9E,0x88,0x6A,0x2A,0x33,0xE5,0x19,0xAA);


MIDL_DEFINE_GUID(IID, IID_IItemsEvents,0x62F95F25,0x2C5E,0x4CBB,0xB2,0x43,0x28,0x57,0x3F,0xAF,0xBC,0x84);


MIDL_DEFINE_GUID(IID, IID_IRibbonCallback,0xCE895442,0x9981,0x4315,0xAA,0x85,0x4B,0x9A,0x5C,0x77,0x39,0xD8);


MIDL_DEFINE_GUID(IID, LIBID_JMSOAddinLib,0x1B76F27C,0x6DDD,0x4E68,0xAA,0xB7,0x7E,0x1E,0x50,0xE1,0x7B,0x02);


MIDL_DEFINE_GUID(CLSID, CLSID_JMSOAddin,0x39C059C5,0x2959,0x4440,0xB6,0x87,0xE1,0x9D,0x12,0xEC,0x25,0x0E);


MIDL_DEFINE_GUID(CLSID, CLSID_JMSOAddinPropPage,0xE84495E9,0x5698,0x4A6F,0x94,0x74,0xE9,0x20,0x2E,0xDE,0xB6,0x8B);


MIDL_DEFINE_GUID(CLSID, CLSID_ItemsEvents,0xFAA9B6DA,0x051E,0x43AF,0x81,0x43,0xD0,0x8D,0x44,0x59,0x4F,0xFA);

#undef MIDL_DEFINE_GUID

#ifdef __cplusplus
}
#endif



