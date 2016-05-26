
#pragma once

#define _ATL_APARTMENT_THREADED
#define _ATL_NO_AUTOMATIC_NAMESPACE

#define _ATL_CSTRING_EXPLICIT_CONSTRUCTORS  // some CString constructors will be explicit


#include <atlbase.h>
#include <atlcom.h>
#include <atlstr.h>
#include <atlcoll.h>
#include <atltime.h>
using namespace ATL;
#include "resource.h"

#include <new>
using namespace std;

#import "C:\Program Files (x86)\Common Files\microsoft shared\OFFICE14\mso.dll" rename_namespace("Office"), named_guids, rename("RGB", "Office_RGB"), rename("DocumentProperties", "Office_DocumentProperties"), rename("SearchPath", "Office_SerarchPath")
#import "C:\Program Files (x86)\Microsoft Office\Office14\MSOUTL.olb" rename_namespace("Outlook"), raw_interfaces_only, named_guids, rename("CopyFile", "Outlook_CopyFile"), rename("PlaySound", "Outlook_PlaySound")
#import "C:\Program Files (x86)\Common Files\DESIGNER\MSADDNDR.dll" raw_interfaces_only, raw_native_types, no_namespace, named_guids, auto_search

#define NELEMS(a) (sizeof(a) / sizeof(*a))

extern HINSTANCE GetApplicationHInstance();

