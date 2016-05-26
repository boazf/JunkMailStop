#include "revision.h"

#define STRINGIZE2(s) #s
#define STRINGIZE(s) STRINGIZE2(s)
 
#define VER_MAJOR               1
#define VER_MINOR               0
#define VER_BUILD               0
 
#define VER_FILE_VERSION            VER_MAJOR, VER_MINOR, VER_REVISION, VER_BUILD
#define VER_FILE_VERSION_STR        STRINGIZE(VER_MAJOR)        \
                                    "." STRINGIZE(VER_MINOR)    \
                                    "." STRINGIZE(VER_REVISION) \
                                    "." STRINGIZE(VER_BUILD)    \
 
#define VER_COMPANY_NAME_STR        "Junk Stop"
#define VER_PRODUCTNAME_STR         "Junk Mail Stop Outtlook Addin"
#define VER_PRODUCT_VERSION         VER_FILE_VERSION
#define VER_PRODUCT_VERSION_STR     VER_FILE_VERSION_STR
#define VER_COPYRIGHT_STR           "Copyright (C) 2016"
 
#ifdef _DEBUG
  #define VER_VER_DEBUG             VS_FF_DEBUG
#else
  #define VER_VER_DEBUG             0
#endif

#define VER_FILEOS                  VOS_NT_WINDOWS32
#define VER_FILEFLAGS               VER_VER_DEBUG
#define VER_FILEFLAGSMASK           0x3FL
