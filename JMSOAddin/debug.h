// Debug trace macros and functions.
// TODO: patch for the global trace
#undef TRACE_INFO

extern DWORD g_dwDebugFlags;    // Global flags to switch debug trace categories.

#define TRACE_ERRORS    1           // Enable tracing of errors.
#define TRACE_INFO      2           // Enable tracing of information.
#define TRACE_VERBOSE_INFO  4       // Enable tracing of verbose information.
#define LOG_TO_DEBUGGER 0x10000     // Enable log to debugger output window.
#define LOG_TO_FILE     0x20000     // Enable log to file.
#define APPEND_LOG_FILE 0x40000     // Append current log to pervious log from previous run.

// Default debug flags.
#ifdef _DEBUG
#define DEFAULT_DEBUG_FLAGS (TRACE_ERRORS | TRACE_INFO | TRACE_VERBOSE_INFO | LOG_TO_DEBUGGER | LOG_TO_FILE | APPEND_LOG_FILE)
#else
#define DEFAULT_DEBUG_FLAGS APPEND_LOG_FILE
#endif
// Default log file name.
#define DEFAULT_LOG_FILE_NAME _T("JMSOAddin.log")

#define DEBUG_FLAGS_REGNAME _T("DebugFlags")            // Name of registry value to store the debug flags.
#define DEBUG_LOG_FILE_REGNAME _T("DebugLogFileName")   // Name of registry value to store the log file name.

#define ADDIN_REG_KEY_PATH _T("Software\\JunkStop\\Addin") // Registry path to store debug values.

// Macro to generate trace code.
#define DEBUG_TRACE(fl, x) \
    if (fl & g_dwDebugFlags) \
    {\
        if (g_dwDebugFlags & LOG_TO_DEBUGGER) OutputDebugStringW(L"VCOA: "); \
        OutputDebugPrintf(L"%hs: ", __FUNCTION__); \
        OutputDebugPrintf ##x ; \
    }

// Function that generates the trace output.
void __cdecl OutputDebugPrintf(LPCWSTR f, ...);
