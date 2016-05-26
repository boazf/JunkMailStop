#include "stdafx.h"
#include "debug.h"

#pragma warning (disable:4996)  // Don't worry it is only in debug builds.

DWORD g_dwDebugFlags = DEFAULT_DEBUG_FLAGS; // Registry may override this.

// A class that automizes file handle operations.
// It actually can work for any handle returned by the system but in practice it is 
// used here only with file handles.
class CFileHandle
{
public:
    CFileHandle() :
        m_hFile(INVALID_HANDLE_VALUE)
    {
    }
    ~CFileHandle()
    {
        if (m_hFile != INVALID_HANDLE_VALUE)
        {
            CloseHandle(m_hFile);
        }
    }

    inline CFileHandle &operator =(HANDLE hFile) throw()
    {
        ATLASSERT(m_hFile == INVALID_HANDLE_VALUE);
        m_hFile = hFile;
        return *this;
    }

    inline bool operator ==(HANDLE hfile) throw()
    {
        return m_hFile == hfile;
    }

    inline bool operator !=(HANDLE hfile) throw()
    {
        return !operator==(hfile);
    }

    inline operator HANDLE() throw()
    {
        return m_hFile;
    }

private:
    HANDLE m_hFile;
};

// A global handle object of the lof file.
static CFileHandle g_hDebugLogFile;

// Function that generates the trace output.
void __cdecl OutputDebugPrintf(LPCWSTR szFmt, ...)
{
    WCHAR s[2048];
    va_list args;

    // Format the trace string.
    va_start(args, szFmt);
    vswprintf(s, NELEMS(s), szFmt, args);
    if (g_dwDebugFlags & LOG_TO_DEBUGGER)
        // Output trace info to debugger output window
        OutputDebugStringW(s);
    if (g_hDebugLogFile != INVALID_HANDLE_VALUE)
    {
        // Log trace info to log file.
        DWORD dwBytesWritten;
        DWORD dwLen = (DWORD)wcslen(s);
        CHAR mbs[NELEMS(s)];

        if (dwLen && (s[dwLen-1] == L'\n') && ((dwLen == 1) || (s[dwLen-2] != '\r')))
        {
            // Replace a single '\n' at the end of the line to "\r\n" 
            // so it'll show nicely in notepad. It doesn't work of course 
            // for '\n' in the middle of the line.
            s[dwLen-1] = L'\r';
            s[dwLen++] = L'\n';
        }

        // Convert the string to multi-byte string.
        dwLen = WideCharToMultiByte(CP_ACP, 0, s, dwLen, mbs, sizeof(mbs), NULL, NULL);
        // Write trace info to log file
        WriteFile(g_hDebugLogFile, mbs, dwLen, &dwBytesWritten, NULL);
        // Flush file buffers to be crash safe (as much as it depends on us...)
        FlushFileBuffers(g_hDebugLogFile);
    }
}

// Initializes g_dwDebugFlags and log file handle from the registry data.
class CInitDebug
{
public:
    CInitDebug()
    {
        // If anything fails here, g_dwDebugFlags will remain with the default value above.
        CRegKey RegKey;

        if (RegKey.Create(HKEY_CURRENT_USER, ADDIN_REG_KEY_PATH) == ERROR_SUCCESS)
        {
            // Retrieve debug flags from registry.
            RegKey.QueryDWORDValue(DEBUG_FLAGS_REGNAME, g_dwDebugFlags);
        }

        if (g_dwDebugFlags & LOG_TO_FILE)
        {
            // Retrieve the log file name, if this fails we remain with the default fine name,
            WCHAR wszDebugLogFileName[MAX_PATH];
            ULONG ulDebugLogFileNameLen = (ULONG)(sizeof(wszDebugLogFileName)/sizeof(WCHAR));

            if (RegKey.QueryStringValue(
                    DEBUG_LOG_FILE_REGNAME, 
                    wszDebugLogFileName, 
                    &ulDebugLogFileNameLen) != ERROR_SUCCESS)
            {
                wcscpy(wszDebugLogFileName, DEFAULT_LOG_FILE_NAME);
            }

            // Create and/or Open the log file.
            g_hDebugLogFile = CreateFile(
                wszDebugLogFileName, 
                FILE_WRITE_DATA, 
                FILE_SHARE_READ, 
                NULL, 
                (g_dwDebugFlags & APPEND_LOG_FILE)  ? OPEN_EXISTING : CREATE_ALWAYS, 
                FILE_ATTRIBUTE_NORMAL,
                NULL);

            if (g_hDebugLogFile == INVALID_HANDLE_VALUE)
            {
                if ((g_dwDebugFlags & APPEND_LOG_FILE) && (GetLastError() == ERROR_FILE_NOT_FOUND))
                {
                    g_hDebugLogFile = CreateFile(
                        wszDebugLogFileName, 
                        FILE_WRITE_DATA, 
                        FILE_SHARE_READ, 
                        NULL, 
                        CREATE_NEW, 
                        FILE_ATTRIBUTE_NORMAL,
                        NULL);
                }
            }
            else
            {
                // In case we append to previous log, move the file pointer to the end of the file.
                SetFilePointer(g_hDebugLogFile, 0, NULL, FILE_END);
            }

            if (g_hDebugLogFile == INVALID_HANDLE_VALUE)
            {
                DEBUG_TRACE(
                    TRACE_ERRORS,
                    (L"Failed to Create debug log file, Error = %d\n", GetLastError()));
            }
        }
        else
        {
            g_hDebugLogFile = INVALID_HANDLE_VALUE;
        }
    }
};

// Create an instance of CInitDebug to initialize the debug trace info.
static CInitDebug g_InitDebug;

#pragma warning (default:4996)
