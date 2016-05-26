// DbgConfigDlg.cpp : implementation file
//

#include "stdafx.h"
#include "DbgConfig.h"
#include "DbgConfigDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

DWORD gdwDebugFlags;


// CAboutDlg dialog used for App About

class CAboutDlg : public CDialog
{
public:
    CAboutDlg();

// Dialog Data
    enum { IDD = IDD_ABOUTBOX };

    protected:
    virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

// Implementation
protected:
    DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialog(CAboutDlg::IDD)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
    CDialog::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialog)
END_MESSAGE_MAP()


// CDbgConfigDlg dialog




CDbgConfigDlg::CDbgConfigDlg(CWnd* pParent /*=NULL*/)
    : CDialog(CDbgConfigDlg::IDD, pParent)
{
    m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CDbgConfigDlg::DoDataExchange(CDataExchange* pDX)
{
    CDialog::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CDbgConfigDlg, CDialog)
    ON_WM_SYSCOMMAND()
    ON_WM_PAINT()
    ON_WM_QUERYDRAGICON()
    //}}AFX_MSG_MAP
    ON_BN_CLICKED(IDC_CHECK_OUTPUT_FILE, &CDbgConfigDlg::OnBnClickedCheckOutputFile)
    ON_BN_CLICKED(IDC_CHECK_OUTPUT_DEBUG, &CDbgConfigDlg::OnBnClickedCheckOutputDebug)
    ON_BN_CLICKED(IDC_CHECK_APPEND_LOG_FILE, &CDbgConfigDlg::OnBnClickedCheckAppendLogFile)
    ON_EN_CHANGE(IDC_EDIT_LOG_FILE_NAME, &CDbgConfigDlg::OnEnChangeEditLogFileName)
    ON_BN_CLICKED(IDOK, &CDbgConfigDlg::OnBnClickedOk)
    ON_BN_CLICKED(IDC_BUTTON_BROWSE_LOG_FILE, &CDbgConfigDlg::OnBnClickedButtonBrowseLogFile)
    ON_BN_CLICKED(IDC_RADIO_NONE, &CDbgConfigDlg::OnBnClickedRadioNone)
    ON_BN_CLICKED(IDC_RADIO_ERRORS, &CDbgConfigDlg::OnBnClickedRadioErrors)
    ON_BN_CLICKED(IDC_RADIO_INFO, &CDbgConfigDlg::OnBnClickedRadioInfo)
    ON_BN_CLICKED(IDC_RADIO_VERBOSE, &CDbgConfigDlg::OnBnClickedRadioVerbose)
END_MESSAGE_MAP()


// CDbgConfigDlg message handlers

BOOL CDbgConfigDlg::OnInitDialog()
{
    CDialog::OnInitDialog();

    // Add "About..." menu item to system menu.

    // IDM_ABOUTBOX must be in the system command range.
    ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
    ASSERT(IDM_ABOUTBOX < 0xF000);

    CMenu* pSysMenu = GetSystemMenu(FALSE);
    if (pSysMenu != NULL)
    {
        CString strAboutMenu;
        strAboutMenu.LoadString(IDS_ABOUTBOX);
        if (!strAboutMenu.IsEmpty())
        {
            pSysMenu->AppendMenu(MF_SEPARATOR);
            pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
        }
    }

    // Set the icon for this dialog.  The framework does this automatically
    //  when the application's main window is not a dialog
    SetIcon(m_hIcon, TRUE);         // Set big icon
    SetIcon(m_hIcon, FALSE);        // Set small icon

    if (RegCreateKeyEx(HKEY_CURRENT_USER, ADDIN_REG_KEY_PATH, 0, NULL, 0, KEY_ALL_ACCESS , NULL, &m_hDbgRegKey, NULL) != ERROR_SUCCESS)
    {
        MessageBox(_T("Could not access registry!"), _T("Error"), MB_OK | MB_ICONEXCLAMATION);
        return FALSE;
    }

    DWORD dwType;
    DWORD dwSize = (DWORD)sizeof(gdwDebugFlags);

    if (RegQueryValueEx(m_hDbgRegKey, DEBUG_FLAGS_REGNAME, 0, &dwType, (LPBYTE)&gdwDebugFlags, &dwSize) != ERROR_SUCCESS)
    {
        gdwDebugFlags = DEFAULT_DEBUG_FLAGS;
    }

    dwSize = (DWORD)sizeof(m_szLogFileName);

    if (RegQueryValueEx(m_hDbgRegKey, DEBUG_LOG_FILE_REGNAME, 0, &dwType, (LPBYTE)m_szLogFileName, &dwSize) != ERROR_SUCCESS)
    {
        wcscpy(m_szLogFileName, DEFAULT_LOG_FILE_NAME);
    }

    if (gdwDebugFlags & TRACE_VERBOSE_INFO)
    {
        CButton *pBtn = (CButton *)GetDlgItem(IDC_RADIO_VERBOSE);
        pBtn->SetCheck(TRUE);
        m_dwTraceInfoType = TRACE_VERBOSE_INFO | TRACE_INFO | TRACE_ERRORS;
    }
    else
    {
        if (gdwDebugFlags & TRACE_INFO)
        {
            CButton *pBtn = (CButton *)GetDlgItem(IDC_RADIO_INFO);
            pBtn->SetCheck(TRUE);
            m_dwTraceInfoType = TRACE_INFO | TRACE_ERRORS;
        }
        else
        {
            if (gdwDebugFlags & TRACE_ERRORS)
            {
                CButton *pBtn = (CButton *)GetDlgItem(IDC_RADIO_ERRORS);
                pBtn->SetCheck(TRUE);
                m_dwTraceInfoType = TRACE_ERRORS;
            }
            else
            {
                CButton *pBtn = (CButton *)GetDlgItem(IDC_RADIO_NONE);
                pBtn->SetCheck(TRUE);
                m_dwTraceInfoType = 0;
            }
        }
    }

    m_pLogToDebuggerBtn = (CButton *)GetDlgItem(IDC_CHECK_OUTPUT_DEBUG);
    m_pLogToDebuggerBtn->SetCheck(m_bDebugLog = ((gdwDebugFlags & LOG_TO_DEBUGGER) != 0));

    m_pLogToFileBtn = (CButton *)GetDlgItem(IDC_CHECK_OUTPUT_FILE);
    m_pLogToFileBtn->SetCheck(m_bFileLog = ((gdwDebugFlags & LOG_TO_FILE) != 0));

    m_pAppendLogFileBtn = (CButton *)GetDlgItem(IDC_CHECK_APPEND_LOG_FILE);
    m_pAppendLogFileBtn->SetCheck(m_bAppendLog = ((gdwDebugFlags & APPEND_LOG_FILE) != 0));
    m_pAppendLogFileBtn->EnableWindow(m_bFileLog);

    m_pLogFileNameEdit = (CEdit *)GetDlgItem(IDC_EDIT_LOG_FILE_NAME);
    m_pLogFileNameEdit->SetWindowText(m_szLogFileName);
    m_pLogFileNameEdit->EnableWindow(m_bFileLog);

    m_pStaticFileName = GetDlgItem(IDC_STATIC_FILE_NAME);
    m_pStaticFileName->EnableWindow(m_bFileLog);

    m_pBrowseFileNameBtn = (CButton *)GetDlgItem(IDC_BUTTON_BROWSE_LOG_FILE);
    m_pBrowseFileNameBtn->EnableWindow(m_bFileLog);
    
    return TRUE;  // return TRUE  unless you set the focus to a control
}

void CDbgConfigDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
    if ((nID & 0xFFF0) == IDM_ABOUTBOX)
    {
        CAboutDlg dlgAbout;
        dlgAbout.DoModal();
    }
    else
    {
        CDialog::OnSysCommand(nID, lParam);
    }
}

// If you add a minimize button to your dialog, you will need the code below
//  to draw the icon.  For MFC applications using the document/view model,
//  this is automatically done for you by the framework.

void CDbgConfigDlg::OnPaint()
{
    if (IsIconic())
    {
        CPaintDC dc(this); // device context for painting

        SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

        // Center icon in client rectangle
        int cxIcon = GetSystemMetrics(SM_CXICON);
        int cyIcon = GetSystemMetrics(SM_CYICON);
        CRect rect;
        GetClientRect(&rect);
        int x = (rect.Width() - cxIcon + 1) / 2;
        int y = (rect.Height() - cyIcon + 1) / 2;

        // Draw the icon
        dc.DrawIcon(x, y, m_hIcon);
    }
    else
    {
        CDialog::OnPaint();
    }
}

// The system calls this function to obtain the cursor to display while the user drags
//  the minimized window.
HCURSOR CDbgConfigDlg::OnQueryDragIcon()
{
    return static_cast<HCURSOR>(m_hIcon);
}

void CDbgConfigDlg::OnBnClickedCheckOutputFile()
{
    m_bFileLog = m_pLogToFileBtn->GetCheck();
    m_pAppendLogFileBtn->EnableWindow(m_bFileLog);
    m_pLogFileNameEdit->EnableWindow(m_bFileLog);
    m_pStaticFileName->EnableWindow(m_bFileLog);
    m_pBrowseFileNameBtn->EnableWindow(m_bFileLog);
    if (m_bFileLog)
        gdwDebugFlags |= LOG_TO_FILE;
    else
        gdwDebugFlags &= ~LOG_TO_FILE;
}

void CDbgConfigDlg::OnBnClickedCheckOutputDebug()
{
    m_bDebugLog = m_pLogToDebuggerBtn->GetCheck();
    if (m_bDebugLog)
        gdwDebugFlags |= LOG_TO_DEBUGGER;
    else
        gdwDebugFlags &= ~LOG_TO_DEBUGGER;
}

void CDbgConfigDlg::OnBnClickedCheckAppendLogFile()
{
    m_bAppendLog = m_pAppendLogFileBtn->GetCheck();
    if (m_bAppendLog)
        gdwDebugFlags |= APPEND_LOG_FILE;
    else
        gdwDebugFlags &= ~APPEND_LOG_FILE;
}

void CDbgConfigDlg::OnEnChangeEditLogFileName()
{
    m_pLogFileNameEdit->GetWindowText(m_szLogFileName, MAX_PATH);
}

void CDbgConfigDlg::OnBnClickedOk()
{
    gdwDebugFlags = (gdwDebugFlags & ~(TRACE_ERRORS | TRACE_INFO | TRACE_VERBOSE_INFO)) | m_dwTraceInfoType;
    RegSetValueEx(m_hDbgRegKey, DEBUG_FLAGS_REGNAME, 0, REG_DWORD, (LPBYTE)&gdwDebugFlags, (DWORD)sizeof(gdwDebugFlags));
    RegSetValueEx(m_hDbgRegKey, DEBUG_LOG_FILE_REGNAME, 0, REG_SZ, (LPBYTE)m_szLogFileName, (DWORD)((_tcslen(m_szLogFileName) + 1) * sizeof(TCHAR)));
    OnOK();
}

void CDbgConfigDlg::OnBnClickedButtonBrowseLogFile()
{
    OPENFILENAME ofn;

    memset(&ofn, 0, sizeof(ofn));
    ofn.lStructSize = sizeof(OPENFILENAME);
    ofn.hwndOwner = this->m_hWnd;
    ofn.hInstance = NULL;
    ofn.lpstrFilter = _T("Log file\0*.log\0Text File\0*.txt\0All Files\0*.*\0");
    ofn.lpstrCustomFilter = NULL;
    ofn.nMaxCustFilter = 0;
    ofn.nFilterIndex = 1;
    ofn.lpstrFile = m_szLogFileName;
    ofn.nMaxFile = sizeof(m_szLogFileName) / sizeof(TCHAR);
    ofn.lpstrFileTitle = NULL;
    ofn.nMaxFileTitle = 0;
    ofn.lpstrInitialDir = NULL;
    ofn.lpstrTitle = _T("Select Log File");
    ofn.Flags = OFN_DONTADDTORECENT | OFN_HIDEREADONLY | OFN_NOREADONLYRETURN;
    ofn.lpstrDefExt = _T("log");
    ofn.FlagsEx = 0;
    if (GetSaveFileName(&ofn))
        m_pLogFileNameEdit->SetWindowText(m_szLogFileName);
}

void CDbgConfigDlg::OnBnClickedRadioNone()
{
    m_dwTraceInfoType = 0;
}

void CDbgConfigDlg::OnBnClickedRadioErrors()
{
    m_dwTraceInfoType = TRACE_ERRORS;
}

void CDbgConfigDlg::OnBnClickedRadioInfo()
{
    m_dwTraceInfoType = TRACE_ERRORS | TRACE_INFO;
}

void CDbgConfigDlg::OnBnClickedRadioVerbose()
{
    m_dwTraceInfoType = TRACE_ERRORS | TRACE_INFO | TRACE_VERBOSE_INFO;
}
