// DbgConfigDlg.h : header file
//

#pragma once


// CDbgConfigDlg dialog
class CDbgConfigDlg : public CDialog
{
// Construction
public:
    CDbgConfigDlg(CWnd* pParent = NULL);    // standard constructor

// Dialog Data
    enum { IDD = IDD_DBGCONFIG_DIALOG };

    protected:
    virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support


// Implementation
protected:
    HICON m_hIcon;

    // Generated message map functions
    virtual BOOL OnInitDialog();
    afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
    afx_msg void OnPaint();
    afx_msg HCURSOR OnQueryDragIcon();
    DECLARE_MESSAGE_MAP()
public:
    afx_msg void OnBnClickedCheck1();
public:
    afx_msg void OnBnClickedCheckTraceErrors();

private:
    CButton *m_pLogToDebuggerBtn;
    CButton *m_pLogToFileBtn;
    CButton *m_pAppendLogFileBtn;
    CWnd *m_pStaticFileName;
    CEdit *m_pLogFileNameEdit;
    CButton *m_pBrowseFileNameBtn;

    DWORD m_dwTraceInfoType;
    BOOL m_bDebugLog;
    BOOL m_bFileLog;
    BOOL m_bAppendLog;
    TCHAR m_szLogFileName[MAX_PATH];
    HKEY m_hDbgRegKey;
    afx_msg void OnBnClickedCheckOutputFile();
    afx_msg void OnBnClickedCheckOutputDebug();
    afx_msg void OnBnClickedCheckAppendLogFile();
    afx_msg void OnEnChangeEditLogFileName();
    afx_msg void OnBnClickedOk();
    afx_msg void OnBnClickedButtonBrowseLogFile();
    afx_msg void OnBnClickedRadioNone();
    afx_msg void OnBnClickedRadioErrors();
    afx_msg void OnBnClickedRadioInfo();
    afx_msg void OnBnClickedRadioVerbose();
};
