// Copyright(c) 2015 Andrew L. Sandoval (http://www.andrewlsandoval.com)
// 
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files(the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and / or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions :
// 
// The above copyright notice and this permission notice shall be included in
// all copies or substantial portions of the Software.
// 
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
// THE SOFTWARE.

// TCLogger2ExcelDlg.h : header file
//

#pragma once
#include "afxwin.h"


// TCLogger2ExcelDlg dialog
class TCLogger2ExcelDlg : public CDialogEx
{
// Construction
public:
	TCLogger2ExcelDlg(CWnd* pParent = nullptr);	// standard constructor

	// Dialog Data
	enum { IDD = IDD_TCLOGGER2EXCEL_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV support

// Implementation
protected:
	HICON m_hIcon;

	// Generated message map functions
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()
private:
	afx_msg void OnBnClickedAddCTube();
	afx_msg void OnBnClickedButtonRemoveCTube();
	afx_msg void OnEnChangeEditPmax();
	afx_msg void OnEnChangeEditGrainwt();
	afx_msg void OnEnChangeEditGrainlen();
	afx_msg void OnEnChangeEditGraindia();
	afx_msg void OnEnChangeEditGraincore();
	afx_msg void OnLbnSelchangeListCtubes();
	CListBox m_CTubeListBox;
	CEdit m_editPmax;
	CEdit m_editGrainWeight;
	CEdit m_editGrainLength;
	CEdit m_editGrainDiameter;
	CEdit m_editGrainCore;
public:
	afx_msg void OnBnClickedOk();
};
