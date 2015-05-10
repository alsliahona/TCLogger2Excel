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

// TCLogger2ExcelDlg.cpp : implementation file
//

#include "stdafx.h"
#include "boost/algorithm/string.hpp"
#include "TCLogger2Excel.h"
#include "TCLogger2ExcelDlg.h"
#include "afxdialogex.h"
#include <string>
#include <array>
#include <vector>
#include <map>
#include "AddCastingDialog.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// CAboutDlg dialog used for App About

class CAboutDlg : public CDialogEx
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

CAboutDlg::CAboutDlg() : CDialogEx(CAboutDlg::IDD)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialogEx)
END_MESSAGE_MAP()


// TCLogger2ExcelDlg dialog



TCLogger2ExcelDlg::TCLogger2ExcelDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(TCLogger2ExcelDlg::IDD, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void TCLogger2ExcelDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_LIST_CTUBES, m_CTubeListBox);
	DDX_Control(pDX, IDC_EDIT_PMAX, m_editPmax);
	DDX_Control(pDX, IDC_EDIT_GRAINWT, m_editGrainWeight);
	DDX_Control(pDX, IDC_EDIT_GRAINLEN, m_editGrainLength);
	DDX_Control(pDX, IDC_EDIT_GRAINDIA, m_editGrainDiameter);
	DDX_Control(pDX, IDC_EDIT_GRAINCORE, m_editGrainCore);
}

BEGIN_MESSAGE_MAP(TCLogger2ExcelDlg, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDC_BUTTON_ADDCTUBE, &TCLogger2ExcelDlg::OnBnClickedAddCTube)
	ON_BN_CLICKED(IDC_BUTTON_REMOVECTUBE, &TCLogger2ExcelDlg::OnBnClickedButtonRemoveCTube)
	ON_EN_CHANGE(IDC_EDIT_PMAX, &TCLogger2ExcelDlg::OnEnChangeEditPmax)
	ON_EN_CHANGE(IDC_EDIT_GRAINWT, &TCLogger2ExcelDlg::OnEnChangeEditGrainwt)
	ON_EN_CHANGE(IDC_EDIT_GRAINLEN, &TCLogger2ExcelDlg::OnEnChangeEditGrainlen)
	ON_EN_CHANGE(IDC_EDIT_GRAINDIA, &TCLogger2ExcelDlg::OnEnChangeEditGraindia)
	ON_EN_CHANGE(IDC_EDIT_GRAINCORE, &TCLogger2ExcelDlg::OnEnChangeEditGraincore)
	ON_LBN_SELCHANGE(IDC_LIST_CTUBES, &TCLogger2ExcelDlg::OnLbnSelchangeListCtubes)
	ON_BN_CLICKED(IDOK, &TCLogger2ExcelDlg::OnBnClickedOk)
END_MESSAGE_MAP()


// TCLogger2ExcelDlg message handlers

BOOL TCLogger2ExcelDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// Add "About..." menu item to system menu.

	// IDM_ABOUTBOX must be in the system command range.
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (nullptr != pSysMenu)
	{
		BOOL bNameValid;
		CString strAboutMenu;
		bNameValid = strAboutMenu.LoadString(IDS_ABOUTBOX);
		ASSERT(bNameValid);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// Set the icon for this dialog.  The framework does this automatically
	//  when the application's main window is not a dialog
	// SetIcon(m_hIcon, TRUE);			// Set big icon
	// SetIcon(m_hIcon, FALSE);		// Set small icon
	std::wstring wstrPmax = g_theApp.GetProfileStringW(L"TCLogger2Excel", L"PmaxThreshold", L"10");
	SetDlgItemTextW(IDC_EDIT_PMAX, wstrPmax.c_str());
	std::wstring wstrGrainWt = g_theApp.GetProfileStringW(L"TCLogger2Excel", L"GrainWeightDefault", L"175.0");
	SetDlgItemTextW(IDC_EDIT_GRAINWT, wstrGrainWt.c_str());
	std::wstring wstrGrainLength = g_theApp.GetProfileStringW(L"TCLogger2Excel", L"GrainLengthDefault", L"3.0625");
	SetDlgItemTextW(IDC_EDIT_GRAINLEN, wstrGrainLength.c_str());
	std::wstring wstrGrainDiameter = g_theApp.GetProfileStringW(L"TCLogger2Excel", L"GrainDiameterDefault", L"1.75");
	SetDlgItemTextW(IDC_EDIT_GRAINDIA, wstrGrainDiameter.c_str());
	std::wstring wstrGrainCore = g_theApp.GetProfileStringW(L"TCLogger2Excel", L"GrainCoreDefault", L"0.625");
	SetDlgItemTextW(IDC_EDIT_GRAINCORE, wstrGrainCore.c_str());

	//
	// Setup Defaults if no entries...
	std::wstring wstrZero(g_theApp.GetProfileStringW(L"TCLogger2Excel", L"CTubeConfig_0", L""));
	std::wstring wstrOne(g_theApp.GetProfileStringW(L"TCLogger2Excel", L"CTubeConfig_1", L""));
	std::wstring wstrTwo(g_theApp.GetProfileStringW(L"TCLogger2Excel", L"CTubeConfig_2", L""));
	if(wstrZero.empty() && wstrOne.empty() && wstrTwo.empty())
	{
		g_theApp.WriteProfileStringW(L"TCLogger2Excel", L"CTubeConfig_0", L"Tru-Core 54mm,1.75,3.030928");
		g_theApp.WriteProfileStringW(L"TCLogger2Excel", L"CTubeConfig_1", L"AlwaysReady 54mm,1.75,1.9197342");
		g_theApp.WriteProfileStringW(L"TCLogger2Excel", L"CTubeConfig_2", L"AlwaysReady 75mm,1.75,3.060583");
	}

	//
	// Read up to 100 casting tube configurations from the registry...
	UINT uiLastTube = g_theApp.GetProfileIntW(L"TCLogger2Excel", L"LastUsedCastingTube", 0);
	int iSelectItem = 0;
	for(UINT uiIndex = 0; uiIndex < 100; ++uiIndex)
	{
		std::wstring wstrCTubeConfig(L"CTubeConfig_");
		wstrCTubeConfig += std::to_wstring(uiIndex);
		std::wstring wstrCTubeConfigLine(g_theApp.GetProfileStringW(L"TCLogger2Excel", wstrCTubeConfig.c_str(), L""));
		if(wstrCTubeConfigLine.empty())
		{
			continue;		// A removed or non-existent entry
		}

		//
		//
		std::vector<std::wstring> vEntries;
		boost::algorithm::split(vEntries, wstrCTubeConfigLine, boost::is_any_of(L","));
		if(vEntries.size() != 3)
		{
			continue;		// @TODO: Warn user of invalid entry, shouldn't be possible though
		}
		std::wstring wstrCTube(vEntries[0] + L" - ");
		wstrCTube += (vEntries[1] + L"\" O.D.: ");
		wstrCTube += (vEntries[2] + L" g/in.");
		int iItem = m_CTubeListBox.AddString(wstrCTube.c_str());
		if(uiIndex == uiLastTube)
		{
			iSelectItem = iItem;
		}
		m_CTubeListBox.SetItemData(iItem, uiIndex);
	}
	m_CTubeListBox.SetCurSel(iSelectItem);

	return TRUE;  // return TRUE  unless you set the focus to a control
}

void TCLogger2ExcelDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialogEx::OnSysCommand(nID, lParam);
	}
}

// If you add a minimize button to your dialog, you will need the code below
//  to draw the icon.  For MFC applications using the document/view model,
//  this is automatically done for you by the framework.

void TCLogger2ExcelDlg::OnPaint()
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
		CDialogEx::OnPaint();
	}
}

// The system calls this function to obtain the cursor to display while the user drags
//  the minimized window.
HCURSOR TCLogger2ExcelDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

void TCLogger2ExcelDlg::OnBnClickedAddCTube()
{
	AddCastingDialog dlgCastingTube;
	if(IDOK != dlgCastingTube.DoModal())
	{
		return;
	}
	const std::wstring wstrName(dlgCastingTube.GetTubeName());
	double dDiameter = dlgCastingTube.GetTubeDiameter();
	double dWeight = dlgCastingTube.GetTubeWeight();

	//
	// Make a new entry
	std::wstring wstrCTube(wstrName + L" - ");
	wstrCTube += std::to_wstring(dDiameter);
	wstrCTube += L"\" O.D.: ";
	wstrCTube += std::to_wstring(dWeight);
	wstrCTube += L" g/in.";

	//
	// Save in first empty registry slot:
	UINT uiIndex = 0;
	for(uiIndex = 0; uiIndex < 100; ++uiIndex)
	{
		std::wstring wstrCTubeConfig(L"CTubeConfig_");
		wstrCTubeConfig += std::to_wstring(uiIndex);
		std::wstring wstrCTubeConfigLine(g_theApp.GetProfileStringW(L"TCLogger2Excel", wstrCTubeConfig.c_str(), L""));
		if(wstrCTubeConfigLine.empty())
		{
			break;
		}
	}

	if(uiIndex > 99)
	{
		MessageBoxW(L"There are too many Casting Tube entries, please delete one before adding another.", L"Too many casting tubes...", MB_ICONEXCLAMATION);
		return;
	}
	int iItem = m_CTubeListBox.AddString(wstrCTube.c_str());
	std::wstring wstrEntry(wstrName + L",");
	wstrEntry += std::to_wstring(dDiameter);
	wstrEntry += L",";
	wstrEntry += std::to_wstring(dWeight);
	m_CTubeListBox.SetItemData(iItem, uiIndex);
	std::wstring wstrCTubeConfig(L"CTubeConfig_");
	wstrCTubeConfig += std::to_wstring(uiIndex);
	g_theApp.WriteProfileStringW(L"TCLogger2Excel", wstrCTubeConfig.c_str(), wstrEntry.c_str());
}

void TCLogger2ExcelDlg::OnBnClickedButtonRemoveCTube()
{
	int iItem = m_CTubeListBox.GetCurSel();
	UINT uiIndex = m_CTubeListBox.GetItemData(iItem);
	CString strItem;
	m_CTubeListBox.GetText(iItem, strItem);
	if(IDNO == MessageBoxW(strItem.GetString(), L"Delete CTube?", MB_YESNO|MB_ICONQUESTION))
	{
		return;
	}
	m_CTubeListBox.DeleteString(iItem);
	std::wstring wstrCTubeConfig(L"CTubeConfig_");
	wstrCTubeConfig += std::to_wstring(uiIndex);
	g_theApp.WriteProfileStringW(L"TCLogger2Excel", wstrCTubeConfig.c_str(), L"");
}


void TCLogger2ExcelDlg::OnEnChangeEditPmax()
{
	OnEnChangeEditDecimalOnly(m_editPmax);
}


void TCLogger2ExcelDlg::OnEnChangeEditGrainwt()
{
	OnEnChangeEditDecimalOnly(m_editGrainWeight);
}


void TCLogger2ExcelDlg::OnEnChangeEditGrainlen()
{
	OnEnChangeEditDecimalOnly(m_editGrainLength);
}


void TCLogger2ExcelDlg::OnEnChangeEditGraindia()
{
	OnEnChangeEditDecimalOnly(m_editGrainDiameter);
}


void TCLogger2ExcelDlg::OnEnChangeEditGraincore()
{
	OnEnChangeEditDecimalOnly(m_editGrainCore);
}


void TCLogger2ExcelDlg::OnLbnSelchangeListCtubes()
{
	//
	// Update the grain diameter
	int iItem = m_CTubeListBox.GetCurSel();
	UINT uiIndex = m_CTubeListBox.GetItemData(iItem);

	std::wstring wstrCTubeConfig(L"CTubeConfig_");
	wstrCTubeConfig += std::to_wstring(uiIndex);
	std::wstring wstrCTubeConfigLine(g_theApp.GetProfileStringW(L"TCLogger2Excel", wstrCTubeConfig.c_str(), L""));
	if(wstrCTubeConfigLine.empty())
	{
		return;
	}

	std::vector<std::wstring> vEntries;
	boost::algorithm::split(vEntries, wstrCTubeConfigLine, boost::is_any_of(L","));
	if(vEntries.size() != 3)
	{
		return;		// @TODO: Warn user of invalid entry, shouldn't be possible though
	}

	std::wstring wstrDiameter(vEntries[1]);
	m_editGrainDiameter.SetWindowTextW(vEntries[1].c_str());
}


void TCLogger2ExcelDlg::OnBnClickedOk()
{
	//
	// Get Values
	std::array<wchar_t, 1024> wzTemp;
	m_editPmax.GetWindowTextW(&wzTemp[0], wzTemp.size());
	const std::wstring wstrPmaxThreshold(&wzTemp[0]);

	m_editGrainWeight.GetWindowTextW(&wzTemp[0], wzTemp.size());
	const std::wstring wstrGrainWeight(&wzTemp[0]);

	m_editGrainLength.GetWindowTextW(&wzTemp[0], wzTemp.size());
	const std::wstring wstrGrainLength(&wzTemp[0]);

	m_editGrainDiameter.GetWindowTextW(&wzTemp[0], wzTemp.size());
	const std::wstring wstrGrainDiameter(&wzTemp[0]);

	m_editGrainCore.GetWindowTextW(&wzTemp[0], wzTemp.size());
	const std::wstring wstrGrainCore(&wzTemp[0]);

	UINT uiIndex = 0;
	std::wstring wstrCastingTubeWeight(L"3.030928");
	{
		int iItem = m_CTubeListBox.GetCurSel();
		uiIndex = m_CTubeListBox.GetItemData(iItem);

		std::wstring wstrCTubeConfig(L"CTubeConfig_");
		wstrCTubeConfig += std::to_wstring(uiIndex);
		std::wstring wstrCTubeConfigLine(g_theApp.GetProfileStringW(L"TCLogger2Excel", wstrCTubeConfig.c_str(), L""));
		if(wstrCTubeConfigLine.empty())
		{
			wstrCTubeConfigLine = L"Default 54mm,1.75,3.030928";
		}

		std::vector<std::wstring> vwstrCastingTubeElements;
		boost::algorithm::split(vwstrCastingTubeElements, wstrCTubeConfigLine, boost::is_any_of(L","));
		if(vwstrCastingTubeElements.size() == 3)
		{
			wstrCastingTubeWeight = vwstrCastingTubeElements[2];
		}
	}

	//
	// Update Registry
	g_theApp.WriteProfileStringW(L"TCLogger2Excel", L"PmaxThreshold", wstrPmaxThreshold.c_str());
	g_theApp.WriteProfileStringW(L"TCLogger2Excel", L"GrainWeightDefault", wstrGrainWeight.c_str());
	g_theApp.WriteProfileStringW(L"TCLogger2Excel", L"GrainLengthDefault", wstrGrainLength.c_str());
	g_theApp.WriteProfileStringW(L"TCLogger2Excel", L"GrainDiameterDefault", wstrGrainDiameter.c_str());
	g_theApp.WriteProfileStringW(L"TCLogger2Excel", L"GrainCoreDefault", wstrGrainCore.c_str());
	g_theApp.WriteProfileInt(L"TCLogger2Excel", L"LastUsedCastingTube", uiIndex);

	//
	// Select Files and Populate Excel
	extern void CreateExcelSpreadsheet(std::map<std::wstring, std::wstring> &mapProperties);
	std::map<std::wstring, std::wstring> mapProperties
	{
		{ L"PmaxThreshold", wstrPmaxThreshold },
		{ L"GrainWeight", wstrGrainWeight },
		{ L"GrainLength", wstrGrainLength },
		{ L"GrainDiameter", wstrGrainDiameter },
		{ L"GrainCore", wstrGrainCore },
		{ L"CastingTubeWeight", wstrCastingTubeWeight }
	};
	CreateExcelSpreadsheet(mapProperties);

	//
	// Dialog wrap-up
	CDialogEx::OnOK();
}
