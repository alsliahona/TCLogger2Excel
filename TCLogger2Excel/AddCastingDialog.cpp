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

// AddCastingDialog.cpp : implementation file
//

#include "stdafx.h"
#include "TCLogger2Excel.h"
#include "AddCastingDialog.h"
#include "afxdialogex.h"
#include <array>


// AddCastingDialog dialog

IMPLEMENT_DYNAMIC(AddCastingDialog, CDialogEx)

AddCastingDialog::AddCastingDialog(CWnd* pParent)
	: CDialogEx(AddCastingDialog::IDD, pParent)
{
}

AddCastingDialog::~AddCastingDialog()
{
}

void AddCastingDialog::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_EDIT_CTDIAMETER, m_editTubeDiameter);
	DDX_Control(pDX, IDC_EDIT_CTWEIGHT, m_editTubeWeight);
	DDX_Control(pDX, IDC_EDIT_CTNAME, m_editTubeName);
}


BEGIN_MESSAGE_MAP(AddCastingDialog, CDialogEx)
	ON_EN_KILLFOCUS(IDC_EDIT_CTDIAMETER, &AddCastingDialog::OnEnKillfocusEditCtdiameter)
	ON_EN_KILLFOCUS(IDC_EDIT_CTWEIGHT, &AddCastingDialog::OnEnKillfocusEditCtweight)
	ON_EN_CHANGE(IDC_EDIT_CTDIAMETER, &AddCastingDialog::OnEnChangeEditCtdiameter)
	ON_EN_CHANGE(IDC_EDIT_CTWEIGHT, &AddCastingDialog::OnEnChangeEditCtweight)
	ON_EN_KILLFOCUS(IDC_EDIT_CTNAME, &AddCastingDialog::OnEnKillfocusEditCtname)
	ON_EN_CHANGE(IDC_EDIT_CTNAME, &AddCastingDialog::OnEnChangeEditCtname)
END_MESSAGE_MAP()


// AddCastingDialog message handlers


void AddCastingDialog::OnEnKillfocusEditCtdiameter()
{
	//
	// Validate:
	std::array<wchar_t, 1024> wzText;
	if(m_editTubeDiameter.GetWindowTextW(&wzText[0], static_cast<int>(wzText.size())) > 0)
	{
		m_dTubeDiameter = std::stod(&wzText[0]);
	}
	else
	{
		m_dTubeDiameter = 0.0f;
	}
}

void AddCastingDialog::OnEnKillfocusEditCtweight()
{
	//
	// Validate:
	std::array<wchar_t, 1024> wzText;
	if(m_editTubeWeight.GetWindowTextW(&wzText[0], static_cast<int>(wzText.size())) > 0)
	{
		m_dTubeWeight = std::stod(&wzText[0]);
	}
	else
	{
		m_dTubeWeight = 0.0f;
	}
}

void OnEnChangeEditDecimalOnly(CEdit &edit)
{
	//
	// Validate:
	std::array<wchar_t, 1024> wzText;
	edit.GetWindowTextW(&wzText[0], static_cast<int>(wzText.size()));
	std::wstring wstrText(&wzText[0]);
	std::wstring wstrSanitized;
	bool bSanityFailed = false;
	for(auto wc : wstrText)
	{
		if(wc == L'.' || (wc >= L'0' && wc <= L'9'))
		{
			wstrSanitized += wc;
		}
		else
		{
			bSanityFailed = true;
			MessageBeep(-1);
		}
	}
	if(true == bSanityFailed)
	{
		edit.SetWindowTextW(wstrSanitized.c_str());
		edit.SetSel(0, -1);
		edit.SetSel(-1);
	}
}

void OnEnChangeEditTextNoComma(CEdit &edit)
{
	//
	// Validate:
	std::array<wchar_t, 1024> wzText;
	edit.GetWindowTextW(&wzText[0], static_cast<int>(wzText.size()));
	std::wstring wstrText(&wzText[0]);
	std::wstring wstrSanitized;
	bool bSanityFailed = false;
	for(auto wc : wstrText)
	{
		if(wc != L',')
		{
			wstrSanitized += wc;
		}
		else
		{
			bSanityFailed = true;
			MessageBeep(-1);
		}
	}
	if(true == bSanityFailed)
	{
		edit.SetWindowTextW(wstrSanitized.c_str());
		edit.SetSel(0, -1);
		edit.SetSel(-1);
	}
}

void AddCastingDialog::OnEnChangeEditCtdiameter()
{
	//
	// Validate:
	OnEnChangeEditDecimalOnly(m_editTubeDiameter);
}


void AddCastingDialog::OnEnChangeEditCtweight()
{
	OnEnChangeEditDecimalOnly(m_editTubeWeight);
}

void AddCastingDialog::OnEnKillfocusEditCtname()
{
	std::array<wchar_t, 1024> wzText;
	m_editTubeName.GetWindowTextW(&wzText[0], static_cast<int>(wzText.size()));
	m_wstrTubeName = &wzText[0];
}


void AddCastingDialog::OnEnChangeEditCtname()
{
	OnEnChangeEditTextNoComma(m_editTubeName);
}
