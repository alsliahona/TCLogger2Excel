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

#pragma once
#include "afxwin.h"
#include <string>

//
// For enforcing only .0-9
void OnEnChangeEditDecimalOnly(CEdit &edit);
void OnEnChangeEditTextNoComma(CEdit &edit);

// AddCastingDialog dialog
class AddCastingDialog : public CDialogEx
{
	DECLARE_DYNAMIC(AddCastingDialog)

public:
	AddCastingDialog(CWnd* pParent = nullptr);   // standard constructor
	virtual ~AddCastingDialog();

	const std::wstring& GetTubeName() const
	{
		return m_wstrTubeName;
	}

	double GetTubeWeight() const
	{
		return m_dTubeWeight;
	}

	double GetTubeDiameter() const
	{
		return m_dTubeDiameter;
	}

// Dialog Data
	enum { IDD = IDD_DIALOG_ADDCASTING };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()
private:
	CEdit m_editTubeDiameter;
	CEdit m_editTubeWeight;
	CEdit m_editTubeName;
	afx_msg void OnEnKillfocusEditCtdiameter();
	afx_msg void OnEnKillfocusEditCtweight();
	afx_msg void OnEnKillfocusEditCtname();
	afx_msg void OnEnChangeEditCtdiameter();
	afx_msg void OnEnChangeEditCtweight();

	double m_dTubeDiameter{ 0.0f };
	double m_dTubeWeight{ 0.0f };
	std::wstring m_wstrTubeName;
public:
	afx_msg void OnEnChangeEditCtname();
};
