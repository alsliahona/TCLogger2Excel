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

// TCLoggerToExcel.cpp:
// Andrew L. Sandoval - Copyright (C) April 2015
// Open TCLogger .MTD1 files in Excel, add a graph, mark the Pmax (10% default) line,
// And if three files or more are opened graph the a & n and average delivered ISP...

#include "stdafx.h"
#include <windows.h>
#include <string>
#include <vector>
#include <set>
#include <map>
#include <array>
#include <tuple>
#include <list>
#include <iostream>
#include <locale>
#include <codecvt>
#include <algorithm>
#include <io.h>
#include <stdio.h>
#include <sys/stat.h>
#include "boost/property_tree/xml_parser.hpp"
#include "boost/property_tree/ptree.hpp"
#include "boost/algorithm/string.hpp"
#include <N4189/scope_exit>
#include <N4189/unique_resource>
#include <atlbase.h>

#pragma region Import the type libraries

#import "libid:2DF8D04C-5BFA-101B-BDE5-00AA0044DE52" \
	rename("RGB", "MSORGB") \
	rename("DocumentProperties", "MSODocumentProperties")

using namespace Office;

#import "libid:0002E157-0000-0000-C000-000000000046"

using namespace VBIDE;

#import "libid:00020813-0000-0000-C000-000000000046" \
	rename("DialogBox", "ExcelDialogBox") \
	rename("RGB", "ExcelRGB") \
	rename("CopyFile", "ExcelCopyFile") \
	rename("ReplaceText", "ExcelReplaceText") \
	no_auto_exclude

#pragma endregion

enum eDataField
{
	eTime = 0,
	eThrust,
	ePressure,
	eTimeDelta,
	eThrustElementN
};

const double g_dNewtonsPerPound = 4.44822162f;

//
// Remove characters from a name (e.g. filename) that will break Excel, such as operators
void SanitizeStringName(std::wstring &wstrName)
{
	boost::replace_all(wstrName, L"-", L"_");
	boost::replace_all(wstrName, L"+", L"_");
	boost::replace_all(wstrName, L"*", L"_");
	boost::replace_all(wstrName, L"/", L"_");
	boost::replace_all(wstrName, L"=", L"_");
	boost::replace_all(wstrName, L"$", L"_");
	boost::replace_all(wstrName, L"#", L"_");
	boost::replace_all(wstrName, L"~", L"_");
	boost::replace_all(wstrName, L"^", L"_");
	boost::replace_all(wstrName, L"&", L"_");
	boost::replace_all(wstrName, L"!", L"_");
	boost::replace_all(wstrName, L"@", L"_");
	boost::replace_all(wstrName, L"(", L"_");
	boost::replace_all(wstrName, L")", L"_");
	boost::replace_all(wstrName, L",", L"_");
	boost::replace_all(wstrName, L".", L"_");
	boost::replace_all(wstrName, L"?", L"_");
	boost::replace_all(wstrName, L":", L"_");
	boost::replace_all(wstrName, L"[", L"_");
	boost::replace_all(wstrName, L"]", L"_");
	boost::replace_all(wstrName, L"|", L"_");
}

HRESULT SafeArrayPutDouble(SAFEARRAY& sa, LONG lIndex, LONG lColumn, double dValue)
{
	// Set the first name.
	long lIndice[] = { lIndex, lColumn };
	VARIANT v;
	v.vt = VT_R8;
	v.dblVal = dValue;
	// Copies the VARIANT into the SafeArray
	return SafeArrayPutElement(&sa, lIndice, (void*)&v);
}

HRESULT SafeArrayPutULONG(SAFEARRAY& sa, LONG lIndex, LONG lColumn, ULONG ulValue)
{
	// Set the first name.
	long lIndice[] = { lIndex, lColumn };
	VARIANT v;
	v.vt = VT_UI4;
	v.ulVal = ulValue;
	// Copies the VARIANT into the SafeArray
	return SafeArrayPutElement(&sa, lIndice, (void*)&v);
}

HRESULT SafeArrayPutString(SAFEARRAY& sa, LONG lIndex, LONG lColumn, const std::wstring &wstrString)
{
	// Set the first name.
	long lIndice[] = { lIndex, lColumn };
	VARIANT v;
	v.vt = VT_BSTR;
	v.bstrVal = SysAllocString(wstrString.c_str());
	auto clean = std::experimental::make_scope_exit([&v]() -> void
	{
		VariantClear(&v);
	});
	// Copies the VARIANT into the SafeArray
	return SafeArrayPutElement(&sa, lIndice, (void*)&v);
}

//
// Allow multiple input files and parse the property map to obtain test stand log data
// Create a new workbook for each file in Excel...

bool ObtainInputFiles(std::vector<std::wstring> &vwstrInputFiles)
{
	//
	// Present the open dialog, allow for multiple selections...
	std::vector<wchar_t> wzFilename(8192, 0);
	std::array<wchar_t, 128> wzFilter{ L"TCLogger Data File\0*.MTD1\0All Files\0*.*\0\0" };
	OPENFILENAMEW fn = { 0 };
	fn.lStructSize = sizeof(fn);
	fn.hInstance = nullptr;
	fn.hwndOwner = nullptr;
	fn.lpstrFilter = &wzFilter[0];
	fn.lpstrFile = &wzFilename[0];
	fn.nMaxFile = wzFilename.size();
	fn.lpstrTitle = L"Open TC Logger Data file";
	fn.Flags = OFN_EXPLORER | OFN_FILEMUSTEXIST | OFN_PATHMUSTEXIST | OFN_ALLOWMULTISELECT;

	BOOL bObtainedFiles = GetOpenFileNameW(&fn);
	if(FALSE == bObtainedFiles)
	{
		fprintf(stderr, "No input files selected.\n");
		return 0;
	}
	const wchar_t *pwzFilename = &wzFilename[0];
	const wchar_t *pwzEndOfBuffer = &wzFilename[wzFilename.size() - 1];
	while(pwzFilename < pwzEndOfBuffer && pwzFilename[0] != 0)
	{
		std::wstring wstrFile(pwzFilename);
		pwzFilename += wstrFile.length();
		++pwzFilename;	// Skip null seperator
		errno_t err = _waccess_s(wstrFile.c_str(), 04);
		if(0 == err)
		{
			vwstrInputFiles.push_back(wstrFile);
		}
		else
		{
			//
			// Check for any flags here...
			fprintf(stderr, "%ls is not a valid file.\n", wstrFile.c_str());
		}
	}
	return !vwstrInputFiles.empty();
}

void PlaceStringInCell(Excel::_WorksheetPtr &pXLSheet, wchar_t wcColumn, int row, const std::wstring &wstrString)
{
	std::wstring wstrRange;
	wstrRange += wcColumn;
	wstrRange += std::to_wstring(row);
	Excel::RangePtr pDataRange = pXLSheet->Range[_variant_t(wstrRange.c_str())];
	_ASSERT(pDataRange);
	if(nullptr == pDataRange)
	{
		return;
	}
	pDataRange->Value2 = _bstr_t(wstrString.c_str());
}

double GetPmax(double &dTimeEntry,
	const std::vector<std::tuple<double, double, double, double, double> > &vReadings)
{
	double dPmax = 0.0f;
	dTimeEntry = 0.0f;
	for(auto entry : vReadings)
	{
		double dPressure = std::get<ePressure>(entry);
		if(dPressure > dPmax)
		{
			dPmax = dPressure;
			dTimeEntry = std::get<eTime>(entry);
		}
	}
	return dPmax;
}

bool ProduceExcelWorkbook(Excel::_WorkbookPtr &pXLBook,
	const std::wstring &wstrFilename,
	const std::string &strPropellant,
	bool bMetric,
	const boost::property_tree::ptree &tree,
	const std::vector<std::tuple<double, double, double, double, double> > &vReadings,
	std::map<std::wstring, std::wstring> &mapProperties)
{
	if(vReadings.empty())
	{
		return false;
	}
	// Get the active Worksheet and set its name.
	Excel::_WorksheetPtr pXLSheet = pXLBook->ActiveSheet;
	_ASSERT(pXLSheet);
	if(nullptr == pXLSheet)
	{
		return false;
	}

	//
	// Get Max Pressure to calculate the 10% line
	double dPmaxtime = 0.0f;
	double dPmax = GetPmax(dPmaxtime, vReadings);
	double dpTenPercent = dPmax * 0.10f;

	//
	// Need to know the start and stop burn time based on 10% of Pmax:
	double dBurnStartTime = 0.0f;
	double dBurnEndTime = 0.0f;
	LONG lIndexBurnStart = 0;
	LONG lIndexBurnEnd = 0;
	LONG lExcelIndex = 1;
	bool bHaveEndTime = false;
	for(auto entry : vReadings)
	{
		double dPressure = std::get<ePressure>(entry);
		double dTime = std::get<eTime>(entry);
		++lExcelIndex;
		if(dTime <= dPmaxtime)
		{
			if(dPressure < dpTenPercent)
			{
				continue;
			}
			if(0 == lIndexBurnStart)
			{
				dBurnStartTime = dTime;
				lIndexBurnStart = lExcelIndex;
			}
		}
		else
		{
			// Watching the curve downward...
			lIndexBurnEnd = lExcelIndex;
			if(dPressure < dpTenPercent)
			{
				break;
			}
			dBurnEndTime = dTime;
		}
	}
	double dBurnTime = dBurnEndTime - dBurnStartTime;

	//
	// Convert name to a wide string and assign it to the sheet
	std::wstring wstrTabName(wstrFilename);
	SanitizeStringName(wstrTabName);

	//
	// Remove any math operators from the name:
	_bstr_t bstrTabName(SysAllocString(wstrTabName.c_str()));
	pXLSheet->Name = bstrTabName;

	std::wstring_convert<std::codecvt_utf8_utf16<wchar_t>> converter;
	// std::wstring wstrPropellantName(converter.from_bytes(strPropellant.c_str()));

	//
	// Create seven columns of data: Time, Thrust (LBS), Thrust (N), and Pressure (PSI), ...
	// Construct a safearray of the data
	LONG lIndex = 2;		// Note, 1 based not 0 - will contain # of rows...
	VARIANT saData;
	VARIANT saDataReferences;
	saData.vt = VT_ARRAY | VT_VARIANT;
	saDataReferences.vt = VT_ARRAY | VT_VARIANT;
	{
		SAFEARRAYBOUND sab[2];
		sab[0].lLbound = 1;
		sab[0].cElements = vReadings.size() + 2;
		sab[1].lLbound = 1;
		sab[1].cElements = 7;
		saData.parray = SafeArrayCreate(VT_VARIANT, 2, sab);
		if(saData.parray == nullptr)
		{
			MessageBoxW(HWND_TOP, L"Unable to create safearray for passing data to Excel.", L"Memory error...", MB_TOPMOST | MB_ICONERROR);
			return false;
		}

		//
		// Clean-up safe array when done...
		auto cleanArray = std::experimental::make_scope_exit([&saData]() -> void
		{
			VariantClear(&saData);
		});

		sab[0].lLbound = 1;
		sab[0].cElements = vReadings.size() + 2;
		sab[1].lLbound = 1;
		sab[1].cElements = 5;
		saDataReferences.parray = SafeArrayCreate(VT_VARIANT, 2, sab);
		if(saDataReferences.parray == nullptr)
		{
			MessageBoxW(HWND_TOP, L"Unable to create safearray for passing data to Excel.", L"Memory error...", MB_TOPMOST | MB_ICONERROR);
			return false;
		}

		//
		// Clean-up safe array when done...
		auto cleanRefArray = std::experimental::make_scope_exit([&saDataReferences]() -> void
		{
			VariantClear(&saDataReferences);
		});

		//
		// Labels:
		SafeArrayPutString(*saData.parray, 1, 1, L"Time");
		SafeArrayPutString(*saData.parray, 1, 2, L"Thrust\n(lbs.)");
		SafeArrayPutString(*saData.parray, 1, 3, L"Thrust\n(N)");
		SafeArrayPutString(*saData.parray, 1, 4, L"Pressure\n(PSI)");
		SafeArrayPutString(*saData.parray, 1, 5, L"Threshold");
		SafeArrayPutString(*saData.parray, 1, 6, L"Time\nDelta");
		SafeArrayPutString(*saData.parray, 1, 7, L"Thrust Element\n(N)");

		//
		// Labels for reference data
		SafeArrayPutString(*saDataReferences.parray, 1, 1, L"Time Delta");				// AA
		SafeArrayPutString(*saDataReferences.parray, 1, 2, L"Thrust\n(lbs.)");			// AB
		SafeArrayPutString(*saDataReferences.parray, 1, 3, L"Thrust\n(N)");				// AC
		SafeArrayPutString(*saDataReferences.parray, 1, 4, L"Pressure\n(PSI)");			// AD
		SafeArrayPutString(*saDataReferences.parray, 1, 5, L"Thrust Element\n(N)");		// AE

		LONG lIndexBurnStart = lIndex;
		LONG lIndexBurnEnd = lIndex;
		for(auto data : vReadings)
		{
			double dTime = std::get<eTime>(data);
			double dThrust = std::get<eThrust>(data);
			double dPressure = std::get<ePressure>(data);
			double dThrustLBS = bMetric ? dThrust / g_dNewtonsPerPound : dThrust;
			double dThrustN = bMetric ? dThrust : dThrust * g_dNewtonsPerPound;
			double dTimeDelta = std::get<eTimeDelta>(data);
			double dThrustElementN = std::get<eThrustElementN>(data);

			SafeArrayPutDouble(*saData.parray, lIndex, 1, dTime);			// A
			SafeArrayPutDouble(*saData.parray, lIndex, 2, dThrustLBS);		// B
			SafeArrayPutDouble(*saData.parray, lIndex, 3, dThrustN);		// C
			SafeArrayPutDouble(*saData.parray, lIndex, 4, dPressure);		// D
			std::wstring wstrPmaxCalcCell(L"=IF($D$");
			wstrPmaxCalcCell += std::to_wstring(lIndex);
			wstrPmaxCalcCell += L"<$Q$35,0,$Q$35";
			SafeArrayPutString(*saData.parray, lIndex, 5, wstrPmaxCalcCell);// E (pMax threshold)
			SafeArrayPutDouble(*saData.parray, lIndex, 6, dTimeDelta);		// F
			SafeArrayPutDouble(*saData.parray, lIndex, 7, dThrustElementN);	// G

			//
			// Generate a set of data (to be hidden) which is blank when the real value is below the threshold
			auto fnGenThresholdIFStatement = ([lIndex](std::wstring &wstrStatement, wchar_t wcCol)
			{
				wstrStatement = L"=IF($E$";
				wstrStatement += std::to_wstring(lIndex);
				wstrStatement += L",$";
				wstrStatement += wcCol;
				wstrStatement += L"$";
				wstrStatement += std::to_wstring(lIndex);
				wstrStatement += L",\"\")";
			});

			//
			// Now the references to included data  =IF($E$n,$F$n,"")
			std::wstring wstrRef;
			fnGenThresholdIFStatement(wstrRef, L'F');
			SafeArrayPutString(*saDataReferences.parray, lIndex, 1, wstrRef);// AA (Time Delta)
			fnGenThresholdIFStatement(wstrRef, L'B');
			SafeArrayPutString(*saDataReferences.parray, lIndex, 2, wstrRef);// AB (Thrust LBS.)
			fnGenThresholdIFStatement(wstrRef, L'C');
			SafeArrayPutString(*saDataReferences.parray, lIndex, 3, wstrRef);// AC (Thrust N.)
			fnGenThresholdIFStatement(wstrRef, L'D');
			SafeArrayPutString(*saDataReferences.parray, lIndex, 4, wstrRef);// AD (Pressure)
			fnGenThresholdIFStatement(wstrRef, L'G');
			SafeArrayPutString(*saDataReferences.parray, lIndex, 5, wstrRef);// AE (Thrust Element)

			++lIndex;
		}

		//
		// Insert the data into A1-Gx
		std::wstring wstrDataRange(L"A1:G");
		wstrDataRange += std::to_wstring(lIndex);

		Excel::RangePtr pXLRange = pXLSheet->Range[_variant_t(wstrDataRange.c_str())];
		_ASSERT(pXLRange);
		if(nullptr == pXLRange)
		{
			return false;
		}
		pXLRange->Value2 = saData;

		//
		// Insert the Reference Data in AA - EE
		std::wstring wstrDataRangeRef(L"AA1:AE");
		wstrDataRangeRef += std::to_wstring(lIndex);

		Excel::RangePtr pXLRangeRef = pXLSheet->Range[_variant_t(wstrDataRangeRef.c_str())];
		_ASSERT(pXLRangeRef);
		if(nullptr == pXLRangeRef)
		{
			return false;
		}
		pXLRangeRef->Value2 = saDataReferences;
	}

	//
	// Create a chart showing pressure and thrust curves
	// With a secondary axis for thrust so that curves are near in scale
	// e.g. "ASBLUE3_16!$A$1:$B$457,ASBLUE3_16!$D$1:$D$457"
	std::wstring wstrChartRange(wstrTabName + L"!$A$1:$B$");
	wstrChartRange += std::to_wstring(lIndex);
	wstrChartRange += L",";
	wstrChartRange += wstrTabName;
	wstrChartRange += L"!$D$1:$D$";
	wstrChartRange += std::to_wstring(lIndex);
	_variant_t varChartRangeString(wstrChartRange.c_str());

	//
	// Create the range...
	Excel::RangePtr pChartRange = pXLSheet->Range[varChartRangeString];
	_ASSERT(pChartRange);
	if(nullptr == pChartRange)
	{
		return false;
	}

	//
	// Create the chart -- odd, but this the method for getting it embedded on
	// the same page as the data...
	auto pShapes = pXLSheet->GetShapes();
	_ASSERT(pShapes);
	if(nullptr == pShapes)
	{
		return false;
	}

	// auto pChartShape = pShapes->AddChart2(240, Excel::xlXYScatter); <- Requires Excel 2013
	auto pChartShape = pShapes->AddChart(Excel::xlXYScatter);
	_ASSERT(pChartShape);
	if(nullptr == pChartShape)
	{
		return false;
	}
	pChartShape->Select();

	//
	// With the newly created Chart Shape we can access it's chart with Workbook::GetActiveChart()
	// And set the source data to the range created previously...
	auto pChart = pXLBook->GetActiveChart();
	_ASSERT(pChart);
	if(nullptr == pChart)
	{
		return false;
	}
	pChart->SetSourceData(pChartRange);

	//
	// Set the Chart Title
	pChart->PutHasTitle(LOCALE_USER_DEFAULT, VARIANT_TRUE);
	auto pChartTitle = pChart->ChartTitle;
	_ASSERT(pChartTitle);
	if(nullptr != pChartTitle)
	{
		pChartTitle->Select();
		pChartTitle->Text = _bstr_t(L"Pressure and Thrust");
	}

	//
	// Setup as a custom chart with stacked lines and a secondary axis
	Excel::SeriesPtr pFSC1 = pChart->SeriesCollection(variant_t(1));
	Excel::SeriesPtr pFSC2 = pChart->SeriesCollection(variant_t(2));
	_ASSERT(pFSC1);
	_ASSERT(pFSC2);
	if(nullptr == pFSC1 || nullptr == pFSC2)
	{
		return false;
	}
	pFSC1->ChartType = Excel::xlLineStacked;
	pFSC1->AxisGroup = Excel::xlSecondary;
	pFSC2->ChartType = Excel::xlLineStacked;
	pFSC2->AxisGroup = Excel::xlPrimary;

	//
	// Make the axis labels appear (left for PSI, bottom for time, and right for thrust)
	pChart->SetElement(Office::msoElementPrimaryCategoryAxisTitleAdjacentToAxis);
	pChart->SetElement(Office::msoElementPrimaryValueAxisTitleRotated);
	pChart->SetElement(Office::msoElementSecondaryValueAxisTitleAdjacentToAxis);

	//
	// Set the bottom (Primary - Category) axis title to "time":
	Excel::AxisPtr pPrimaryAxes = pChart->Axes(Excel::xlCategory, Excel::xlPrimary);
	_ASSERT(pPrimaryAxes);
	if(nullptr == pPrimaryAxes)
	{
		return false;
	}

	Excel::AxisTitlePtr pPrimaryAxisTitle = pPrimaryAxes->AxisTitle;
	_ASSERT(pPrimaryAxisTitle);
	if(nullptr == pPrimaryAxisTitle)
	{
		return false;
	}
	pPrimaryAxisTitle->Select();
	_bstr_t bstrTime(L"Time");
	pPrimaryAxisTitle->Text = bstrTime;

	//
	// Set the left (Primary - Value) axis title to "PSI":
	_bstr_t bstrPSI(L"PSI");
	Excel::AxisPtr pPrimaryVAxes = pChart->Axes(Excel::xlValue, Excel::xlPrimary);
	_ASSERT(pPrimaryVAxes);
	if(nullptr == pPrimaryVAxes)
	{
		return false;
	}
	Excel::AxisTitlePtr pPrimaryVAxisTitle = pPrimaryVAxes->AxisTitle;
	_ASSERT(pPrimaryVAxisTitle);
	if(nullptr == pPrimaryVAxisTitle)
	{
		return false;
	}
	pPrimaryVAxisTitle->Select();
	pPrimaryVAxisTitle->Text = bstrPSI;

	//
	// Set the right (Secondary - Value) axis title to "Pounds of Thrust"
	Excel::AxisPtr pSecondaryAxes = pChart->Axes(Excel::xlValue, Excel::xlSecondary);
	_ASSERT(pSecondaryAxes);
	if(nullptr == pSecondaryAxes)
	{
		return false;
	}
	Excel::AxisTitlePtr pSecondaryAxisTitle = pSecondaryAxes->AxisTitle;
	_ASSERT(pSecondaryAxisTitle);
	if(nullptr == pSecondaryAxisTitle)
	{
		return false;
	}
	pSecondaryAxisTitle->Text = _bstr_t(L"Pounds of Thrust");

	//
	// Move chart into place and give a good size:
	Excel::ChartAreaPtr pChartArea = pChart->GetChartArea(0);
	_ASSERT(pChartArea);
	if(nullptr == pChartArea)
	{
		return false;
	}
	pChartArea->PutLeft(365.0f);
	pChartArea->PutTop(52.0f);
	pChartArea->PutHeight(300.0f);
	pChartArea->PutWidth(500.0f);

	//
	// Set Gridlines
	pChart->SetElement(Office::msoElementPrimaryCategoryGridLinesNone);
	pChart->SetElement(Office::msoElementPrimaryValueGridLinesNone);
	pChart->SetElement(Office::msoElementPrimaryValueGridLinesMajor);
	pChart->SetElement(Office::msoElementPrimaryValueGridLinesNone);
	pChart->SetElement(Office::msoElementPrimaryValueGridLinesMinor);

	//
	// Add the 10% Pmax line:
	Excel::SeriesCollectionPtr pSeriesCol = pChart->SeriesCollection();
	_ASSERT(pSeriesCol);
	if(nullptr == pSeriesCol)
	{
		return false;
	}
	Excel::SeriesPtr pPmaxSeries = pSeriesCol->NewSeries();
	_ASSERT(pPmaxSeries);
	if(nullptr == pPmaxSeries)
	{
		return false;
	}
	pPmaxSeries->Name = _bstr_t(L"Action Time");

	std::wstring wstrPmaxValue(L"=");
	wstrPmaxValue += wstrTabName;
	std::wstring wstrPmaxXValue(wstrPmaxValue);
	wstrPmaxValue += L"!$E$2:$E$";
	wstrPmaxXValue += L"!$A$2:$A$";
	wstrPmaxValue += std::to_wstring(lIndex);
	wstrPmaxXValue += std::to_wstring(lIndex);
	pPmaxSeries->Values = _bstr_t(wstrPmaxValue.c_str());
	pPmaxSeries->XValues = _bstr_t(wstrPmaxXValue.c_str());
	pPmaxSeries->ChartType = Excel::xlLine;
	pPmaxSeries->AxisGroup = Excel::xlPrimary;

	std::wstring wstrPressureRange(L"(D2:D");
	wstrPressureRange += std::to_wstring(lIndex);
	wstrPressureRange += L")";

	//
	// Display other file data...
	const std::string strGrains(tree.get("Document.MotorData.Grains", "0"));
	const std::string strCaseDiameter(tree.get("Document.MotorData.CaseDiameter", "0 mm"));
	const std::string strNozzleThroat(tree.get("Document.MotorData.NozzleThroatDiameter", "0"));
	std::string strInfo(strGrains + " grain " + strCaseDiameter + " " + strPropellant + " with " + strNozzleThroat + "\" throat 'nozzle'");
	std::wstring wstrInfo(converter.from_bytes(strInfo.c_str()));
	std::wstring wstrInfoRange(L"K1");

	Excel::RangePtr pInfoRange = pXLSheet->Range[_variant_t(wstrInfoRange.c_str())];
	_ASSERT(pInfoRange);
	if(nullptr == pInfoRange)
	{
		return false;
	}
	pInfoRange->Value2 = _bstr_t(wstrInfo.c_str());

	//
	// Add calculated data: ???  5% line and other calculated data...
	// <doc name, display name, display location row letter, display location cell number, metric conversion>
	std::vector<std::tuple<std::string, std::wstring, wchar_t, int, bool> > vFileItemDisplay;
	vFileItemDisplay.push_back(std::make_tuple("Document.MotorData.MaxThrust", L"Maximum Thrust:", L'K', 25, true));
	vFileItemDisplay.push_back(std::make_tuple("Document.MotorData.AvgThrust", L"Average Thrust:", L'K', 26, true));
	vFileItemDisplay.push_back(std::make_tuple("Document.MotorData.MaxPressure", L"Maximum Pressure:", L'K', 28, false));
	vFileItemDisplay.push_back(std::make_tuple("Document.MotorData.AvgPressure", L"Average Pressure:", L'K', 29, false));
	vFileItemDisplay.push_back(std::make_tuple("Document.MotorData.BurnTime", L"Burn Time:", L'K', 31, false));
	vFileItemDisplay.push_back(std::make_tuple("Document.MotorData.Impuse", L"Impulse:", L'K', 32, false));	// NOTE: Spelling mistake in the TC Logger data!

	PlaceStringInCell(pXLSheet, L'L', 24, L"TC Logger File Data:");
	for(auto entry : vFileItemDisplay)
	{
		std::wstring wstrDisplay(std::get<1>(entry));

		double dData = tree.get(std::get<0>(entry).c_str(), 0.0f);
		bool bConversion = std::get<4>(entry);

		const wchar_t wcColumn = std::get<2>(entry);
		const wchar_t wcDataColumn = wcColumn + 2;
		int row = std::get<3>(entry);

		PlaceStringInCell(pXLSheet, wcColumn, row, wstrDisplay);
		if(bConversion)
		{
			double dThrustLBS = bMetric ? dData / g_dNewtonsPerPound : dData;
			double dThrustN = bMetric ? dData : dData * g_dNewtonsPerPound;

			PlaceStringInCell(pXLSheet, wcDataColumn, row, std::to_wstring(dThrustLBS));
			PlaceStringInCell(pXLSheet, wcDataColumn + 1, row, L"LBS.");
			PlaceStringInCell(pXLSheet, wcDataColumn + 2, row, std::to_wstring(dThrustN));
			PlaceStringInCell(pXLSheet, wcDataColumn + 3, row, L"N");
		}
		else
		{
			PlaceStringInCell(pXLSheet, wcDataColumn, row, std::to_wstring(dData));
		}
	}

	PlaceStringInCell(pXLSheet, L'I', 34, L"VARIABLES:");
	//
	// Percent of Pmax used for calculations
	PlaceStringInCell(pXLSheet, L'I', 35, L"% Pmax for Burn threshold:");
	PlaceStringInCell(pXLSheet, L'L', 35, mapProperties[L"PmaxThreshold"]);		// 10% by default gives Action Time according to Sutton
	PlaceStringInCell(pXLSheet, L'M', 35, L"Pmax:");
	std::wstring wstrPmaxCalc(L"=MAX");
	wstrPmaxCalc += wstrPressureRange;
	PlaceStringInCell(pXLSheet, L'N', 35, wstrPmaxCalc);
	PlaceStringInCell(pXLSheet, L'O', 35, L"Pmax Threshold:");
	PlaceStringInCell(pXLSheet, L'Q', 35, L"=($N$35*($L$35/100))");		// Q35

	//
	// Grain Weight (g):			[178]
	PlaceStringInCell(pXLSheet, L'I', 36, L"Grain Weight (g):");
	PlaceStringInCell(pXLSheet, L'L', 36, mapProperties[L"GrainWeight"]);

	// Liner Weight g/in:			[ 1.9197342 ] or [ 3.030928 ]
	PlaceStringInCell(pXLSheet, L'I', 37, L"Casting Tube Weight (g/in.):");
	PlaceStringInCell(pXLSheet, L'L', 37, mapProperties[L"CastingTubeWeight"]);
	// PlaceStringInCell(pXLSheet, L'M', 37, L"White 54mm casting tube: 1.9197342 g/in.  Tru-Core Waxy 54mm: 3.030928 g/in");

	// Grain Length:				[ 3.0625 ]
	PlaceStringInCell(pXLSheet, L'I', 38, L"Grain Length (in.):");
	PlaceStringInCell(pXLSheet, L'L', 38, mapProperties[L"GrainLength"]);

	// Grain Diameter (in.):		[ 1.75 ]
	PlaceStringInCell(pXLSheet, L'I', 39, L"Grain Diameter (in.):");
	PlaceStringInCell(pXLSheet, L'L', 39, mapProperties[L"GrainDiameter"]);

	// Grain Core (in.):			[ 0.625 ]
	PlaceStringInCell(pXLSheet, L'I', 40, L"Grain Core (in.):");
	PlaceStringInCell(pXLSheet, L'L', 40, mapProperties[L"GrainCore"]);

	//
	// Our calculations (using Excel formulas so input can be changed)...
	PlaceStringInCell(pXLSheet, L'I', 42, L"Calculations based on Pmax Threshold:");

	// Density:
	PlaceStringInCell(pXLSheet, L'I', 43, L"Density:");
	PlaceStringInCell(pXLSheet, L'L', 43, L"=((L36-(L38*L37))/453.59237)/(((((L39/2)^2)-(L40/2)^2))*PI()*L38)");

	// Web
	PlaceStringInCell(pXLSheet, L'I', 44, L"Web:");
	PlaceStringInCell(pXLSheet, L'L', 44, L"=(L39-L40)/2");

	// Burn Time (s)
	PlaceStringInCell(pXLSheet, L'I', 45, L"Burn Time (s):");
	std::wstring wstrBurnTime(L"=SUM(AA1:AA");
	wstrBurnTime += std::to_wstring(lIndex);
	PlaceStringInCell(pXLSheet, L'L', 45, wstrBurnTime);
	PlaceStringInCell(pXLSheet, L'M', 45, std::wstring(L"10% = " + std::to_wstring(dBurnTime)));

	// Max Pressure
	PlaceStringInCell(pXLSheet, L'I', 46, L"Max Pressure (PSI):");
	std::wstring wstrBtPressureRange(L"(AD1:AD");
	wstrBtPressureRange += std::to_wstring(lIndex);
	wstrBtPressureRange += L")";

	std::wstring wstrMaxPressure(L"=MAX");
	wstrMaxPressure += wstrBtPressureRange;
	PlaceStringInCell(pXLSheet, L'L', 46, wstrMaxPressure);

	// Average Pressure
	PlaceStringInCell(pXLSheet, L'I', 47, L"Average Pressure (PSI):");
	std::wstring wstrAvgPressure(L"=AVERAGE");
	wstrAvgPressure += wstrBtPressureRange;
	PlaceStringInCell(pXLSheet, L'L', 47, wstrAvgPressure);

	std::wstring wstrThrustRange(L"(AB1:AB");
	std::wstring wstrThrustRangeN(L"(AC1:AC");
	std::wstring wstrThrustElementRangeN(L"(AE1:AE");
	wstrThrustRange += std::to_wstring(lIndex);
	wstrThrustRangeN += std::to_wstring(lIndex);
	wstrThrustElementRangeN += std::to_wstring(lIndex);
	wstrThrustRange += L")";
	wstrThrustRangeN += L")";
	wstrThrustElementRangeN += L")";

	// Max Thrust (LBS)
	PlaceStringInCell(pXLSheet, L'I', 48, L"Max Thrust (LBS):");
	std::wstring wstrMaxThrust(L"=MAX");
	wstrMaxThrust += wstrThrustRange;
	PlaceStringInCell(pXLSheet, L'L', 48, wstrMaxThrust);

	// Max Thrust (N)
	PlaceStringInCell(pXLSheet, L'I', 49, L"Max Thrust (N):");
	std::wstring wstrMaxThrustN(L"=MAX");
	wstrMaxThrustN += wstrThrustRangeN;
	PlaceStringInCell(pXLSheet, L'L', 49, wstrMaxThrustN);

	// Average Thrust (LBS)
	PlaceStringInCell(pXLSheet, L'I', 50, L"Average Thrust (LBS):");
	std::wstring wstrAvgThrust(L"=AVERAGE");
	wstrAvgThrust += wstrThrustRange;
	PlaceStringInCell(pXLSheet, L'L', 50, wstrAvgThrust);

	// Average Thrust (N)
	PlaceStringInCell(pXLSheet, L'I', 51, L"Average Thrust (N):");
	std::wstring wstrAvgThrustN(L"=AVERAGE");
	wstrAvgThrustN += wstrThrustRangeN;
	PlaceStringInCell(pXLSheet, L'L', 51, wstrAvgThrustN);

	// Total Thrust:
	PlaceStringInCell(pXLSheet, L'I', 52, L"Total Thrust (N):");
	std::wstring wstrTotalThrustN(L"=SUM");
	wstrTotalThrustN += wstrThrustElementRangeN;
	PlaceStringInCell(pXLSheet, L'L', 52, wstrTotalThrustN);

	// Grain Weight minus Liner:
	PlaceStringInCell(pXLSheet, L'I', 53, L"Propellant Weight (g):");
	PlaceStringInCell(pXLSheet, L'L', 53, L"=(L36-(L38*L37))");

	// ISP:
	PlaceStringInCell(pXLSheet, L'I', 55, L"ISP:");
	PlaceStringInCell(pXLSheet, L'L', 55, L"=L52/(9.8*(L53/1000))");

	// Burn Rate:
	PlaceStringInCell(pXLSheet, L'I', 56, L"Burn Rate:");
	PlaceStringInCell(pXLSheet, L'L', 56, L"=(L44/L45)");

	//
	// Make the INPUT area GREEN:
	Excel::RangePtr pVariableRange = pXLSheet->Range[_variant_t(L"I35:L40")];
	_ASSERT(pVariableRange);
	if(nullptr == pVariableRange)
	{
		return false;
	}
	Excel::InteriorPtr pIntVarRange = pVariableRange->Interior;
	_ASSERT(pIntVarRange);
	if(nullptr == pIntVarRange)
	{
		return false;
	}
	pIntVarRange->Pattern = xlSolid;
	pIntVarRange->PatternColorIndex = Excel::xlAutomatic;
	pIntVarRange->Color = 5287936;
	pIntVarRange->TintAndShade = 0;
	pIntVarRange->PatternTintAndShade = 0;

	//
	// Output Area
	Excel::RangePtr pCalcRange = pXLSheet->Range[_variant_t(L"I43:L56")];
	_ASSERT(pCalcRange);
	if(nullptr == pCalcRange)
	{
		return false;
	}
	Excel::InteriorPtr pIntCalcRange = pCalcRange->Interior;
	_ASSERT(pIntCalcRange);
	if(nullptr == pIntCalcRange)
	{
		return false;
	}
	pIntCalcRange->PatternColorIndex = Excel::xlAutomatic;
	pIntCalcRange->Color = 7373816;
	pIntCalcRange->TintAndShade = 0;
	pIntCalcRange->PatternTintAndShade = 0;

	return true;
}

void ProduceAandNTab(Excel::_WorkbookPtr &pXLBook,
	const std::vector<std::wstring> &vTabs)
{
	// Get the active Worksheet and set its name.
	Excel::_WorksheetPtr pXLSheet = pXLBook->ActiveSheet;
	_ASSERT(pXLSheet);
	if(nullptr == pXLSheet)
	{
		return;
	}
	pXLSheet->Name = _bstr_t(L"a & n");

	//
	// Create four columns of data: Name, Pc, Br, and ISP
	// Construct a safearray of the data
	LONG lIndex = 2;		// Note, 1 based not 0 - will contain # of rows...
	VARIANT saData;
	saData.vt = VT_ARRAY | VT_VARIANT;
	{
		SAFEARRAYBOUND sab[2];
		sab[0].lLbound = 1;
		sab[0].cElements = vTabs.size() + 2;
		sab[1].lLbound = 1;
		sab[1].cElements = 4;
		saData.parray = SafeArrayCreate(VT_VARIANT, 2, sab);
		if(saData.parray == nullptr)
		{
			MessageBoxW(HWND_TOP, L"Unable to create safearray for passing data to Excel.", L"Memory error...", MB_TOPMOST | MB_ICONERROR);
			return;
		}

		//
		// Clean-up safe array when done...
		auto cleanArray = std::experimental::make_scope_exit([&saData]() -> void
		{
			VariantClear(&saData);
		});

		//
		// Labels:
		SafeArrayPutString(*saData.parray, 1, 1, L"Motor");
		SafeArrayPutString(*saData.parray, 1, 2, L"Pressure");
		SafeArrayPutString(*saData.parray, 1, 3, L"Burn Rate");
		SafeArrayPutString(*saData.parray, 1, 4, L"ISP");

		for(auto entry : vTabs)
		{
			std::wstring wstrLink(L"=");
			wstrLink += entry;
			wstrLink += L"!$L$";
			const std::wstring wstrPc(wstrLink + L"47");
			const std::wstring wstrBr(wstrLink + L"56");
			const std::wstring wstrISP(wstrLink + L"55");

			SafeArrayPutString(*saData.parray, lIndex, 1, entry);		// A Name
			SafeArrayPutString(*saData.parray, lIndex, 2, wstrPc);		// B Pc
			SafeArrayPutString(*saData.parray, lIndex, 3, wstrBr);		// C Br
			SafeArrayPutString(*saData.parray, lIndex, 4, wstrISP);		// D Isp
			++lIndex;
		}

		//
		// Insert the data into A1-Dx
		std::wstring wstrDataRange(L"A1:D");
		wstrDataRange += std::to_wstring(lIndex);

		Excel::RangePtr pXLRange = pXLSheet->Range[_variant_t(wstrDataRange.c_str())];
		pXLRange->Value2 = saData;
	}

	// Create a chart of the data
	std::wstring wstrChartRange(L"='a & n'!$B$1:$C$");
	wstrChartRange += std::to_wstring(lIndex - 1);
	_variant_t varChartRangeString(wstrChartRange.c_str());

	//
	// Create the range...
	Excel::RangePtr pChartRange = pXLSheet->Range[varChartRangeString];
	_ASSERT(pChartRange);
	if(nullptr == pChartRange)
	{
		return;
	}
	pChartRange->Select();

	//
	// Create the chart -- odd, but this the method for getting it embedded on
	// the same page as the data...
	auto pShapes = pXLSheet->GetShapes();
	_ASSERT(pShapes);
	if(nullptr == pShapes)
	{
		return;
	}

	// auto pChartShape = pShapes->AddChart2(240, Excel::xlXYScatterLines);
	auto pChartShape = pShapes->AddChart(Excel::xlXYScatterLines);
	_ASSERT(pChartShape);
	if(nullptr == pChartShape)
	{
		return;
	}
	pChartShape->Select();

	//
	// With the newly created Chart Shape we can access it's chart with Workbook::GetActiveChart()
	// And set the source data to the range created previously...
	auto pChart = pXLBook->GetActiveChart();
	_ASSERT(pChart);
	if(nullptr == pChart)
	{
		return;
	}
	pChart->PutHasTitle(LOCALE_USER_DEFAULT, VARIANT_TRUE);
	pChart->SetSourceData(pChartRange);

	//
	// Set the Chart Title
	auto pChartTitle = pChart->ChartTitle;
	_ASSERT(pChartTitle);
	if(nullptr != pChartTitle)
	{
		pChartTitle->Select();
		pChartTitle->Text = _bstr_t(L"a & n");
	}

	// Add Trendline:
	Excel::SeriesPtr pFSC1 = pChart->SeriesCollection(variant_t(1));
	_ASSERT(pFSC1);
	if(nullptr == pFSC1)
	{
		return;
	}
	pFSC1->Select();

	Excel::TrendlinesPtr pTrendLines = pFSC1->Trendlines();
	_ASSERT(pTrendLines);
	if(nullptr == pTrendLines)
	{
		return;
	}
	pTrendLines->Add(Excel::xlPower);

	Excel::TrendlinePtr pTrendLine = pTrendLines->Item(_variant_t(1));
	_ASSERT(pTrendLine);
	if(nullptr == pTrendLine)
	{
		return;
	}
	pTrendLine->Select();
	pTrendLine->Type = Excel::xlPower;
	pTrendLine->Forward = 200;
	pTrendLine->Backward = 200;
	pTrendLine->DisplayEquation = VARIANT_TRUE;
	pTrendLine->DisplayRSquared = VARIANT_TRUE;

	// Enlarge Text
	Excel::DataLabelPtr pDataLabel = pTrendLine->DataLabel;
	_ASSERT(pDataLabel);
	if(nullptr == pDataLabel)
	{
		return;
	}
	pDataLabel->Select();

	Excel::FontPtr pFont = pDataLabel->Font;
	_ASSERT(pFont);
	if(nullptr == pFont)
	{
		return;
	}
	pFont->Size = 16;

	// Move to top right...
	pDataLabel->Top = 40;
	pDataLabel->Left = 40;

	//
	// Move chart into place and give a good size:
	Excel::ChartAreaPtr pChartArea = pChart->GetChartArea(0);
	_ASSERT(pChartArea);
	if(nullptr == pChartArea)
	{
		return;
	}
	pChartArea->PutLeft(240.0f);
	pChartArea->PutTop(0.0f);

	//
	// Now give an Average ISP:
	PlaceStringInCell(pXLSheet, L'H', 17, L"Average Delivered ISP:");
	std::wstring wstrAvgISP(L"=AVERAGE(D2:D");
	wstrAvgISP += std::to_wstring(lIndex - 1);
	wstrAvgISP += L")";
	PlaceStringInCell(pXLSheet, L'K', 17, wstrAvgISP);
	//
	// Place a & n numbers on the sheet:
	// a: =EXP(INDEX(LINEST(LN(C2:C4),LN(B2:B4),,)*1,2))
	// n: =INDEX(LINEST(LN(C2:C4), LN(B2:B4),,),1) 
	PlaceStringInCell(pXLSheet, L'H', 18, L"Burn Rate Coefficient (a):");
	PlaceStringInCell(pXLSheet, L'H', 19, L"Burn Rate Exponent (n):");

	std::wstring wstrCoefficient(L"=EXP(INDEX(LINEST(LN(C2:C");
	wstrCoefficient += std::to_wstring(vTabs.size() + 1);
	wstrCoefficient += L"), LN(B2:B";
	wstrCoefficient += std::to_wstring(vTabs.size() + 1);
	wstrCoefficient += L"), , ) * 1, 2))";
	PlaceStringInCell(pXLSheet, L'K', 18, wstrCoefficient);

	std::wstring wstrExponent(L"=INDEX(LINEST(LN(C2:C");
	wstrExponent += std::to_wstring(vTabs.size() + 1);
	wstrExponent += L"), LN(B2:B";
	wstrExponent += std::to_wstring(vTabs.size() + 1);
	wstrExponent += L"), , ), 1)";
	PlaceStringInCell(pXLSheet, L'K', 19, wstrExponent);

	//
	// Draw attention to the results:
	Excel::RangePtr pResultsRange = pXLSheet->Range[_bstr_t(L"H17:K19")];
	_ASSERT(pResultsRange);
	if(nullptr == pResultsRange)
	{
		return;
	}
	pResultsRange->Style = _bstr_t(L"Output");
}

void SortInputFiles(std::vector<std::wstring> &vwstrInputFiles)
{
	std::map<double, std::wstring> mapSortedFiles;
	for(auto wstrFile : vwstrInputFiles)
	{
		boost::property_tree::ptree tree;
		std::filebuf file;
		if(false == file.open(wstrFile.c_str(), std::ios::in))
		{
			continue;
		}
		std::istream inputStream(&file);
		boost::property_tree::xml_parser::read_xml(inputStream, tree);
		double dExitDiameter = tree.get("Document.MotorData.ExitDiameter", 0.0f);
		while(mapSortedFiles.count(dExitDiameter) != 0)
		{
			dExitDiameter += 0.0000001f;		// Hack to order according to input order if multiples of same diameter
		}
		mapSortedFiles[dExitDiameter] = wstrFile;
	}
	vwstrInputFiles.clear();
	std::for_each(mapSortedFiles.rbegin(), mapSortedFiles.rend(),
		[&vwstrInputFiles](decltype(mapSortedFiles)::value_type &entry) ->void
	{
		vwstrInputFiles.push_back(entry.second);
	});
}

void CreateExcelSpreadsheet(std::map<std::wstring, std::wstring> &mapProperties)
{
	CoInitialize(nullptr);
	auto aShutdown = std::experimental::make_scope_exit([]() ->void
	{
		CoUninitialize();
	});

	//
	// Make sure Excel is present...
	CLSID clsid;
	HRESULT hr = CLSIDFromProgID(L"Excel.Application", &clsid);
	if(FAILED(hr))
	{
		MessageBoxW(HWND_TOP, L"Unable to find Excel.Application's CLSID.", L"Excel unavailable...", MB_TOPMOST | MB_ICONERROR);
		return;
	}

	Excel::_ApplicationPtr pXLApp;
	hr = pXLApp.CreateInstance(__uuidof(Excel::Application));
	if(FAILED(hr))
	{
		MessageBoxW(HWND_TOP, L"Unable to connect with Excel.", L"Excel unavailable...", MB_TOPMOST | MB_ICONERROR);
		return;
	}

	std::vector<std::wstring> vwstrInputFiles;
	if(false == ObtainInputFiles(vwstrInputFiles))
	{
		return;
	}

	//
	// Sort input files by nozzle diameter
	SortInputFiles(vwstrInputFiles);

	// Make Excel invisible. (i.e. Application.Visible = 0)
	pXLApp->Visible[0] = VARIANT_TRUE;

	// Create a new Excel Workbook. (i.e. Application.Workbooks.Add)
	Excel::WorkbooksPtr pXLBooks = pXLApp->Workbooks;
	_ASSERT(pXLBooks);
	if(nullptr == pXLBooks)
	{
		return;
	}

	Excel::_WorkbookPtr pXLBook = pXLBooks->Add();
	_ASSERT(pXLBook);
	if(nullptr == pXLBook)
	{
		return;
	}
	pXLApp->ActiveWindow->PutWindowState(Excel::xlMaximized);

	std::vector<std::wstring> vTabnames;
	bool bFirst = true;
	for(auto wstrFile : vwstrInputFiles)
	{
		std::array<wchar_t, 1024> wzFilename;
		_wsplitpath_s(wstrFile.c_str(), nullptr, 0, nullptr, 0, &wzFilename[0], wzFilename.size(), nullptr, 0);

		boost::property_tree::ptree tree;
		std::filebuf file;
		if(false == file.open(wstrFile.c_str(), std::ios::in))
		{
			continue;
		}
		std::istream inputStream(&file);
		boost::property_tree::xml_parser::read_xml(inputStream, tree);
		std::string strPropellantName(tree.get("Document.MotorData.Propellant", "-"));
		printf("NAME: %hs\n", strPropellantName.c_str());

		const std::string strMetric(tree.get("Document.MotorData.Metric", "False"));
		bool bMetric = (strMetric == "True");

		double dSelectedAvgStartTime = tree.get("Document.MotorData.SelectedAvgStartTime", 0.0f);
		double dSelectedAvgEndTime = tree.get("Document.MotorData.SelectedAvgEndTime", 0.0f);

		auto motorData = tree.get_child("Document.MotorData");
		std::vector<std::tuple<double, double, double, double, double> > vReadings;
		double dLastTime = dSelectedAvgStartTime;

		for(auto data : motorData)
		{
			if(data.first != "Item")
			{
				continue;
			}
			double dTime = data.second.get<double>("Time", -1.0f);
			double dThrust = data.second.get<double>("Thrust", -1.0f);
			double dPressure = data.second.get<double>("Pressure", -1.0f);
			//
			// Skip readings that were "selected" out in the TC Logger program
			// Useful for example to skip over an ignition spike...
			if(dSelectedAvgEndTime > 0.0f)
			{
				if(dTime < dSelectedAvgStartTime ||
					dTime > dSelectedAvgEndTime)
				{
					continue;
				}
			}
			double dTimeDelta = dTime - dLastTime;
			dLastTime = dTime;
			double dThrustElementN = bMetric ? (dTimeDelta * dThrust) : (dTimeDelta * (dThrust * g_dNewtonsPerPound));
			vReadings.push_back(std::make_tuple(dTime, dThrust, dPressure, dTimeDelta, dThrustElementN));
		}

		if(bFirst)
		{
			bFirst = false;
		}
		else
		{
			//
			// Create a new worksheet
			Excel::SheetsPtr pWorksheets = pXLBook->Worksheets;
			pWorksheets->Add();		// Becomes the active sheet...
		}

		//
		// Generate a workbook with data and a chart marking the 10% line, etc.
		// If successful and there are at least 3 files we will generate a sheet with a & n and average ISP
		std::wstring sanitizedFilename(&wzFilename[0]);
		SanitizeStringName(sanitizedFilename);
		if(true == ProduceExcelWorkbook(pXLBook, sanitizedFilename, strPropellantName, bMetric, tree, vReadings, mapProperties))
		{
			vTabnames.push_back(sanitizedFilename);
		}
	}

	//
	// If we have three or more sets of data we can generate an a & n:
	if(vTabnames.size() > 2)
	{
		Excel::SheetsPtr pWorksheets = pXLBook->Worksheets;
		pWorksheets->Add();		// Becomes the active sheet...
		ProduceAandNTab(pXLBook, vTabnames);
	}
	return;
}
