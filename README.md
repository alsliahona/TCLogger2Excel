# TCLogger2Excel
Converts data from files to Excel (2010 or newer), charting the data and adding burn-time line.  If 3 or more files are selected an a&amp;n tab is added where the Burn Rate Coefficent, Burn Rate Exponent, and average delivered ISP are calculated and displayed.

TCLogger files are produced by the TC Logger USB program that works with the TC Logger hardware sold by NASSA.  See http://www.tclogger.com/.

Dependencies:
* Boost 1.57.0 should be in a folder at the root level (same level as the TCLogger2Excel folder) named boost_1_57_0.  Only header files are used from Boost.  If a different version of Boost is used the project settings will need to be changed to point to the proper location.  Boost is available for download at http://www.boost.org/
* N4189 files are included.  N4189 is a proposed addition to the C++ Standard Library.  At the time of publication of this file N4189 had passed the C++ ISO Library Evolution Working Group committee and was being handled by the Library Working Group.  It is the work of the author (Andrew L. Sandoval) and Professor Peter Sommerlad (a member of the ISO C++ Committee).
* Visual Studio 2013 or better with C++ is required to compile TCLogger2Excel for Windows
* Microsoft Excel 2010 or newer is required

