/*******************************************************************************
 smartrename_pch.h : include file for libraries and windows system include files

 This code is free for any use by anyone under the terms of
 Creative Commons Attribution-NonCommercial 4.0 International Public License
 See the the file license.txt and https://creativecommons.org

 smartrename source code

 2023-04-20 Jerker Bäck created

================================================================================
 RcsID = $Id$ */

// clang-format off

#pragma once

#if defined(_MSC_VER) && !defined(__clang__)
#define MSDISABLE_WARNING(x)		__pragma(warning(disable:x))
#define MSDISABLE_WARNING_PUSH(x)	__pragma(warning(push));__pragma(warning(disable:x))
#define MSDISABLE_WARNING_POP		__pragma(warning(pop))
#define MSSUPPRESS_WARNING(x)		__pragma(warning(suppress:x))
#define MSADD_LIBRARY(x)			__pragma(comment(lib,x))
#define GCCDISABLE_WARNING(x)
#define GCCDISABLE_WARNING_PUSH(x)
#define GCCDISABLE_WARNING_POP
#elif defined(__GNUC__) || defined(__clang__)
#define MSDISABLE_WARNING(x)
#define MSDISABLE_WARNING_PUSH(x)
#define MSDISABLE_WARNING_POP
#define MSSUPPRESS_WARNING(x)
#define MSADD_LIBRARY(x)
#define GCCPRAGMA(x)				_Pragma(#x)
#define GCCDISABLE_WARNING(x)		GCCPRAGMA(GCC diagnostic ignored #x)
#define GCCDISABLE_WARNING_PUSH(x)  GCCPRAGMA(GCC diagnostic push) GCCPRAGMA(GCC diagnostic ignored #x)
#define GCCDISABLE_WARNING_POP      GCCPRAGMA(GCC diagnostic pop)
#else
#define MSDISABLE_WARNING(x)
#define MSDISABLE_WARNING_PUSH(x)
#define MSDISABLE_WARNING_POP
#define MSSUPPRESS_WARNING(x)
#define MSADD_LIBRARY(x)
#define GCCDISABLE_WARNING(x)
#define GCCDISABLE_WARNING_PUSH(x)
#define GCCDISABLE_WARNING_POP
#endif

GCCDISABLE_WARNING(-Wc++98-compat-pedantic)
GCCDISABLE_WARNING(-Wunused-command-line-argument)
GCCDISABLE_WARNING_PUSH(-Wreserved-id-macro)

#define STRICT
#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#define CRTDBG_MAP_ALLOC

#define  _WIN32_WINNT   0x0601	// minimum Windows 7
#include <winsdkver.h>
#include <sdkddkver.h>
#ifndef WINAPI_FAMILY
#define WINAPI_FAMILY WINAPI_FAMILY_DESKTOP_APP
#endif

MSDISABLE_WARNING_PUSH(4710)
#include <cstddef>
#include <cstdint>
#include <cstdlib>
#include <clocale>
#include <cstring>
#include <cwchar>
#include <cmath>
#include <limits>
#include <locale>
#include <string>
#include <stdexcept>
#include <chrono>
#include <algorithm>
#include <condition_variable>
#include <memory>
#include <mutex>
#include <thread>
#include <atomic>
#include <regex>
#include <filesystem>
#include <vector>
#include <map>
#include <set>
#include <unordered_set>
#include <tuple>
#include <functional>
#include <iostream>
#include <fstream>
#include <memory.h>
#include <malloc.h>
#include <crtdbg.h>
MSDISABLE_WARNING_POP

//MSDISABLE_WARNING_PUSH(4191 4365 4514 4626 4820 4986 5027 5039 5045 5204 5246 6011 6387 28204);
MSDISABLE_WARNING_PUSH(4191 4365 4514 4626 4820 5027 5039 5045 5204 5246);
// ATL
#define _ATL_ALL_USER_WARNINGS
#define _ATL_ALL_WARNINGS
#define _ATL_APARTMENT_THREADED
#define _ATL_NO_AUTOMATIC_NAMESPACE
#include <atlbase.h>
#include <atlcom.h>
#include <atlctl.h>
#include <atlwin.h>
MSDISABLE_WARNING_POP

MSDISABLE_WARNING_PUSH(4005 4514 4820 5039)
#define INITGUID
#include <guiddef.h>
#include <commctrl.h>
#include <pathcch.h>
#include <objbase.h>
#include <shellapi.h>
#include <shldisp.h>
#include <shlguid.h>
#include <shlobj.h>
#include <shlobj_core.h>
#include <shobjidl.h>
#include <shlwapi.h>
#include <unknwn.h>
#include <versionhelpers.h>
//#include <windowsx.h>
MSDISABLE_WARNING_POP

MSDISABLE_WARNING_PUSH(4191 4265 4355 4365 4514 4619 4625 4626 4668 4710 4820 5026 5027 5039 5045 5204 5220 5264 6001 26451)
#include <strsafe.h>
//#include <wil/result.h>
//#include <wil/resource.h>
//#include <wil/com.h>
//#include <wil/win32_helpers.h>
//#include <wil/cppwinrt.h>
//#include <wrl/wrappers/corewrappers.h>
MSDISABLE_WARNING_POP

using std::min;
using std::max;

#if defined(_DEBUG) && !defined(TESTEXE)
#define USE_UNITTEST    1
#define USE_MSTEST      1
#define USE_GTEST       0
#define USE_DOCTEST     0
//#include "CppUnitTest.h"
//#include "test/unittesting.h"
#else
#define USE_UNITTEST    0
#endif

#pragma comment(linker,"/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")

// global disable warnings
MSDISABLE_WARNING(4514 4571 4710 4711)
//MSDISABLE_WARNING(4005 4191 4365 4820 5039 5045 5204)
//GCCDISABLE_WARNING_POP

// clang-format on

