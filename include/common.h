/*******************************************************************************
 common.h : common custom definitions for the Windows API

 SPDX-License-Identifier: BSD-3-Clause
 This file is an independent addition to WTL, the Windows Template Library
 and is free for use by anyone under the terms of
 Creative Commons Attribution-NonCommercial 4.0 International Public License
 See https://creativecommons.org

 WTL source code

 2021-04-04 Jerker Bäck, created

================================================================================
 RcsID = $Id$ */

#pragma once

#include <type_traits> // _BITMASK_OPS

#ifdef BUILDDLL
#define IMPEXPORT extern "C" __declspec(dllexport)
#else
#define IMPEXPORT extern "C" __declspec(dllimport)
#endif

// give more meaning to disabled code
#define CONSIDERED_OBSOLETE 0
#define CONSIDERED_UNSUPPORTED 0
#define CONSIDERED_DISABLED 0
#define CONSIDERED_UNUSED 0

#ifdef _DEBUG
#define NODEFAULT _ASSERTE(0)
#else
#define NODEFAULT __assume(0)
#endif

// enum bit mask operations
#ifndef ENABLE_ENUM_BITMASKS
#define ENABLE_ENUM_BITMASKS(e) _BITMASK_OPS(, e) // <type_traits>
// constexpr bool _Bitmask_includes_any (_Left & _Elements) != _Bitmask{};
// constexpr bool _Bitmask_includes_all (_Left & _Elements) == _Elements;
#endif

#ifndef __cpp_lib_to_underlying
template <class Tenum>
constexpr std::underlying_type_t<Tenum> std::to_underlying(Tenum e) noexcept // (since C++23)
{
	return static_cast<std::underlying_type_t<Tenum>>(e);
}
#endif

// clang-format off
namespace
{

// -----------------------------------------------------------------------------
// exceptions => wil

//#define THROW_IF_FAILED(hr)                                     __R_FN(Throw_IfFailed)(__R_INFO(#hr) wil::verify_hresult(hr))
//#define THROW_IF_WIN32_BOOL_FALSE(win32BOOL)                    __R_FN(Throw_IfWin32BoolFalse)(__R_INFO(#win32BOOL) wil::verify_BOOL(win32BOOL))
//#define THROW_IF_WIN32_ERROR(win32err)                          __R_FN(Throw_IfWin32Error)(__R_INFO(#win32err) win32err)
//#define THROW_IF_NULL_ALLOC(ptr)                                __R_FN(Throw_IfNullAlloc)(__R_INFO(#ptr) ptr)
//#define THROW_HR_IF(hr, condition)                              __R_FN(Throw_HrIf)(__R_INFO(#condition) wil::verify_hresult(hr), wil::verify_bool(condition))
//#define THROW_HR_IF_NULL(hr, ptr)                               __R_FN(Throw_HrIfNull)(__R_INFO(#ptr) wil::verify_hresult(hr), ptr)
//#define THROW_LAST_ERROR_IF(condition)                          __R_FN(Throw_GetLastErrorIf)(__R_INFO(#condition) wil::verify_bool(condition))
//#define THROW_LAST_ERROR_IF_NULL(ptr)                           __R_FN(Throw_GetLastErrorIfNull)(__R_INFO(#ptr) ptr)
//#define THROW_IF_NTSTATUS_FAILED(status)                        __R_FN(Throw_IfNtStatusFailed)(__R_INFO(#status) status)

// Throw a std::system_error if the HRESULT indicates failure.
inline HRESULT ThrowIfFailed(HRESULT hr)
{
	THROW_IF_FAILED(hr);
	return hr;
}

// Windows API functions that return false on failure and set LastErrorValue
inline BOOL ThrowIfWinFalse(BOOL expr)
{
	BOOL result = expr;
	THROW_IF_WIN32_BOOL_FALSE(result);
	return result;
}

// Expression that return false on failure and don't set LastErrorValue
inline bool ThrowIfFalse(HRESULT hr, bool expr)
{
	bool result = expr;
	THROW_HR_IF(hr, result);
	return result;
}

// Windows API functions that return a nullptr pointer on failure and set LastErrorValue
template <typename Ttype>
Ttype ThrowIfNullptr(Ttype ptr)
{
	Ttype result = ptr;
	THROW_LAST_ERROR_IF_NULL(result);
	return result;
}

// Windows API functions that return 0 on failure and set LastErrorValue
template <typename Ttype>
Ttype ThrowIfZero(Ttype expr)
{
	if (expr == 0)
	{
		THROW_LAST_ERROR();
	}
	return expr;
}

// Windows API functions that return 0 on failure but don't set LastErrorValue
template <typename Ttype>
Ttype ThrowHresultIfZero(Ttype expr, HRESULT hr)
{
	if (expr == 0)
	{
		THROW_HR(hr);
	}
	return expr;
}
} // namespace

namespace util
{

// constants for step by step increasing buffer size for Win32 at < UINT16_MAX
enum class bufferstep : uint16_t
{
	bufferslice = UINT8_MAX, // common buffer size, 255 0xFF (MAX_PATH, UINT8_MAX)
	countmax = 8, // 0 1 2 3 4 5 6 7 8
	countlimit = 1 << countmax, // buffercount * 2 ... => buffer size is increased 8 times => max size 0xFF00 (65280) < UINT16_MAX
};

//uint16_t i = 1u;
//do
//{
//	uint16_t buffercount = (1u << i) * std::to_underlying(bufferstep::bufferslice);
//	buffer = std::make_unique<wchar_t[]>(buffercount);
//    ...
//	if (cond)
//		break;
//} while(i++ <= std::to_underlying(bufferstep::countmax));

//#define MAKEWORD(a, b)      ((WORD)(((BYTE)(((DWORD_PTR)(a)) & 0xff)) | ((WORD)((BYTE)(((DWORD_PTR)(b)) & 0xff))) << 8))
//#define MAKELONG(a, b)      ((LONG)(((WORD)(((DWORD_PTR)(a)) & 0xffff)) | ((DWORD)((WORD)(((DWORD_PTR)(b)) & 0xffff))) << 16))
//#define LOWORD(l)           ((WORD)(((DWORD_PTR)(l)) & 0xffff))
//#define HIWORD(l)           ((WORD)((((DWORD_PTR)(l)) >> 16) & 0xffff))
//#define LOBYTE(w)           ((BYTE)(((DWORD_PTR)(w)) & 0xff))
//#define HIBYTE(w)           ((BYTE)((((DWORD_PTR)(w)) >> 8) & 0xff))
//#define GET_X_LPARAM(lp)    ((int)(short)LOWORD(lp))
//#define GET_Y_LPARAM(lp)    ((int)(short)HIWORD(lp))
//#define MAKEINTRESOURCEW(i) ((LPWSTR)((ULONG_PTR)((WORD)(i))))
//#define POINTTOPOINTS(pt)   (MAKELONG((short)((pt).x), (short)((pt).y)))
//#define MAKEWPARAM(l, h)    ((WPARAM)(DWORD)MAKELONG(l, h))
//#define MAKELPARAM(l, h)    ((LPARAM)(DWORD)MAKELONG(l, h))
//#define MAKELRESULT(l, h)   ((LRESULT)(DWORD)MAKELONG(l, h))
//#define IS_INTRESOURCE(i)   ((((ULONG_PTR)(i)) >> 16) == 0)
//#define MAKELANGID(p, s)    ((((WORD)(s)) << 10) | (WORD)(p))
//#define PRIMARYLANGID(lgid) ((WORD  )(lgid) & 0x3ff)
//#define SUBLANGID(lgid)     ((WORD  )(lgid) >> 10)

template <typename Ttype> constexpr uint8_t			// LOBYTE
LoByte(Ttype type)
{
	return static_cast<uint8_t>(static_cast<uintptr_t>(type) & 0xff);
}
template <typename Ttype> constexpr uint8_t			// HIBYTE
HiByte(Ttype type)
{
	return static_cast<uint8_t>((static_cast<uintptr_t>(type) >> 8) & 0xff);
}
template <typename Ttype> constexpr uint16_t		// LOWORD
LoWord(Ttype type)
{
	return static_cast<uint16_t>(static_cast<uintptr_t>(type) & 0xffff);
}
template <typename Ttype> constexpr uint16_t		// HIWORD
HiWord(Ttype type)
{
	return static_cast<uint16_t>((static_cast<uintptr_t>(type) >> 16) & 0xffff);
}
template <typename Ttype> constexpr int32_t			// GET_X_LPARAM
GetXlparam(Ttype type)
{
	return static_cast<int32_t>(static_cast<int16_t>(static_cast<uint16_t>(static_cast<uintptr_t>(type) & 0xffff)));
}
template <typename Ttype> constexpr int32_t			// GET_Y_LPARAM
GetYlparam(Ttype type)
{
	return static_cast<int32_t>(static_cast<int16_t>(static_cast<uint16_t>(static_cast<uintptr_t>(type) >> 16) & 0xffff));
}
inline constexpr uint16_t							// MAKELANGID
MakeLangid(uint8_t lhs, uint8_t rhs)
{
	return static_cast<uint16_t>(((static_cast<uint16_t>(rhs) << 10u) | static_cast<uint16_t>(lhs)));
}
inline constexpr uint16_t							// MAKEWORD
MakeWord(uint8_t lhs, uint8_t rhs)
{
	return static_cast<uint16_t>((static_cast<uint8_t>((static_cast<uintptr_t>(lhs)) & 0xff)) 
		| (static_cast<uint16_t>(static_cast<uint8_t>((static_cast<uintptr_t>(rhs)) & 0xff)) << 8));
}
template <typename Ttype32, typename Ttype16> constexpr Ttype32
Make32bit(Ttype16 lhs, Ttype16 rhs)
{
	return static_cast<Ttype32>((static_cast<uint16_t>((static_cast<uintptr_t>(lhs)) & 0xffff)) 
		| (static_cast<uint32_t>(static_cast<uint16_t>((static_cast<uintptr_t>(rhs)) & 0xffff)) << 16));
}
inline constexpr int32_t							// MAKELONG
MakeLong(uint16_t lhs, uint16_t rhs)
{
	return Make32bit<int32_t, uint16_t>(lhs, rhs);
}
inline constexpr uint32_t
MakeUlong(uint16_t lhs, uint16_t rhs)
{
	return Make32bit<uint32_t, uint16_t>(lhs, rhs);
}
inline constexpr intptr_t
MakeIntptr(uint16_t lhs, uint16_t rhs)
{
	return Make32bit<intptr_t, uint16_t>(lhs, rhs);
}
inline constexpr uintptr_t
MakeUintptr(uint16_t lhs, uint16_t rhs)
{
	return Make32bit<uintptr_t, uint16_t>(lhs, rhs);
}

// C++ function aliases
const auto MakeWparam = MakeUintptr;				// MAKEWPARAM
const auto MakeLparam = MakeIntptr;					// MAKELPARAM
const auto MakeLresult = MakeIntptr;				// MAKELRESULT

// constant id numbers casted to pointers cannot be constexpr and is not expected to be by the API - used in resources
template <typename TtypePtr> TtypePtr
MakeIdptr(uint16_t id)
{
	return std::bit_cast<TtypePtr>(static_cast<uintptr_t>(id) & UINT16_MAX);
}
template <typename Tchar> Tchar*
MakeIntResource(uint16_t resid)
{
	return MakeIdptr<Tchar*>(resid);
}
inline PWSTR
MakeIntResourceW(uint16_t resid)
{
	return MakeIdptr<PWSTR>(resid);
}
template <typename Tenum> PWSTR
MakeEnumResource(Tenum resid)
{
	return MakeIdptr<PWSTR>(std::to_underlying(resid));
}
template <typename TtypePtr> bool
IsIntResource(TtypePtr* resptr)						// IS_INTRESOURCE
{
	return ((std::bit_cast<uintptr_t>(resptr) >> 16) == 0);		// uint16_t
}

static_assert(LoByte(0x1000u)                       == LOBYTE(0x1000u), "LoByte conversion template is wrong");
static_assert(HiByte(0x1000u)                       == HIBYTE(0x1000u), "HiByte conversion template is wrong");
static_assert(LoWord(0x10000000u)                   == LOWORD(0x10000000u), "LoWord conversion template is wrong");
static_assert(HiWord(0x10000000u)                   == HIWORD(0x10000000u), "HiWord conversion template is wrong");
static_assert(GetXlparam(0xFFFFFFFFu)               == GET_X_LPARAM(0xFFFFFFFFu), "GetXlparam conversion template is wrong");
static_assert(GetYlparam(0xFFFFFFFFu)               == GET_Y_LPARAM(0xFFFFFFFFu), "GetYlparam conversion template is wrong");
static_assert(MakeLangid(0x09, 0x01)                == MAKELANGID(LANG_ENGLISH, SUBLANG_ENGLISH_US), "MakeLangid conversion template is wrong");
static_assert(MakeWord(0xFEu, 0xEFu)                == MAKEWORD(0xFEu, 0xEFu), "MakeWord conversion is wrong");
static_assert(MakeLong(0xBEEFu, 0xDEADu)            == MAKELONG(0xBEEFu, 0xDEADu), "MakeLong conversion is wrong");
static_assert(MakeWparam(0xBEEFu, 0xDEADu)          == MAKEWPARAM(0xBEEFu, 0xDEADu), "MakeWparam conversion is wrong");
static_assert(MakeLparam(0xBEEFu, 0xDEADu)          == MAKELPARAM(0xBEEFu, 0xDEADu), "MakeLparam conversion is wrong");
static_assert(MakeLresult(0xBEEFu, 0xDEADu)         == MAKELRESULT(0xBEEFu, 0xDEADu), "MakeLresult conversion is wrong");

} // namespace util

// clang-format on

///////////////////////////////////////////////////////////////////////////////
// ModuleHelper - helper functions for ATL (deprecated)

namespace ModuleHelper
{
inline ATL::CAtlModule* GetModulePtr()
{
	return ATL::_pAtlModule;
}

inline HINSTANCE GetModuleInstance(void) noexcept
{
	return ATL::_AtlBaseModule.GetModuleInstance();
}

inline HINSTANCE GetResourceInstance(void) noexcept
{
	return ATL::_AtlBaseModule.GetResourceInstance();
}

inline void AddCreateWndData(ATL::_AtlCreateWndData* data, void* pObject) noexcept
{
	ATL::_AtlWinModule.AddCreateWndData(data, pObject);
}

inline void* ExtractCreateWndData(void) noexcept
{
	return ATL::_AtlWinModule.ExtractCreateWndData();
}
} // namespace ModuleHelper

template <typename... T>
std::wstring concat(T... args)
{
	std::wstring result;
	std::wstring_view substrings[] {args...};
	std::wstring::size_type fullsize = 0;
	for (auto sw : substrings)
		fullsize += sw.size();
	result.reserve(fullsize);
	for (auto sw : substrings)
		result.append(sw);
	return result;
}
