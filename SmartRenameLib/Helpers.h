#pragma once

namespace helpers
{

HRESULT LoadResourceString(_In_ HMODULE instancehandle, _Out_ std::wstring_view& resstring, _In_ uint16_t resid, _In_ uint16_t langid /* = UINT16_MAX*/);
HRESULT LoadResourceString(_In_ uint16_t resid, _Out_ std::wstring_view& resstring);
std::wstring_view LoadResourceString(_In_ uint16_t resid);

HRESULT GetIconIndexFromPath(_In_ PCWSTR path, _Out_ int* index);
HBITMAP CreateBitmapFromIcon(_In_ HICON hIcon, _In_opt_ uint32_t width = 0, _In_opt_ uint32_t height = 0);
HWND CreateMsgWindow(_In_ HINSTANCE hInst, _In_ WNDPROC pfnWndProc, _In_ void* p);
HRESULT GetShellItemArrayFromUnknown(_In_ IUnknown* punk, _COM_Outptr_ IShellItemArray** items);
BOOL GetEnumeratedFileName(
	__out_ecount(cchMax) PWSTR pszUniqueName,
	uint32_t cchMax,
	__in PCWSTR pszTemplate,
	__in_opt PCWSTR pszDir,
	unsigned long ulMinLong,
	__inout unsigned long* pulNumUsed);

} // namespace helpers
