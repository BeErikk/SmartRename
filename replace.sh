#!/bin/sh
# replace names of variables and types in code
# 2022-06-05 Jerker Bäck

find . -type f \( -name "*.h" -o -name "*.cpp" \) -exec sed -i \
	-e 's|\<APIENTRY\>|__stdcall|g' \
	-e 's|\<WINAPI\>|__stdcall|g' \
	-e 's|\<CONST\>|const|g' \
	-e 's|\<VOID\>|void|g' \
	-e 's|\<NULL\>|nullptr|g' \
	-e 's|\<UINT\>|uint32_t|g' \
	-e 's|\<INT\>|int32_t|g' \
	-e 's|\<LONG\>|int32_t|g' \
	-e 's|\<INT_PTR\>|intptr_t|g' \
	-e 's|\<UINT_PTR\>|uintptr_t|g' \
	-e 's|\<ULONG_PTR\>|uintptr_t|g' \
	-e 's|\<LONG_PTR\>|intptr_t|g' \
	-e 's|\<DWORD_PTR\>|uintptr_t|g' \
	-e 's|\<DWORD\>|uint32_t|g' \
	-e 's|\<WORD\>|uint16_t|g' \
	-e 's|\<hWnd\>|hwnd|g' \
	-e 's|\<hDlg\>|hdlg|g' \
	-e 's|\<wParam\>|wparam|g' \
	-e 's|\<lParam\>|lparam|g' \
	-e 's|\<CHAR\>|char|g' \
	-e 's|\<TCHAR\>|wchar_t|g' \
	-e 's|\<LPTSTR\>|PWSTR|g' \
	-e 's|\<LPCTSTR\>|PCWSTR|g' \
	{} \;

find . -type f \( -name "*.h" -o -name "*.cpp" \) -exec sed -i \
	-e 's|\<LOBYTE\>|util::LoByte|g' \
	-e 's|\<HIBYTE\>|util::HiByte|g' \
	-e 's|\<LOWORD\>|util::LoWord|g' \
	-e 's|\<HIWORD\>|util::HiWord|g' \
	-e 's|\<GET_X_LPARAM\>|util::GetXlparam|g' \
	-e 's|\<GET_Y_LPARAM\>|util::GetYlparam|g' \
	-e 's|\<MAKELANGID\>|util::MakeLangid|g' \
	-e 's|\<MAKEWORD\>|util::MakeWord|g' \
	-e 's|\<MAKELONG\>|util::MakeLong|g' \
	-e 's|\<MAKEWPARAM\>|util::MakeWparam|g' \
	-e 's|\<MAKELPARAM\>|util::MakeLparam|g' \
	-e 's|\<MAKELRESULT\>|util::MakeLresult|g' \
	{} \;
