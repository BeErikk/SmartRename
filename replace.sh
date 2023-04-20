#!/bin/sh
# replace names of variables and types in code
# 2022-06-05 Jerker Bäck

find ./winxp/ -type f \( -name "*.h" -o -name "*.cpp" \) -exec sed -i \
	-e 's|\<APIENTRY\>|__stdcall|g' \
	-e 's|\<CONST\>|const|g' \
	-e 's|\<VOID\>|void|g' \
	-e 's|\<NULL\>|nullptr|g' \
	-e 's|\<UINT\>|uint32_t|g' \
	-e 's|\<INT\>|int32_t|g' \
	-e 's|\<LONG\>|int32_t|g' \
	-e 's|\<INT_PTR\>|intptr_t|g' \
	-e 's|\<UINT_PTR\>|uintptr_t|g' \
	-e 's|\<ULONG_PTR\>|uintptr_t|g' \
	-e 's|\<DWORD\>|uint32_t|g' \
	-e 's|\<WORD\>|uint16_t|g' \
	-e 's|\<hWnd\>|hwnd|g' \
	-e 's|\<hDlg\>|hdlg|g' \
	-e 's|\<wParam\>|wparam|g' \
	-e 's|\<lParam\>|lparam|g' \
	-e 's|\<message\>|messageid|g' \
	-e 's|\<pszValue\>|valuename|g' \
	-e 's|\<ARRAYSIZE\>|std::size|g' \
	-e 's|\<CHAR\>|char|g' \
	-e 's|\<TCHAR\>|wchar_t|g' \
	-e 's|\<LPTSTR\>|PWSTR|g' \
	-e 's|\<LPCTSTR\>|PCWSTR|g' \
	{} \;
