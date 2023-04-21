#include "smartrename_pch.h"
#include "common.h"

#include "Helpers.h"
//#include <ShlGuid.h>

namespace helpers
{

// STRINGTABLE resource
struct RES_STRINGTABLE
{
	uint16_t count;
#pragma warning(suppress : 4200)
	wchar_t charray[];
};

/// <summary>
/// <c>LoadResourceString</c> is a modified version of the LoadStringW function.
/// </summary>
/// <remarks>See atlcore.h _AtlGetStringResourceImage() and
/// https://blogs.msdn.microsoft.com/oldnewthing/20040130-00/?p=40813/
/// </remarks>
/// <param name="insthandle">resource DLL handle</param>
/// <param name="resstring">string to recieve resource</param>
/// <param name="resid">resource ID</param>
/// <param name="langid">resource language</param>
/// <returns>HRESULT</returns>
HRESULT LoadResourceString(_In_ HMODULE instancehandle, _Out_ std::wstring_view& resstring,
	_In_ uint16_t resid, _In_ uint16_t langid /* = UINT16_MAX*/)
{
	// string tables are counted UTF16 strings grouped together in bundles of sixteen
	// https://blogs.msdn.microsoft.com/oldnewthing/20040130-00/?p=40813/

	wchar_t* txtid = std::bit_cast<wchar_t*>(static_cast<uintptr_t>(resid / 16 + 1) & UINT16_MAX);
	if (langid == UINT16_MAX) // invalid
	{
		langid = ::GetUserDefaultUILanguage();
	}

	HRSRC reshandle = ::FindResourceExW(instancehandle, RT_STRING, txtid, langid);
	if (reshandle == nullptr)
	{
		langid = 0u; //util::MakeLangid(LANG_NEUTRAL, SUBLANG_NEUTRAL);
		reshandle = ::FindResourceExW(instancehandle, RT_STRING, txtid, langid);
		if (reshandle == nullptr)
		{
			// try strings in English as a last resort
			langid = util::MakeLangid(LANG_ENGLISH, SUBLANG_ENGLISH_US); // 1033, 0x409
			reshandle = ThrowIfNullptr(::FindResourceExW(instancehandle, RT_STRING, txtid, langid));
		}
	}

	HGLOBAL memhandle = ThrowIfNullptr(::LoadResource(instancehandle, reshandle));
	RES_STRINGTABLE* pstringtable = ThrowIfNullptr(static_cast<RES_STRINGTABLE*>(::LockResource(memhandle)));

	// from atlcore.h _AtlGetStringResourceImage()
	uint32_t ressize = ThrowIfZero(::SizeofResource(instancehandle, reshandle));
	uintptr_t resptr = std::bit_cast<uintptr_t>(pstringtable);
	uintptr_t resend = resptr + ressize;

	auto index = resid & 0x000fu;
	while ((index > 0) && (resptr < resend))
	{
		resptr += sizeof(*pstringtable) + (pstringtable->count * sizeof(wchar_t));
		pstringtable = std::bit_cast<RES_STRINGTABLE*>(resptr);
		index--;
	}

	// create an exact view of the resource in the supplied string reference
	// the resource string may not null-terminated
	resstring = std::wstring_view(pstringtable->charray, pstringtable->count);

	return S_OK;
}

HRESULT LoadResourceString(_In_ uint16_t resid, _Out_ std::wstring_view& resstring)
{
	return LoadResourceString(ModuleHelper::GetResourceInstance(), resstring, resid, UINT16_MAX);
}

std::wstring_view LoadResourceString(_In_ uint16_t resid)
{
	std::wstring_view resstring;
	HRESULT hr = LoadResourceString(ModuleHelper::GetResourceInstance(), resstring, resid, UINT16_MAX);
	return SUCCEEDED(hr) ? resstring : std::wstring_view();
}

HRESULT GetIconIndexFromPath(_In_ PCWSTR path, _Out_ int* index)
{
	*index = 0;

	HRESULT hr = E_FAIL;

	SHFILEINFOW shFileInfo = {0};

	if (!::PathIsRelativeW(path))
	{
		uint32_t attrib = GetFileAttributes(path);
		HIMAGELIST himl = (HIMAGELIST)::SHGetFileInfoW(path, attrib, &shFileInfo, sizeof(shFileInfo), (SHGFI_SYSICONINDEX | SHGFI_SMALLICON | SHGFI_USEFILEATTRIBUTES));
		if (himl)
		{
			*index = shFileInfo.iIcon;
			// We shouldn't free the HIMAGELIST.
			hr = S_OK;
		}
	}

	return hr;
}

HBITMAP CreateBitmapFromIcon(_In_ HICON hIcon, _In_opt_ uint32_t width, _In_opt_ uint32_t height)
{
	HBITMAP hBitmapResult = nullptr;

	// Create compatible DC
	HDC hDC = CreateCompatibleDC(nullptr);
	if (hDC != nullptr)
	{
		// Get bitmap rectangle size
		RECT rc = {0};
		rc.left = 0;
		rc.right = (width != 0) ? width : GetSystemMetrics(SM_CXSMICON);
		rc.top = 0;
		rc.bottom = (height != 0) ? height : GetSystemMetrics(SM_CYSMICON);

		// Create bitmap compatible with DC
		BITMAPINFO BitmapInfo;
		ZeroMemory(&BitmapInfo, sizeof(BITMAPINFO));

		BitmapInfo.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
		BitmapInfo.bmiHeader.biWidth = rc.right;
		BitmapInfo.bmiHeader.biHeight = rc.bottom;
		BitmapInfo.bmiHeader.biPlanes = 1;
		BitmapInfo.bmiHeader.biBitCount = 32;
		BitmapInfo.bmiHeader.biCompression = BI_RGB;

		HDC hDCBitmap = GetDC(nullptr);

		HBITMAP hBitmap = CreateDIBSection(hDCBitmap, &BitmapInfo, DIB_RGB_COLORS, nullptr, nullptr, 0);

		ReleaseDC(nullptr, hDCBitmap);

		if (hBitmap != nullptr)
		{
			// Select bitmap into DC
			HBITMAP hBitmapOld = (HBITMAP)SelectObject(hDC, hBitmap);
			if (hBitmapOld != nullptr)
			{
				// Draw icon into DC
				if (DrawIconEx(hDC, 0, 0, hIcon, rc.right, rc.bottom, 0, nullptr, DI_NORMAL))
				{
					// Restore original bitmap in DC
					hBitmapResult = (HBITMAP)SelectObject(hDC, hBitmapOld);
					hBitmapOld = nullptr;
					hBitmap = nullptr;
				}

				if (hBitmapOld != nullptr)
				{
					SelectObject(hDC, hBitmapOld);
				}
			}

			if (hBitmap != nullptr)
			{
				DeleteObject(hBitmap);
			}
		}

		DeleteDC(hDC);
	}

	return hBitmapResult;
}

HWND CreateMsgWindow(_In_ HINSTANCE hInst, _In_ WNDPROC pfnWndProc, _In_ void* p)
{
	WNDCLASSWl wc = {0};
	PCWSTR wndClassName = L"MsgWindow";

	wc.lpfnWndProc = ::DefWindowProcWl;
	wc.cbWndExtra = sizeof(void*);
	wc.hInstance = hInst;
	wc.hbrBackground = (HBRUSH)(COLOR_BTNFACE + 1);
	wc.lpszClassName = wndClassName;

	::RegisterClassW(&wc);

	HWND hwnd = ::CreateWindowExW(
		0, wndClassName, nullptr, 0, 0, 0, 0, 0, HWND_MESSAGE, 0, hInst, nullptr);
	if (hwnd)
	{
		::SetWindowLongPtrW(hwnd, 0, (intptr_t)p);
		if (pfnWndProc)
		{
			::SetWindowLongPtrW(hwnd, GWLP_WNDPROC, (intptr_t)pfnWndProc);
		}
	}

	return hwnd;
}

HRESULT GetShellItemArrayFromUnknown(_In_ IUnknown* punk, _COM_Outptr_ IShellItemArray** ppsia)
{
	*ppsia = nullptr;
	ATL::CComPtr<IDataObject> spdo;
	HRESULT hr;
	if (SUCCEEDED(punk->QueryInterface(IID_PPV_ARGS(&spdo))))
	{
		hr = SHCreateShellItemArrayFromDataObject(spdo, IID_PPV_ARGS(ppsia));
	}
	else
	{
		hr = punk->QueryInterface(IID_PPV_ARGS(ppsia));
	}

	return hr;
}

BOOL GetEnumeratedFileName(__out_ecount(cchMax) PWSTR pszUniqueName, uint32_t cchMax, __in PCWSTR pszTemplate, __in_opt PCWSTR pszDir, unsigned long ulMinLong, __inout unsigned long* pulNumUsed)
{
	PWSTR pszName = nullptr;
	HRESULT hr = S_OK;
	BOOL fRet = FALSE;
	int cchDir = 0;

	if (0 != cchMax && pszUniqueName)
	{
		*pszUniqueName = 0;
		if (pszDir)
		{
			hr = StringCchCopy(pszUniqueName, cchMax, pszDir);
			if (SUCCEEDED(hr))
			{
				size_t cch = wcslen(pszUniqueName);
				pszName = pszUniqueName + cch;
				size_t cchRemaining = cchMax - cch;
				if (cch && (L'\\' != *(pszName - 1)))
				{
					hr = StringCchCopyEx(pszName, cchRemaining, L"\\", &pszName, &cchRemaining, 0);
				}

				if (SUCCEEDED(hr))
				{
					cchDir = lstrlen(pszDir);
				}
			}
		}
		else
		{
			cchDir = 0;
			pszName = pszUniqueName;
		}
	}
	else
	{
		hr = E_INVALIDARG;
	}

	int cchTmp = 0;
	int cchStem = 0;
	PCWSTR pszStem = nullptr;
	PCWSTR pszRest = nullptr;
	wchar_t szFormat[MAX_PATH] = {0};

	if (SUCCEEDED(hr))
	{
		pszStem = pszTemplate;

		pszRest = ::StrChrW(pszTemplate, L'(');
		while (pszRest)
		{
			PCWSTR pszEndUniq = ::CharNextW(pszRest);
			while (*pszEndUniq && *pszEndUniq >= L'0' && *pszEndUniq <= L'9')
			{
				pszEndUniq++;
			}

			if (*pszEndUniq == L')')
			{
				break;
			}

			pszRest = ::StrChrW(::CharNextW(pszRest), L'(');
		}

		if (!pszRest)
		{
			pszRest = ::PathFindExtensionW(pszTemplate);
			cchStem = (int)(pszRest - pszTemplate);

			hr = StringCchCopy(szFormat, std::size(szFormat), L" (%lu)");
		}
		else
		{
			pszRest++;

			cchStem = (int)(pszRest - pszTemplate);

			while (*pszRest && *pszRest >= L'0' && *pszRest <= L'9')
			{
				pszRest++;
			}

			hr = StringCchCopy(szFormat, std::size(szFormat), L"%lu");
		}
	}

	unsigned long ulMax = 0;
	unsigned long ulMin = 0;
	if (SUCCEEDED(hr))
	{
		int cchFormat = lstrlen(szFormat);
		if (cchFormat < 3)
		{
			*pszUniqueName = L'\0';
			return FALSE;
		}
		ulMin = ulMinLong;
		cchTmp = cchMax - cchDir - cchStem - (cchFormat - 3);
		switch (cchTmp)
		{
		case 1:
			ulMax = 10;
			break;
		case 2:
			ulMax = 100;
			break;
		case 3:
			ulMax = 1000;
			break;
		case 4:
			ulMax = 10000;
			break;
		case 5:
			ulMax = 100000;
			break;
		default:
			if (cchTmp <= 0)
			{
				ulMax = ulMin;
			}
			else
			{
				ulMax = 1000000;
			}
			break;
		}
	}

	if (SUCCEEDED(hr))
	{
		hr = StringCchCopyN(pszName, pszUniqueName + cchMax - pszName, pszStem, cchStem);
		if (SUCCEEDED(hr))
		{
			PWSTR pszDigit = pszName + cchStem;

			for (unsigned long ul = ulMin; ((ul < ulMax) && (!fRet)); ul++)
			{
				wchar_t szTemp[MAX_PATH] = {0};
				hr = StringCchPrintf(szTemp, std::size(szTemp), szFormat, ul);
				if (SUCCEEDED(hr))
				{
					hr = StringCchCat(szTemp, std::size(szTemp), pszRest);
					if (SUCCEEDED(hr))
					{
						hr = StringCchCopy(pszDigit, pszUniqueName + cchMax - pszDigit, szTemp);
						if (SUCCEEDED(hr))
						{
							if (!::PathFileExistsW(pszUniqueName))
							{
								(*pulNumUsed) = ul;
								fRet = TRUE;
							}
						}
					}
				}
			}
		}
	}

	if (!fRet)
	{
		*pszUniqueName = L'\0';
	}

	return fRet;
}

} // namespace helpers