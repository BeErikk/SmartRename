#include "smartrename_pch.h"
#include "common.h"

#include "smartrenameinterfaces.h"
#include "resource.h"
#include "SmartRenameUI.h"
#include "helpers.h"
#include "settings.h"
#include "smartrenameenum.h"

//#include <commctrl.h>
//#include <Shlobj.h>
//#include <thread>
//#include <windowsx.h>

extern HINSTANCE g_hInst;

int g_rgnMatchModeResIDs[] = {
	IDS_ENTIREITEMNAME,
	IDS_NAMEONLY,
	IDS_EXTENSIONONLY};

enum
{
	MATCHMODE_FULLNAME = 0,
	MATCHMODE_NAMEONLY,
	MATCHMODE_EXTENIONONLY
};

struct FlagCheckboxMap
{
	uint32_t flag;
	uint32_t id;
};

FlagCheckboxMap g_flagCheckboxMap[] = {
	{UseRegularExpressions, IDC_CHECK_USEREGEX},
	{ExcludeSubfolders, IDC_CHECK_EXCLUDESUBFOLDERS},
	{EnumerateItems, IDC_CHECK_ENUMITEMS},
	{ExcludeFiles, IDC_CHECK_EXCLUDEFILES},
	{CaseSensitive, IDC_CHECK_CASESENSITIVE},
	{MatchAllOccurrences, IDC_CHECK_MATCHALLOCCURRENCES},
	{ExcludeFolders, IDC_CHECK_EXCLUDEFOLDERS},
	{NameOnly, IDC_CHECK_NAMEONLY},
	{ExtensionOnly, IDC_CHECK_EXTENSIONONLY}};

struct RepositionMap
{
	uint32_t id;
	uint32_t flags;
};

enum
{
	Reposition_None = 0,
	Reposition_X = 0x1,
	Reposition_Y = 0x2,
	Reposition_Width = 0x4,
	Reposition_Height = 0x8
};

RepositionMap g_repositionMap[] = {
	{IDC_SEARCHREPLACEGROUP, Reposition_Width},
	{IDC_OPTIONSGROUP, Reposition_Width},
	{IDC_PREVIEWGROUP, Reposition_Width | Reposition_Height},
	{IDC_EDIT_SEARCHFOR, Reposition_Width},
	{IDC_EDIT_REPLACEWITH, Reposition_Width},
	{IDC_LIST_PREVIEW, Reposition_Width | Reposition_Height},
	{IDC_STATUS_MESSAGE, Reposition_Y},
	{ID_RENAME, Reposition_X | Reposition_Y},
	{ID_ABOUT, Reposition_X | Reposition_Y},
	{IDCANCEL, Reposition_X | Reposition_Y}};

inline int RECT_WIDTH(RECT& r)
{
	return r.right - r.left;
}
inline int RECT_HEIGHT(RECT& r)
{
	return r.bottom - r.top;
}

#define MAX_INPUT_STRING_LEN 1024

HRESULT CSmartRenameUI::s_CreateInstance(_In_ ISmartRenameManager* psrm, _In_opt_ IUnknown* punk, _In_ bool enableDragDrop, _Outptr_ ISmartRenameUI** ppsrui)
{
	*ppsrui = nullptr;
	CSmartRenameUI* prui = new CSmartRenameUI();
	HRESULT hr = prui ? S_OK : E_OUTOFMEMORY;
	if (SUCCEEDED(hr))
	{
		// Pass the ISmartRenameManager to the ISmartRenameUI so it can subscribe to events
		hr = prui->_Initialize(psrm, punk, enableDragDrop);
		if (SUCCEEDED(hr))
		{
			hr = prui->QueryInterface(IID_PPV_ARGS(ppsrui));
		}
		prui->Release();
	}
	return hr;
}

// ISmartRenameUI
IFACEMETHODIMP CSmartRenameUI::Show(_In_opt_ HWND hwndParent)
{
	return _DoModeless(hwndParent);
}

IFACEMETHODIMP CSmartRenameUI::Close()
{
	_OnCloseDlg();
	return S_OK;
}

IFACEMETHODIMP CSmartRenameUI::Update()
{
	return S_OK;
}

IFACEMETHODIMP CSmartRenameUI::get_hwnd(_Out_ HWND* hwnd)
{
	*hwnd = m_hwnd;
	return S_OK;
}

IFACEMETHODIMP CSmartRenameUI::get_showUI(_Out_ bool* showUI)
{
	// Let callers know that it is OK to show UI (ex: progress dialog, error dialog and conflict dialog UI)
	*showUI = true;
	return S_OK;
}

// ISmartRenameManagerEvents
IFACEMETHODIMP CSmartRenameUI::OnItemAdded([[maybe_unused]]_In_ ISmartRenameItem* item)
{
	// Check if the user canceled the enumeration from the progress dialog UI
	if (m_srpui.IsCanceled())
	{
		m_srpui.Stop();
		if (m_spsre)
		{
			// Cancel the enumeration
			m_spsre->Cancel();
		}
	}
	return S_OK;
}

IFACEMETHODIMP CSmartRenameUI::OnUpdate(_In_ ISmartRenameItem*)
{
	uint32_t itemCount = 0;
	if (m_spsrm)
	{
		m_spsrm->GetItemCount(&itemCount);
	}
	m_listview.RedrawItems(0, itemCount);
	_UpdateCounts();
	return S_OK;
}

IFACEMETHODIMP CSmartRenameUI::OnError(_In_ ISmartRenameItem*)
{
	return S_OK;
}

IFACEMETHODIMP CSmartRenameUI::OnRegExStarted(_In_ uint32_t threadId)
{
	m_disableCountUpdate = true;
	m_currentRegExId = threadId;
	_UpdateCounts();
	return S_OK;
}

IFACEMETHODIMP CSmartRenameUI::OnRegExCanceled(_In_ uint32_t threadId)
{
	if (m_currentRegExId == threadId)
	{
		m_disableCountUpdate = false;
		_UpdateCounts();
	}

	return S_OK;
}

IFACEMETHODIMP CSmartRenameUI::OnRegExCompleted(_In_ uint32_t threadId)
{
	// Enable list view
	if (m_currentRegExId == threadId)
	{
		m_disableCountUpdate = false;
		_UpdateCounts();
	}
	return S_OK;
}

IFACEMETHODIMP CSmartRenameUI::OnRenameStarted()
{
	// Disable controls
	EnableWindow(m_hwnd, FALSE);
	return S_OK;
}

IFACEMETHODIMP CSmartRenameUI::OnRenameCompleted()
{
	// Enable controls
	EnableWindow(m_hwnd, TRUE);

	// Close the window
	::PostMessageW(m_hwnd, WM_CLOSE, 0u, 0);
	return S_OK;
}

// IDropTarget
IFACEMETHODIMP CSmartRenameUI::DragEnter(_In_ IDataObject* pdtobj, ULONG /* grfKeyState */, POINTL pt, _Inout_ ULONG* pdwEffect)
{
	if (m_spdth)
	{
		POINT ptT = {pt.x, pt.y};
		m_spdth->DragEnter(m_hwnd, pdtobj, &ptT, *pdwEffect);
	}

	return S_OK;
}

IFACEMETHODIMP CSmartRenameUI::DragOver(ULONG /* grfKeyState */, POINTL pt, _Inout_ ULONG* pdwEffect)
{
	if (m_spdth)
	{
		POINT ptT = {pt.x, pt.y};
		m_spdth->DragOver(&ptT, *pdwEffect);
	}

	return S_OK;
}

IFACEMETHODIMP CSmartRenameUI::DragLeave()
{
	if (m_spdth)
	{
		m_spdth->DragLeave();
	}

	return S_OK;
}

IFACEMETHODIMP CSmartRenameUI::Drop(_In_ IDataObject* pdtobj, ULONG, POINTL pt, _Inout_ ULONG* pdwEffect)
{
	if (m_spdth)
	{
		POINT ptT = {pt.x, pt.y};
		m_spdth->Drop(pdtobj, &ptT, *pdwEffect);
	}

	EnableWindow(GetDlgItem(m_hwnd, ID_RENAME), TRUE);
	EnableWindow(m_hwndLV, TRUE);

	// Populate the manager from the data object
	if (m_spsrm)
	{
		_EnumerateItems(pdtobj);
	}

	return S_OK;
}

HRESULT CSmartRenameUI::_Initialize(_In_ ISmartRenameManager* psrm, _In_opt_ IUnknown* punk, _In_ bool enableDragDrop)
{
	// Cache the smart rename manager
	m_spsrm = psrm;

	// Cache the punk for enumeration later
	m_spunk = punk;

	m_enableDragDrop = enableDragDrop;

	HRESULT hr = CoCreateInstance(CLSID_DragDropHelper, nullptr, CLSCTX_INPROC, IID_PPV_ARGS(&m_spdth));
	if (SUCCEEDED(hr))
	{
		// Subscribe to smart rename manager events
		hr = m_spsrm->Advise(this, &m_cookie);
	}

	if (FAILED(hr))
	{
		_Cleanup();
	}

	return hr;
}

HRESULT CSmartRenameUI::_InitAutoComplete()
{
	HRESULT hr = S_OK;
	if (CSettings::GetMRUEnabled())
	{
		hr = CoCreateInstance(CLSID_AutoComplete, nullptr, CLSCTX_INPROC, IID_PPV_ARGS(&m_spSearchAC));
		if (SUCCEEDED(hr))
		{
			hr = CRenameMRUSearch_CreateInstance(&m_spSearchACL);
			if (SUCCEEDED(hr))
			{
				hr = m_spSearchAC->Init(GetDlgItem(m_hwnd, IDC_EDIT_SEARCHFOR), m_spSearchACL, nullptr, nullptr);
				if (SUCCEEDED(hr))
				{
					hr = m_spSearchAC->SetOptions(ACO_AUTOSUGGEST | ACO_AUTOAPPEND | ACO_UPDOWNKEYDROPSLIST);
				}
			}
		}

		if (SUCCEEDED(hr))
		{
			hr = CoCreateInstance(CLSID_AutoComplete, nullptr, CLSCTX_INPROC, IID_PPV_ARGS(&m_spReplaceAC));
			if (SUCCEEDED(hr))
			{
				hr = CRenameMRUReplace_CreateInstance(&m_spReplaceACL);
				if (SUCCEEDED(hr))
				{
					hr = m_spReplaceAC->Init(GetDlgItem(m_hwnd, IDC_EDIT_REPLACEWITH), m_spReplaceACL, nullptr, nullptr);
					if (SUCCEEDED(hr))
					{
						hr = m_spReplaceAC->SetOptions(ACO_AUTOSUGGEST | ACO_AUTOAPPEND | ACO_UPDOWNKEYDROPSLIST);
					}
				}
			}
		}
	}

	return hr;
}

void CSmartRenameUI::_Cleanup()
{
	if (m_spsrm && m_cookie != 0)
	{
		m_spsrm->UnAdvise(m_cookie);
		m_cookie = 0;
		m_spsrm = nullptr;
	}

	m_spunk = nullptr;
	m_spdth = nullptr;

	if (m_enableDragDrop)
	{
		RevokeDragDrop(m_hwnd);
	}

	m_hwnd = nullptr;
}

HRESULT CSmartRenameUI::_EnumerateItems(_In_ IUnknown* punk)
{
	HRESULT hr = S_OK;
	// Enumerate and popuplate the manager
	if (m_spsrm)
	{
		m_disableCountUpdate = true;

		// Ensure we re-create the enumerator
		m_spsre = nullptr;
		hr = CSmartRenameEnum::s_CreateInstance(punk, m_spsrm, IID_PPV_ARGS(&m_spsre));
		if (SUCCEEDED(hr))
		{
			m_srpui.Start();
			hr = m_spsre->Start();
			m_srpui.Stop();
		}

		m_disableCountUpdate = false;

		if (SUCCEEDED(hr))
		{
			uint32_t itemCount = 0;
			m_spsrm->GetItemCount(&itemCount);
			m_listview.SetItemCount(itemCount);

			_UpdateCounts();
		}
	}

	return hr;
}

HRESULT CSmartRenameUI::_ReadSettings()
{
	// Check if we should read flags from settings
	// or the defaults from the manager.
	uint32_t flags = 0;
	if (CSettings::GetPersistState())
	{
		flags = CSettings::GetFlags();
		m_spsrm->put_flags(flags);

		wchar_t buffer[MAX_INPUT_STRING_LEN];
		buffer[0] = L'\0';
		CSettings::GetSearchText(buffer, std::size(buffer));
		::SetDlgItemTextW(m_hwnd, IDC_EDIT_SEARCHFOR, buffer);

		buffer[0] = L'\0';
		CSettings::GetReplaceText(buffer, std::size(buffer));
		::SetDlgItemTextW(m_hwnd, IDC_EDIT_REPLACEWITH, buffer);
	}
	else
	{
		m_spsrm->get_flags(&flags);
	}

	_SetCheckboxesFromFlags(flags);

	return S_OK;
}

HRESULT CSmartRenameUI::_WriteSettings()
{
	// Check if we should store our settings
	if (CSettings::GetPersistState())
	{
		uint32_t flags = 0;
		m_spsrm->get_flags(&flags);
		CSettings::SetFlags(flags);

		wchar_t buffer[MAX_INPUT_STRING_LEN] {};
		::GetDlgItemTextW(m_hwnd, IDC_EDIT_SEARCHFOR, buffer, static_cast<int32_t>(std::size(buffer)));
		CSettings::SetSearchText(buffer);

		if (CSettings::GetMRUEnabled() && m_spSearchACL)
		{
			ATL::CComPtr<ISmartRenameMRU> spSearchMRU;
			if (SUCCEEDED(m_spSearchACL->QueryInterface(IID_PPV_ARGS(&spSearchMRU))))
			{
				spSearchMRU->AddMRUString(buffer);
			}
		}

		buffer[0] = L'\0';
		::GetDlgItemTextW(m_hwnd, IDC_EDIT_REPLACEWITH, buffer, static_cast<int32_t>(std::size(buffer)));
		CSettings::SetReplaceText(buffer);

		if (CSettings::GetMRUEnabled() && m_spReplaceACL)
		{
			ATL::CComPtr<ISmartRenameMRU> spReplaceMRU;
			if (SUCCEEDED(m_spReplaceACL->QueryInterface(IID_PPV_ARGS(&spReplaceMRU))))
			{
				spReplaceMRU->AddMRUString(buffer);
			}
		}
	}

	return S_OK;
}

void CSmartRenameUI::_OnCloseDlg()
{
	if (m_modeless)
	{
		DestroyWindow(m_hwnd);
	}
	else
	{
		EndDialog(m_hwnd, 1);
	}
}

void CSmartRenameUI::_OnDestroyDlg()
{
	_Cleanup();

	if (m_modeless)
	{
		PostQuitMessage(0);
	}
}

void CSmartRenameUI::_OnRename()
{
	if (m_spsrm)
	{
		m_spsrm->Rename(m_hwnd);
	}

	// Persist the current settings.  We only do this when
	// a rename is actually performed.  Not when the user
	// closes/cancels the dialog.
	_WriteSettings();
}

void CSmartRenameUI::_OnAbout()
{
	// Launch github page
	SHELLEXECUTEINFOW info = {0};
	info.cbSize = sizeof(SHELLEXECUTEINFOW);
	info.lpVerb = L"open";
	info.lpFile = L"http://www.github.com/chrdavis";
	info.nShow = SW_SHOWDEFAULT;

	::ShellExecuteExW(&info);
}

HRESULT CSmartRenameUI::_DoModal(__in_opt HWND hwnd)
{
	m_modeless = false;
	HRESULT hr = S_OK;
	intptr_t ret = ::DialogBoxParamW(g_hInst, util::MakeIntResourceW(IDD_MAIN), hwnd, s_DlgProc, std::bit_cast<LPARAM>(this));
	if (ret < 0)
	{
		hr = HRESULT_FROM_WIN32(GetLastError());
	}
	return hr;
}

HRESULT CSmartRenameUI::_DoModeless(__in_opt HWND hwnd)
{
	m_modeless = true;
	HRESULT hr = S_OK;
	if (nullptr != ::CreateDialogParamW(g_hInst, util::MakeIntResourceW(IDD_MAIN), hwnd, s_DlgProc, std::bit_cast<LPARAM>(this)))
	{
		ShowWindow(m_hwnd, SW_SHOWNORMAL);
		MSG msg;
		while (::GetMessageW(&msg, nullptr, 0, 0))
		{
			if (!::IsDialogMessageW(m_hwnd, &msg))
			{
				::TranslateMessage(&msg);
				::DispatchMessageW(&msg);
			}
		}

		DestroyWindow(m_hwnd);
		m_hwnd = nullptr;
	}
	else
	{
		hr = HRESULT_FROM_WIN32(GetLastError());
	}
	return hr;
}

intptr_t CSmartRenameUI::_DlgProc(uint32_t uMsg, WPARAM wparam, LPARAM lparam)
{
	intptr_t bRet = TRUE; // default for all handled cases in switch below

	switch (uMsg)
	{
	case WM_INITDIALOG:
		_OnInitDlg();
		break;

	case WM_COMMAND:
		_OnCommand(wparam, lparam);
		break;

	case WM_NOTIFY:
		bRet = _OnNotify(wparam, lparam);
		break;

	case WM_THEMECHANGED:
		_OnSize(SIZE_RESTORED);
		break;

	case WM_SIZE:
		_OnSize(wparam);
		break;

	case WM_GETMINMAXINFO:
		_OnGetMinMaxInfo(lparam);
		break;

	case WM_CLOSE:
		_OnCloseDlg();
		break;

	case WM_DESTROY:
		_OnDestroyDlg();
		break;

	default:
		bRet = FALSE;
	}
	return bRet;
}

void CSmartRenameUI::_OnInitDlg()
{
	m_hwndLV = GetDlgItem(m_hwnd, IDC_LIST_PREVIEW);

	m_listview.Init(m_hwndLV);

	if (m_spunk)
	{
		// Populate the manager from the data object
		if (FAILED(_EnumerateItems(m_spunk)))
		{
			// Failed during enumeration.  Close the dialog.
			_OnCloseDlg();
			return;
		}
	}

	// Initialize from stored settings
	_ReadSettings();

	// Load the main icon
	LoadIconWithScaleDown(g_hInst, util::MakeIntResourceW(IDI_RENAME), 32, 32, &m_iconMain);

	// Update the icon associated with our main app window
	::SendMessageW(m_hwnd, WM_SETICON, ICON_SMALL, std::bit_cast<LPARAM>(m_iconMain.get()));
	::SendMessageW(m_hwnd, WM_SETICON, ICON_BIG, std::bit_cast<LPARAM>(m_iconMain.get()));

	// TODO: put this behind a setting?
	if (m_enableDragDrop)
	{
		RegisterDragDrop(m_hwnd, this);
	}

	RECT rc = {0};
	GetWindowRect(m_hwnd, &rc);
	m_initialWidth = RECT_WIDTH(rc);
	m_initialHeight = RECT_HEIGHT(rc);
	m_lastWidth = m_initialWidth;
	m_lastHeight = m_initialHeight;

	_InitAutoComplete();

	// Disable rename button by default.  It will be enabled in _UpdateCounts if
	// there are tiems to be renamed
	EnableWindow(GetDlgItem(m_hwnd, ID_RENAME), FALSE);

	// Update UI elements that depend on number of items selected or to be renamed
	_UpdateCounts();

	m_initialized = true;
}

void CSmartRenameUI::_OnCommand(_In_ WPARAM wparam,[[maybe_unused]] _In_ LPARAM lparam)
{
	switch (util::LoWord(wparam))
	{
	case IDOK:
	case IDCANCEL:
		_OnCloseDlg();
		break;

	case ID_RENAME:
		_OnRename();
		break;

	case ID_ABOUT:
		_OnAbout();
		break;

	case IDC_EDIT_REPLACEWITH:
	case IDC_EDIT_SEARCHFOR:
		if (util::HiWord(wparam) == EN_CHANGE) //if (GET_WM_COMMAND_CMD(wparam, lparam) == EN_CHANGE)
		{
			_OnSearchReplaceChanged();
		}
		break;

	case IDC_CHECK_CASESENSITIVE:
	case IDC_CHECK_ENUMITEMS:
	case IDC_CHECK_EXCLUDEFILES:
	case IDC_CHECK_EXCLUDEFOLDERS:
	case IDC_CHECK_EXCLUDESUBFOLDERS:
	case IDC_CHECK_MATCHALLOCCURRENCES:
	case IDC_CHECK_USEREGEX:
	case IDC_CHECK_EXTENSIONONLY:
	case IDC_CHECK_NAMEONLY:
		if (util::HiWord(wparam) == BN_CLICKED)
		{
			_ValidateFlagCheckbox(util::LoWord(wparam));
			_GetFlagsFromCheckboxes();
		}
		break;
	}
}

BOOL CSmartRenameUI::_OnNotify([[maybe_unused]]_In_ WPARAM wparam, _In_ LPARAM lparam)
{
	bool ret = FALSE;
	LPNMHDR pnmdr = std::bit_cast<LPNMHDR>(lparam);
	LPNMLISTVIEW pnmlv = std::bit_cast<LPNMLISTVIEW>(pnmdr);
	NMLVEMPTYMARKUP* pnmMarkup = nullptr;

	if (pnmdr)
	{
		//BOOL checked = FALSE;
		switch (pnmdr->code)
		{
		case HDN_ITEMSTATEICONCLICK:
			if (m_spsrm)
			{
				LPNMHEADERW pnmheader = std::bit_cast<LPNMHEADERW>(lparam);
				m_listview.ToggleAll(m_spsrm, (!(pnmheader->pitem->fmt & HDF_CHECKED)));
				_UpdateCounts();
			}
			break;

		case LVN_GETEMPTYMARKUP:
			pnmMarkup = std::bit_cast<NMLVEMPTYMARKUP*>(lparam);
			pnmMarkup->dwFlags = EMF_CENTERED;
			{
				std::wstring_view resstring { helpers::LoadResourceString(IDS_LISTVIEW_EMPTY) };	// localization
				if (!resstring.empty() && (resstring.size() < std::size(pnmMarkup->szMarkup)))
				{
					resstring.copy(pnmMarkup->szMarkup, resstring.size());
				}
			}
			ret = TRUE;
			break;

		case LVN_BEGINLABELEDITW:
			ret = TRUE;
			break;

		case LVN_KEYDOWN:
			if (m_spsrm)
			{
				m_listview.OnKeyDown(m_spsrm, (NMLVKEYDOWN*)pnmdr);
				_UpdateCounts();
			}
			break;

		case LVN_GETDISPINFOW:
			if (m_spsrm)
			{
				m_listview.GetDisplayInfo(m_spsrm, (NMLVDISPINFOW*)pnmlv);
			}
			break;

		case NM_CLICK:
		{
			if (m_spsrm)
			{
				m_listview.OnClickList(m_spsrm, (NMLISTVIEW*)pnmdr);
				_UpdateCounts();
			}
			break;
		}
		}
	}

	return ret;
}

void CSmartRenameUI::_OnGetMinMaxInfo(_In_ LPARAM lparam)
{
	if (m_initialWidth)
	{
		// Prevent resizing the dialog less than the original size
		MINMAXINFO* pMinMaxInfo = std::bit_cast<MINMAXINFO*>(lparam);
		pMinMaxInfo->ptMinTrackSize.x = m_initialWidth;
		pMinMaxInfo->ptMinTrackSize.y = m_initialHeight;
	}
}

void CSmartRenameUI::_OnSize(_In_ WPARAM wparam)
{
	if ((wparam == SIZE_RESTORED || wparam == SIZE_MAXIMIZED) && m_initialWidth)
	{
		// Calculate window size change delta
		RECT rc = {0};
		GetWindowRect(m_hwnd, &rc);

		const int xDelta = RECT_WIDTH(rc) - m_lastWidth;
		m_lastWidth += xDelta;
		const int yDelta = RECT_HEIGHT(rc) - m_lastHeight;
		m_lastHeight += yDelta;

		for (uint32_t u = 0; u < std::size(g_repositionMap); u++)
		{
			_MoveControl(g_repositionMap[u].id, g_repositionMap[u].flags, xDelta, yDelta);
		}

		m_listview.OnSize();
	}
}

void CSmartRenameUI::_MoveControl(_In_ uint32_t id, _In_ uint32_t repositionFlags, _In_ int xDelta, _In_ int yDelta)
{
	HWND hwnd = GetDlgItem(m_hwnd, id);

	uint32_t flags = SWP_NOOWNERZORDER | SWP_NOZORDER | SWP_NOACTIVATE;
	if (!((repositionFlags & Reposition_X) || (repositionFlags & Reposition_Y)))
	{
		flags |= SWP_NOMOVE;
	}

	if (!((repositionFlags & Reposition_Width) || (repositionFlags & Reposition_Height)))
	{
		flags |= SWP_NOSIZE;
	}

	RECT rcWindow = {0};
	GetWindowRect(hwnd, &rcWindow);

	int cx = RECT_WIDTH(rcWindow);
	int cy = RECT_HEIGHT(rcWindow);

	MapWindowPoints(HWND_DESKTOP, GetParent(hwnd), (LPPOINT)&rcWindow, 2);

	int x = rcWindow.left;
	int y = rcWindow.top;

	if (repositionFlags & Reposition_X)
	{
		x += xDelta;
	}

	if (repositionFlags & Reposition_Y)
	{
		y += yDelta;
	}

	if (repositionFlags & Reposition_Width)
	{
		cx += xDelta;
	}

	if (repositionFlags & Reposition_Height)
	{
		cy += yDelta;
	}

	SetWindowPos(hwnd, nullptr, x, y, cx, cy, flags);

	RedrawWindow(hwnd, nullptr, nullptr, RDW_INVALIDATE);
}

void CSmartRenameUI::_OnSearchReplaceChanged()
{
	// Pass updated search and replace terms to the ISmartRenameRegEx handler
	ATL::CComPtr<ISmartRenameRegEx> spRegEx;
	if (m_spsrm && SUCCEEDED(m_spsrm->get_renameRegEx(&spRegEx)))
	{
		wchar_t buffer[MAX_INPUT_STRING_LEN];
		buffer[0] = L'\0';
		::GetDlgItemTextW(m_hwnd, IDC_EDIT_SEARCHFOR, buffer, std::size(buffer));
		spRegEx->put_searchTerm(buffer);

		buffer[0] = L'\0';
		::GetDlgItemTextW(m_hwnd, IDC_EDIT_REPLACEWITH, buffer, std::size(buffer));
		spRegEx->put_replaceTerm(buffer);
	}
}

uint32_t CSmartRenameUI::_GetFlagsFromCheckboxes()
{
	uint32_t flags = 0;
	for (int i = 0; i < std::size(g_flagCheckboxMap); i++)
	{
		if (Button_GetCheck(GetDlgItem(m_hwnd, g_flagCheckboxMap[i].id)) == BST_CHECKED)
		{
			flags |= g_flagCheckboxMap[i].flag;
		}
	}

	// Ensure we update flags
	if (m_spsrm)
	{
		m_spsrm->put_flags(flags);
	}

	return flags;
}

void CSmartRenameUI::_SetCheckboxesFromFlags(_In_ uint32_t flags)
{
	for (int i = 0; i < std::size(g_flagCheckboxMap); i++)
	{
		Button_SetCheck(GetDlgItem(m_hwnd, g_flagCheckboxMap[i].id), flags & g_flagCheckboxMap[i].flag);
	}
}

void CSmartRenameUI::_ValidateFlagCheckbox(_In_ uint32_t checkBoxId)
{
	if (checkBoxId == IDC_CHECK_NAMEONLY)
	{
		if (Button_GetCheck(GetDlgItem(m_hwnd, IDC_CHECK_NAMEONLY)) == BST_CHECKED)
		{
			Button_SetCheck(GetDlgItem(m_hwnd, IDC_CHECK_EXTENSIONONLY), FALSE);
		}
	}
	else if (checkBoxId == IDC_CHECK_EXTENSIONONLY)
	{
		if (Button_GetCheck(GetDlgItem(m_hwnd, IDC_CHECK_EXTENSIONONLY)) == BST_CHECKED)
		{
			Button_SetCheck(GetDlgItem(m_hwnd, IDC_CHECK_NAMEONLY), FALSE);
		}
	}
}

void CSmartRenameUI::_UpdateCounts()
{
	// This method is CPU intensive.  We disable it during certain operations
	// for performance reasons.
	if (m_disableCountUpdate)
	{
		return;
	}

	uint32_t selectedCount = 0;
	uint32_t renamingCount = 0;
	if (m_spsrm)
	{
		m_spsrm->GetSelectedItemCount(&selectedCount);
		m_spsrm->GetRenameItemCount(&renamingCount);
	}

	if (m_selectedCount != selectedCount || m_renamingCount != renamingCount)
	{
		m_selectedCount = selectedCount;
		m_renamingCount = renamingCount;

		// Update selected and rename count label
		std::wstring_view countslabelformat { helpers::LoadResourceString(IDS_COUNTSLABELFMT) };
		std::wstring countslabel = std::format(countslabelformat, selectedCount, renamingCount);
		::SetDlgItemTextW(m_hwnd, IDC_STATUS_MESSAGE, countslabel);

		// Update Rename button state
		EnableWindow(GetDlgItem(m_hwnd, ID_RENAME), (renamingCount > 0));
	}
}

void CSmartRenameListView::Init(_In_ HWND hwndLV)
{
	if (hwndLV)
	{
		m_hwndLV = hwndLV;

		EnableWindow(m_hwndLV, TRUE);

		// Set the standard styles
		uint32_t dwLVStyle = (uint32_t)::GetWindowLongPtrW(m_hwndLV, GWL_STYLE);
		dwLVStyle |= LVS_ALIGNLEFT | LVS_REPORT | LVS_SHAREIMAGELISTS | LVS_SINGLESEL;
		::SetWindowLongPtrW(m_hwndLV, GWL_STYLE, dwLVStyle);

		// Set the extended view styles
		ListView_SetExtendedListViewStyle(m_hwndLV, LVS_EX_CHECKBOXES | LVS_EX_DOUBLEBUFFER | LVS_EX_AUTOSIZECOLUMNS);

		// Get the system image lists.  Our list view is setup to not destroy
		// these since the image list belongs to the entire explorer process
		HIMAGELIST himlLarge;
		HIMAGELIST himlSmall;
		if (Shell_GetImageLists(&himlLarge, &himlSmall))
		{
			ListView_SetImageList(m_hwndLV, himlSmall, LVSIL_SMALL);
			ListView_SetImageList(m_hwndLV, himlLarge, LVSIL_NORMAL);
		}

		_UpdateColumns();
	}
}

void CSmartRenameListView::ToggleAll(_In_ ISmartRenameManager* psrm, _In_ bool selected)
{
	if (m_hwndLV)
	{
		uint32_t itemCount = 0;
		psrm->GetItemCount(&itemCount);
		for (uint32_t i = 0; i < itemCount; i++)
		{
			ATL::CComPtr<ISmartRenameItem> spItem;
			if (SUCCEEDED(psrm->GetItemByIndex(i, &spItem)))
			{
				spItem->put_selected(selected);
			}
		}

		RedrawItems(0, itemCount);
	}
}

void CSmartRenameListView::ToggleItem(_In_ ISmartRenameManager* psrm, _In_ int item)
{
	ATL::CComPtr<ISmartRenameItem> spItem;
	if (SUCCEEDED(psrm->GetItemByIndex(item, &spItem)))
	{
		bool selected = false;
		spItem->get_selected(&selected);
		spItem->put_selected(!selected);

		RedrawItems(item, item);
	}
}

void CSmartRenameListView::OnKeyDown(_In_ ISmartRenameManager* psrm, _In_ NMLVKEYDOWN* lvKeyDown)
{
	if (lvKeyDown->wVKey == VK_SPACE)
	{
		int selectionMark = ListView_GetSelectionMark(m_hwndLV);
		if (selectionMark != -1)
		{
			ToggleItem(psrm, selectionMark);
		}
	}
}

void CSmartRenameListView::OnClickList(_In_ ISmartRenameManager* psrm, NMLISTVIEW* pnmListView)
{
	LVHITTESTINFO hitinfo;
	//Copy click point
	hitinfo.pt = pnmListView->ptAction;

	//Make the hit test...
	int item = ListView_HitTest(m_hwndLV, &hitinfo);
	if (item != -1)
	{
		if ((hitinfo.flags & LVHT_ONITEM) != 0)
		{
			ToggleItem(psrm, item);
		}
	}
}

void CSmartRenameListView::UpdateItemCheckState(_In_ ISmartRenameManager* psrm, _In_ int iItem)
{
	if (psrm && m_hwndLV && (iItem > -1))
	{
		ATL::CComPtr<ISmartRenameItem> spItem;
		if (SUCCEEDED(psrm->GetItemByIndex(iItem, &spItem)))
		{
			bool checked = ListView_GetCheckState(m_hwndLV, iItem);
			spItem->put_selected(checked);

			uint32_t uSelected = (checked) ? LVIS_SELECTED : 0;
			ListView_SetItemState(m_hwndLV, iItem, uSelected, LVIS_SELECTED);

			// Update the rename column if necessary
			int id = 0;
			spItem->get_id(&id);
			RedrawItems(id, id);
		}

		// Get the total number of list items and compare it to what is selected
		// We need to update the column checkbox if all items are selected or if
		// not all of the items are selected.
		bool checkHeader = (ListView_GetSelectedCount(m_hwndLV) == ListView_GetItemCount(m_hwndLV));
		_UpdateHeaderCheckState(checkHeader);
	}
}

#define COL_ORIGINAL_NAME 0
#define COL_NEW_NAME 1

void CSmartRenameListView::GetDisplayInfo(_In_ ISmartRenameManager* psrm, _Inout_ NMLVDISPINFOW* plvdi)
{
	uint32_t count = 0;
	psrm->GetItemCount(&count);
	if (plvdi->item.iItem < 0 || plvdi->item.iItem > static_cast<int>(count))
	{
		// Invalid index
		return;
	}

	ATL::CComPtr<ISmartRenameItem> renameItem;
	if (SUCCEEDED(psrm->GetItemByIndex((int)plvdi->item.iItem, &renameItem)))
	{
		if (plvdi->item.mask & LVIF_IMAGE)
		{
			renameItem->get_iconIndex(&plvdi->item.iImage);
		}

		if (plvdi->item.mask & LVIF_STATE)
		{
			plvdi->item.stateMask = LVIS_STATEIMAGEMASK;

			bool isSelected = false;
			renameItem->get_selected(&isSelected);
			if (isSelected)
			{
				// Turn check box on
				plvdi->item.state = INDEXTOSTATEIMAGEMASK(2);
			}
			else
			{
				// Turn check box off
				plvdi->item.state = INDEXTOSTATEIMAGEMASK(1);
			}
		}

		if (plvdi->item.mask & LVIF_PARAM)
		{
			int id = 0;
			renameItem->get_id(&id);
			plvdi->item.lparam = static_cast<LPARAM>(id);
		}

		if (plvdi->item.mask & LVIF_INDENT)
		{
			uint32_t depth = 0;
			renameItem->get_depth(&depth);
			plvdi->item.iIndent = static_cast<int>(depth);
		}

		if (plvdi->item.mask & LVIF_TEXT)
		{
			wil::unique_cotaskmem_string subItemText;
			if (plvdi->item.iSubItem == COL_ORIGINAL_NAME)
			{
				renameItem->get_originalName(&subItemText);
			}
			else if (plvdi->item.iSubItem == COL_NEW_NAME)
			{
				uint32_t flags = 0;
				psrm->get_flags(&flags);
				bool shouldRename = false;
				if (SUCCEEDED(renameItem->ShouldRenameItem(flags, &shouldRename)) && shouldRename)
				{
					renameItem->get_newName(&subItemText);
				}
			}

			StringCchCopy(plvdi->item.pszText, plvdi->item.cchTextMax, subItemText ? subItemText : L"");
			CoTaskMemFree(subItemText);
			subItemText = nullptr;
		}
	}
}

void CSmartRenameListView::OnSize()
{
	RECT rc = {0};
	GetClientRect(m_hwndLV, &rc);
	ListView_SetColumnWidth(m_hwndLV, 0, RECT_WIDTH(rc) / 2);
	ListView_SetColumnWidth(m_hwndLV, 1, RECT_WIDTH(rc) / 2);
}

void CSmartRenameListView::RedrawItems(_In_ uint32_t first, _In_ uint32_t last)
{
	ListView_RedrawItems(m_hwndLV, first, last);
}

void CSmartRenameListView::SetItemCount(_In_ uint32_t itemCount)
{
	ListView_SetItemCount(m_hwndLV, itemCount);
}

void CSmartRenameListView::_UpdateColumns()
{
	if (m_hwndLV)
	{
		// And the list view columns
		int iInsertPoint = 0;

		LVCOLUMNW lvc {};
		lvc.mask = LVCF_FMT | LVCF_ORDER | LVCF_WIDTH | LVCF_TEXT;
		lvc.fmt = LVCFMT_LEFT;
		lvc.iOrder = iInsertPoint;

		std::wstring_view original { helpers::LoadResourceString(IDS_ORIGINAL) };
		lvc.pszText = const_cast<PWSTR>(original.data());

		ListView_InsertColumn(m_hwndLV, iInsertPoint, &lvc);

		iInsertPoint++;

		lvc.iOrder = iInsertPoint;
		std::wstring_view renamed { helpers::LoadResourceString(IDS_RENAMED) };
		lvc.pszText = const_cast<PWSTR>(renamed.data());

		ListView_InsertColumn(m_hwndLV, iInsertPoint, &lvc);

		// Get a handle to the header of the columns
		HWND hwndHeader = ListView_GetHeader(m_hwndLV);

		if (hwndHeader)
		{
			// Update the header style to allow checkboxes
			uint32_t dwHeaderStyle = (uint32_t)::GetWindowLongPtrW(hwndHeader, GWL_STYLE);
			dwHeaderStyle |= HDS_CHECKBOXES;
			::SetWindowLongPtrW(hwndHeader, GWL_STYLE, dwHeaderStyle);

			_UpdateHeaderCheckState(TRUE);
		}

		_UpdateColumnSizes();
	}
}

void CSmartRenameListView::_UpdateColumnSizes()
{
	if (m_hwndLV)
	{
		RECT rc;
		GetClientRect(m_hwndLV, &rc);

		ListView_SetColumnWidth(m_hwndLV, 0, (rc.right - rc.left) / 2);
		ListView_SetColumnWidth(m_hwndLV, 1, (rc.right - rc.left) / 2);
	}
}

void CSmartRenameListView::_UpdateHeaderCheckState(_In_ bool check)
{
	// Get a handle to the header of the columns
	HWND hwndHeader = ListView_GetHeader(m_hwndLV);
	if (hwndHeader)
	{
		wchar_t buffer[MAX_PATH] = {0};

		// Retrieve the existing header first so we
		// don't trash the text already there
		HDITEMW hdi = {0};
		hdi.mask = HDI_FORMAT | HDI_TEXT;
		hdi.pszText = buffer;
		hdi.cchTextMax = std::size(buffer);

		Header_GetItem(hwndHeader, 0, &hdi);

		// Set the first column to contain a checkbox
		hdi.fmt |= HDF_CHECKBOX;
		hdi.fmt |= (check) ? HDF_CHECKED : 0;

		Header_SetItem(hwndHeader, 0, &hdi);
	}
}

#define TIMERID_CHECKCANCELED 101
#define CANCEL_CHECK_INTERVAL 500

HRESULT CSmartRenameProgressUI::Start()
{
	_Cleanup();
	m_canceled = false;
	AddRef();
	m_workerThreadHandle.reset(::CreateThread(nullptr, 0, s_workerThread, this, 0, nullptr));
	if (m_workerThreadHandle)
	{
		m_loadingThread = true;
	}
	else
	{
		Release();
	}
	return (m_workerThreadHandle) ? S_OK : E_FAIL;
}

uint32_t __stdcall CSmartRenameProgressUI::s_workerThread(_In_ void* pv)
{
	/*wil::unique_couninitialize_call*/ 
	auto call = wil::CoInitializeEx(COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE);
	//if (SUCCEEDED(CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE)))
	{
		CSmartRenameProgressUI* pThis = std::bit_cast<CSmartRenameProgressUI*>(pv);
		if (pThis)
		{
			HWND hwndMessage = helpers::CreateMsgWindow(g_hInst, s_msgWndProc, pThis);

			SetTimer(hwndMessage, TIMERID_CHECKCANCELED, CANCEL_CHECK_INTERVAL, nullptr);

			ATL::CComPtr<IProgressDialog> sppd;
			if (SUCCEEDED(CoCreateInstance(CLSID_ProgressDialog, nullptr, CLSCTX_INPROC, IID_PPV_ARGS(&sppd))))
			{
				pThis->m_sppd = sppd;
				std::wstring_view loading { helpers::LoadResourceString(IDS_LOADING) };
				sppd->SetLine(1, loading.data(), FALSE, nullptr);
				std::wstring_view loadingmsg { helpers::LoadResourceString(IDS_LOADING_MSG) };
				sppd->SetLine(2, loadingmsg.data(), FALSE, nullptr);
				std::wstring_view apptitle { helpers::LoadResourceString(IDS_APP_TITLE) };
				sppd->SetTitle(apptitle.data());

				SetTimer(hwndMessage, TIMERID_CHECKCANCELED, CANCEL_CHECK_INTERVAL, nullptr);
				sppd->StartProgressDialog(nullptr, nullptr, PROGDLG_MARQUEEPROGRESS, nullptr);
			}

			pThis->m_loadingThread = false;

			while (pThis->m_sppd && !sppd->HasUserCancelled())
			{
				MSG msg;
				while (::PeekMessageW(&msg, nullptr, 0, 0, PM_REMOVE))
				{
					::TranslateMessage(&msg);
					::DispatchMessageW(&msg);
				}
			}

			// Ensure dialog is stopped
			sppd->StopProgressDialog();

			KillTimer(hwndMessage, TIMERID_CHECKCANCELED);
			DestroyWindow(hwndMessage);

			pThis->Release();
		}

		//CoUninitialize();
	}

	return S_OK;
}

HRESULT CSmartRenameProgressUI::Stop()
{
	_Cleanup();
	return S_OK;
}

void CSmartRenameProgressUI::_Cleanup()
{
	while (m_loadingThread.load())
	{
		std::this_thread::sleep_for(std::chrono::milliseconds(5));
	}

	if (m_sppd)
	{
		m_sppd->StopProgressDialog();
		m_sppd = nullptr;
	}

	if (m_workerThreadHandle)
	{
		// Wait for up to 5 seconds for worker thread to finish
		WaitForSingleObject(m_workerThreadHandle, 5000);
		CloseHandle(m_workerThreadHandle);
		m_workerThreadHandle = nullptr;
	}
}

void CSmartRenameProgressUI::_UpdateCancelState()
{
	m_canceled = (m_sppd && m_sppd->HasUserCancelled());
}

LRESULT CALLBACK CSmartRenameProgressUI::s_msgWndProc(_In_ HWND hwnd, _In_ uint32_t uMsg, _In_ WPARAM wparam, _In_ LPARAM lparam)
{
	LRESULT lRes = 0;

	CSmartRenameProgressUI* pThis = (CSmartRenameProgressUI*)::GetWindowLongPtrW(hwnd, 0);
	if (pThis != nullptr)
	{
		lRes = pThis->_WndProc(hwnd, uMsg, wparam, lparam);
		if (uMsg == WM_NCDESTROY)
		{
			::SetWindowLongPtrW(hwnd, 0, nullptr);
		}
	}
	else
	{
		lRes = ::DefWindowProcW(hwnd, uMsg, wparam, lparam);
	}

	return lRes;
}

LRESULT CSmartRenameProgressUI::_WndProc(_In_ HWND hwnd, _In_ uint32_t msg, _In_ WPARAM wparam, _In_ LPARAM lparam)
{
	LRESULT lRes = 0;

	AddRef();

	switch (msg)
	{
	case WM_TIMER:
		if (wparam == TIMERID_CHECKCANCELED)
		{
			_UpdateCancelState();
		}
		break;

	case WM_DESTROY:
		_UpdateCancelState();
		KillTimer(hwnd, TIMERID_CHECKCANCELED);
		break;

	default:
		lRes = ::DefWindowProcW(hwnd, msg, wparam, lparam);
		break;
	}

	Release();

	return lRes;
}
