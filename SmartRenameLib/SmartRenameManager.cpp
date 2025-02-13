#include "smartrename_pch.h"
#include "common.h"

#include "smartrenameinterfaces.h"
#include "srwlock.h"
#include "smartrenamemanager.h"
#include "smartrenameregex.h" // Default RegEx handler
#include "helpers.h"

//#include <algorithm>
//#include <shlobj.h>
//#include <versionhelpers.h>
//#include <filesystem>
//#include <vector>

namespace fs = std::filesystem;

extern HINSTANCE g_hInst;

IFACEMETHODIMP_(ULONG)
CSmartRenameManager::AddRef()
{
	return InterlockedIncrement(&m_refCount);
}

IFACEMETHODIMP_(ULONG)
CSmartRenameManager::Release()
{
	long refCount = InterlockedDecrement(&m_refCount);

	if (refCount == 0)
	{
		delete this;
	}
	return refCount;
}

IFACEMETHODIMP CSmartRenameManager::QueryInterface(_In_ REFIID riid, _Outptr_ void** ppv)
{
	static const QITAB qit[] = {
		QITABENT(CSmartRenameManager, ISmartRenameManager),
		QITABENT(CSmartRenameManager, ISmartRenameRegExEvents),
		{0}};
	return QISearch(this, qit, riid, ppv);
}

IFACEMETHODIMP CSmartRenameManager::Advise(_In_ ISmartRenameManagerEvents* renameOpEvents, _Out_ uint32_t* cookie)
{
	CSRWExclusiveAutoLock lock(&m_lockEvents);
	m_cookie++;
	RENAME_MGR_EVENT srme;
	srme.cookie = m_cookie;
	srme.pEvents = renameOpEvents;
	renameOpEvents->AddRef();
	m_renameManagerEvents.push_back(srme);

	*cookie = m_cookie;

	return S_OK;
}

IFACEMETHODIMP CSmartRenameManager::UnAdvise(_In_ uint32_t cookie)
{
	HRESULT hr = E_FAIL;
	CSRWExclusiveAutoLock lock(&m_lockEvents);

	for (std::vector<RENAME_MGR_EVENT>::iterator it = m_renameManagerEvents.begin(); it != m_renameManagerEvents.end(); ++it)
	{
		if (it->cookie == cookie)
		{
			hr = S_OK;
			it->cookie = 0;
			if (it->pEvents)
			{
				it->pEvents->Release();
				it->pEvents = nullptr;
			}
			break;
		}
	}

	return hr;
}

IFACEMETHODIMP CSmartRenameManager::Start()
{
	return E_NOTIMPL;
}

IFACEMETHODIMP CSmartRenameManager::Stop()
{
	return E_NOTIMPL;
}

IFACEMETHODIMP CSmartRenameManager::Rename(_In_ HWND hwndParent)
{
	m_hwndParent = hwndParent;
	return _PerformFileOperation();
}

IFACEMETHODIMP CSmartRenameManager::Reset()
{
	// Stop all threads and wait
	// Reset all rename items
	return E_NOTIMPL;
}

IFACEMETHODIMP CSmartRenameManager::Shutdown()
{
	_ClearRegEx();
	_Cleanup();
	return S_OK;
}

IFACEMETHODIMP CSmartRenameManager::AddItem(_In_ ISmartRenameItem* pItem)
{
	HRESULT hr = E_FAIL;
	// Scope lock
	{
		CSRWExclusiveAutoLock lock(&m_lockItems);
		int id = 0;
		pItem->get_id(&id);
		// Verify the item isn't already added
		if (m_renameItems.find(id) == m_renameItems.end())
		{
			m_renameItems[id] = pItem;
			pItem->AddRef();
			hr = S_OK;
		}
	}

	if (SUCCEEDED(hr))
	{
		_OnItemAdded(pItem);
	}

	return hr;
}

IFACEMETHODIMP CSmartRenameManager::GetItemByIndex(_In_ uint32_t index, _COM_Outptr_ ISmartRenameItem** ppItem)
{
	*ppItem = nullptr;
	CSRWSharedAutoLock lock(&m_lockItems);
	HRESULT hr = E_FAIL;
	if (index < m_renameItems.size())
	{
		std::map<int, ISmartRenameItem*>::iterator it = m_renameItems.begin();
		std::advance(it, index);
		*ppItem = it->second;
		(*ppItem)->AddRef();
		hr = S_OK;
	}

	return hr;
}

IFACEMETHODIMP CSmartRenameManager::GetItemById(_In_ int id, _COM_Outptr_ ISmartRenameItem** ppItem)
{
	*ppItem = nullptr;

	CSRWSharedAutoLock lock(&m_lockItems);
	HRESULT hr = E_FAIL;
	std::map<int, ISmartRenameItem*>::iterator it;
	it = m_renameItems.find(id);
	if (it != m_renameItems.end())
	{
		*ppItem = m_renameItems[id];
		(*ppItem)->AddRef();
		hr = S_OK;
	}

	return hr;
}

IFACEMETHODIMP CSmartRenameManager::GetItemCount(_Out_ uint32_t* count)
{
	CSRWSharedAutoLock lock(&m_lockItems);
	*count = static_cast<uint32_t>(m_renameItems.size());
	return S_OK;
}

IFACEMETHODIMP CSmartRenameManager::GetSelectedItemCount(_Out_ uint32_t* count)
{
	*count = 0;
	CSRWSharedAutoLock lock(&m_lockItems);

	for (std::map<int, ISmartRenameItem*>::iterator it = m_renameItems.begin(); it != m_renameItems.end(); ++it)
	{
		ISmartRenameItem* pItem = it->second;
		bool selected = false;
		if (SUCCEEDED(pItem->get_selected(&selected)) && selected)
		{
			(*count)++;
		}
	}

	return S_OK;
}

IFACEMETHODIMP CSmartRenameManager::GetRenameItemCount(_Out_ uint32_t* count)
{
	*count = 0;
	CSRWSharedAutoLock lock(&m_lockItems);

	for (std::map<int, ISmartRenameItem*>::iterator it = m_renameItems.begin(); it != m_renameItems.end(); ++it)
	{
		ISmartRenameItem* pItem = it->second;
		bool shouldRename = false;
		if (SUCCEEDED(pItem->ShouldRenameItem(m_flags, &shouldRename)) && shouldRename)
		{
			(*count)++;
		}
	}

	return S_OK;
}

IFACEMETHODIMP CSmartRenameManager::get_flags(_Out_ uint32_t* flags)
{
	_EnsureRegEx();
	*flags = m_flags;
	return S_OK;
}

IFACEMETHODIMP CSmartRenameManager::put_flags(_In_ uint32_t flags)
{
	if (flags != m_flags)
	{
		m_flags = flags;
		_EnsureRegEx();
		m_spRegEx->put_flags(flags);
	}
	return S_OK;
}

IFACEMETHODIMP CSmartRenameManager::get_renameRegEx(_COM_Outptr_ ISmartRenameRegEx** ppRegEx)
{
	*ppRegEx = nullptr;
	HRESULT hr = _EnsureRegEx();
	if (SUCCEEDED(hr))
	{
		*ppRegEx = m_spRegEx;
		(*ppRegEx)->AddRef();
	}
	return hr;
}

IFACEMETHODIMP CSmartRenameManager::put_renameRegEx(_In_ ISmartRenameRegEx* pRegEx)
{
	_ClearRegEx();
	m_spRegEx = pRegEx;
	return S_OK;
}

IFACEMETHODIMP CSmartRenameManager::get_renameItemFactory(_COM_Outptr_ ISmartRenameItemFactory** ppItemFactory)
{
	*ppItemFactory = nullptr;
	HRESULT hr = E_FAIL;
	if (m_spItemFactory)
	{
		hr = S_OK;
		*ppItemFactory = m_spItemFactory;
		(*ppItemFactory)->AddRef();
	}
	return hr;
}

IFACEMETHODIMP CSmartRenameManager::put_renameItemFactory(_In_ ISmartRenameItemFactory* pItemFactory)
{
	m_spItemFactory = pItemFactory;
	return S_OK;
}

IFACEMETHODIMP CSmartRenameManager::OnSearchTermChanged(_In_ PCWSTR /*searchTerm*/)
{
	_PerformRegExRename();
	return S_OK;
}

IFACEMETHODIMP CSmartRenameManager::OnReplaceTermChanged(_In_ PCWSTR /*replaceTerm*/)
{
	_PerformRegExRename();
	return S_OK;
}

IFACEMETHODIMP CSmartRenameManager::OnFlagsChanged(_In_ uint32_t flags)
{
	// Flags were updated in the smart rename regex.  Update our preview.
	m_flags = flags;
	_PerformRegExRename();
	return S_OK;
}

HRESULT CSmartRenameManager::s_CreateInstance(_Outptr_ ISmartRenameManager** ppsrm)
{
	*ppsrm = nullptr;
	CSmartRenameManager* psrm = new CSmartRenameManager();
	HRESULT hr = psrm ? S_OK : E_OUTOFMEMORY;
	if (SUCCEEDED(hr))
	{
		hr = psrm->_Init();
		if (SUCCEEDED(hr))
		{
			hr = psrm->QueryInterface(IID_PPV_ARGS(ppsrm));
		}
		psrm->Release();
	}
	return hr;
}

CSmartRenameManager::CSmartRenameManager()
	: m_refCount(1)
{
	InitializeCriticalSection(&m_critsecReentrancy);
}

CSmartRenameManager::~CSmartRenameManager()
{
	DeleteCriticalSection(&m_critsecReentrancy);
}

HRESULT CSmartRenameManager::_Init()
{
	// Guaranteed to succeed
	m_startFileOpWorkerEvent = CreateEvent(nullptr, TRUE, FALSE, nullptr);
	m_startRegExWorkerEvent = CreateEvent(nullptr, TRUE, FALSE, nullptr);
	m_cancelRegExWorkerEvent = CreateEvent(nullptr, TRUE, FALSE, nullptr);

	m_hwndMessage = helpers::CreateMsgWindow(g_hInst, s_msgWndProc, this);

	return S_OK;
}

// Custom messages for worker threads
enum
{
	SRM_REGEX_ITEM_UPDATED = (WM_APP + 1), // Single smart rename item processed by regex worker thread
	SRM_REGEX_STARTED, // RegEx operation was started
	SRM_REGEX_CANCELED, // Regex operation was canceled
	SRM_REGEX_COMPLETE, // Regex worker thread completed
	SRM_FILEOP_COMPLETE // File Operation worker thread completed
};

struct WorkerThreadData
{
	HWND hwndManager = nullptr;
	HANDLE startEvent = nullptr;
	HANDLE cancelEvent = nullptr;
	HWND hwndParent = nullptr;
	ATL::CComPtr<ISmartRenameManager> spsrm;
};

// Msg-only worker window proc for communication from our worker threads
LRESULT CALLBACK CSmartRenameManager::s_msgWndProc(_In_ HWND hwnd, _In_ uint32_t uMsg, _In_ WPARAM wparam, _In_ LPARAM lparam)
{
	LRESULT lRes = 0;

	CSmartRenameManager* pThis = (CSmartRenameManager*)::GetWindowLongPtrW(hwnd, 0);
	if (pThis != nullptr)
	{
		lRes = pThis->_WndProc(hwnd, uMsg, wparam, lparam);
		if (uMsg == WM_NCDESTROY)
		{
			::SetWindowLongPtrW(hwnd, 0, nullptr);
			pThis->m_hwndMessage = nullptr;
		}
	}
	else
	{
		lRes = ::DefWindowProcW(hwnd, uMsg, wparam, lparam);
	}

	return lRes;
}

LRESULT CSmartRenameManager::_WndProc(_In_ HWND hwnd, _In_ uint32_t msg, _In_ WPARAM wparam, _In_ LPARAM lparam)
{
	LRESULT lRes = 0;

	AddRef();

	switch (msg)
	{
	case SRM_REGEX_ITEM_UPDATED:
	{
		int id = static_cast<int>(lparam);
		ATL::CComPtr<ISmartRenameItem> spItem;
		if (SUCCEEDED(GetItemById(id, &spItem)))
		{
			_OnUpdate(spItem);
		}
		break;
	}
	case SRM_REGEX_STARTED:
		_OnRegExStarted(static_cast<uint32_t>(wparam));
		break;

	case SRM_REGEX_CANCELED:
		_OnRegExCanceled(static_cast<uint32_t>(wparam));
		break;

	case SRM_REGEX_COMPLETE:
		_OnRegExCompleted(static_cast<uint32_t>(wparam));
		break;

	default:
		lRes = ::DefWindowProcW(hwnd, msg, wparam, lparam);
		break;
	}

	Release();

	return lRes;
}

HRESULT CSmartRenameManager::_PerformFileOperation()
{
	// Do we have items to rename?
	uint32_t renameItemCount = 0;
	if (FAILED(GetRenameItemCount(&renameItemCount)) || renameItemCount == 0)
	{
		return E_FAIL;
	}

	// Wait for existing regex thread to finish
	_WaitForRegExWorkerThread();

	// Create worker thread which will perform the actual rename
	HRESULT hr = _CreateFileOpWorkerThread();
	if (SUCCEEDED(hr))
	{
		_OnRenameStarted();

		// Signal the worker thread that they can start working. We needed to wait until we
		// were ready to process thread messages.
		SetEvent(m_startFileOpWorkerEvent);

		while (true)
		{
			// Check if worker thread has exited
			if (WaitForSingleObject(m_fileOpWorkerThreadHandle, 0) == WAIT_OBJECT_0)
			{
				break;
			}

			MSG msg;
			while (::PeekMessageW(&msg, nullptr, 0, 0, PM_REMOVE))
			{
				if (msg.message == SRM_FILEOP_COMPLETE)
				{
					// Worker thread completed
					break;
				}
				else
				{
					::TranslateMessage(&msg);
					::DispatchMessageW(&msg);
				}
			}
		}

		_OnRenameCompleted();
	}

	return hr;
}

HRESULT CSmartRenameManager::_CreateFileOpWorkerThread()
{
	WorkerThreadData* pwtd = new WorkerThreadData;
	HRESULT hr = pwtd ? S_OK : E_OUTOFMEMORY;
	if (SUCCEEDED(hr))
	{
		pwtd->hwndManager = m_hwndMessage;
		pwtd->startEvent = m_startRegExWorkerEvent;
		pwtd->cancelEvent = nullptr;
		pwtd->spsrm = this;
		m_fileOpWorkerThreadHandle = CreateThread(nullptr, 0, s_fileOpWorkerThread, pwtd, 0, nullptr);
		hr = (m_fileOpWorkerThreadHandle) ? S_OK : E_FAIL;
		if (FAILED(hr))
		{
			delete pwtd;
		}
	}

	return hr;
}

// The default FOF flags to use in the rename operations
uint32_t CSmartRenameManager::_GetDefaultFileOpFlags()
{
	uint32_t flags = (FOF_ALLOWUNDO | FOFX_SHOWELEVATIONPROMPT | FOF_RENAMEONCOLLISION);
	if (IsWindows8OrGreater())
	{
		flags |= FOFX_ADDUNDORECORD;
	}
	return flags;
}

uint32_t __stdcall CSmartRenameManager::s_fileOpWorkerThread(_In_ void* pv)
{
	//if (SUCCEEDED(CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE)))
	auto call = wil::CoInitializeEx(COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE);
	{
		WorkerThreadData* pwtd = std::bit_cast<WorkerThreadData*>(pv);
		if (pwtd)
		{
			// Wait to be told we can begin
			if (WaitForSingleObject(pwtd->startEvent, INFINITE) == WAIT_OBJECT_0)
			{
				ATL::CComPtr<ISmartRenameRegEx> spRenameRegEx;
				if (SUCCEEDED(pwtd->spsrm->get_renameRegEx(&spRenameRegEx)))
				{
					// Create IFileOperation interface
					ATL::CComPtr<IFileOperation> spFileOp;
					if (SUCCEEDED(CoCreateInstance(CLSID_FileOperation, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&spFileOp))))
					{
						uint32_t flags = 0;
						spRenameRegEx->get_flags(&flags);

						uint32_t itemCount = 0;
						pwtd->spsrm->GetItemCount(&itemCount);

						// We add the items to the operation in depth-first order.  This allows child items to be
						// renamed before parent items.

						// Creating a vector of vectors of items of the same depth
						std::vector<std::vector<uint32_t>> matrix(itemCount);

						for (uint32_t u = 0; u < itemCount; u++)
						{
							ATL::CComPtr<ISmartRenameItem> spItem;
							if (SUCCEEDED(pwtd->spsrm->GetItemByIndex(u, &spItem)))
							{
								uint32_t depth = 0;
								spItem->get_depth(&depth);
								matrix[depth].push_back(u);
							}
						}

						// From the greatest depth first, add all items of that depth to the operation
						for (int32_t v = itemCount - 1; v >= 0; v--)
						{
							for (auto it : matrix[v])
							{
								ATL::CComPtr<ISmartRenameItem> spItem;
								if (SUCCEEDED(pwtd->spsrm->GetItemByIndex(it, &spItem)))
								{
									bool shouldRename = false;
									if (SUCCEEDED(spItem->ShouldRenameItem(flags, &shouldRename)) && shouldRename)
									{
										wil::unique_cotaskmem_string newName;
										if (SUCCEEDED(spItem->get_newName(&newName)))
										{
											ATL::CComPtr<IShellItem> spShellItem;
											if (SUCCEEDED(spItem->get_shellItem(&spShellItem)))
											{
												spFileOp->RenameItem(spShellItem, newName, nullptr);
											}
											//CoTaskMemFree(newName);
										}
									}
								}
							}
						}

						// Set the operation flags
						if (SUCCEEDED(spFileOp->SetOperationFlags(_GetDefaultFileOpFlags())))
						{
							// Set the parent window
							if (pwtd->hwndParent)
							{
								spFileOp->SetOwnerWindow(pwtd->hwndParent);
							}

							// Perform the operation
							// We don't care about the return code here. We would rather
							// return control back to explorer so the user can cleanly
							// undo the operation if it failed halfway through.
							spFileOp->PerformOperations();
						}
					}
				}
			}

			// Send the manager thread the completion message
			::PostMessageW(pwtd->hwndManager, SRM_FILEOP_COMPLETE, GetCurrentThreadId(), 0);

			delete pwtd;
		}
		//CoUninitialize();
	}

	return 0;
}

HRESULT CSmartRenameManager::_PerformRegExRename()
{
	HRESULT hr = E_FAIL;

	if (!TryEnterCriticalSection(&m_critsecReentrancy))
	{
		// Ensure we do not re-enter since we pump messages here.
		// TODO: If we do, post a message back to ourselves
	}
	else
	{
		// Ensure previous thread is canceled
		_CancelRegExWorkerThread();

		// Create worker thread which will message us progress and completion.
		hr = _CreateRegExWorkerThread();
		if (SUCCEEDED(hr))
		{
			ResetEvent(m_cancelRegExWorkerEvent);

			// Signal the worker thread that they can start working. We needed to wait until we
			// were ready to process thread messages.
			SetEvent(m_startRegExWorkerEvent);
		}
	}

	return hr;
}

HRESULT CSmartRenameManager::_CreateRegExWorkerThread()
{
	WorkerThreadData* pwtd = new WorkerThreadData;
	HRESULT hr = pwtd ? S_OK : E_OUTOFMEMORY;
	if (SUCCEEDED(hr))
	{
		pwtd->hwndManager = m_hwndMessage;
		pwtd->startEvent = m_startRegExWorkerEvent;
		pwtd->cancelEvent = m_cancelRegExWorkerEvent;
		pwtd->hwndParent = m_hwndParent;
		pwtd->spsrm = this;
		m_regExWorkerThreadHandle = CreateThread(nullptr, 0, s_regexWorkerThread, pwtd, 0, nullptr);
		hr = (m_regExWorkerThreadHandle) ? S_OK : E_FAIL;
		if (FAILED(hr))
		{
			delete pwtd;
		}
	}

	return hr;
}

uint32_t __stdcall CSmartRenameManager::s_regexWorkerThread(_In_ void* pv)
{
	//if (SUCCEEDED(CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE)))
	auto call = wil::CoInitializeEx(COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE);
	{
		WorkerThreadData* pwtd = std::bit_cast<WorkerThreadData*>(pv);
		if (pwtd)
		{
			::PostMessageW(pwtd->hwndManager, SRM_REGEX_STARTED, GetCurrentThreadId(), 0);

			// Wait to be told we can begin
			if (WaitForSingleObject(pwtd->startEvent, INFINITE) == WAIT_OBJECT_0)
			{
				ATL::CComPtr<ISmartRenameRegEx> spRenameRegEx;
				if (SUCCEEDED(pwtd->spsrm->get_renameRegEx(&spRenameRegEx)))
				{
					uint32_t flags = 0;
					spRenameRegEx->get_flags(&flags);

					uint32_t itemCount = 0;
					unsigned long itemEnumIndex = 1;
					pwtd->spsrm->GetItemCount(&itemCount);
					for (uint32_t u = 0; u <= itemCount; u++)
					{
						// Check if cancel event is signaled
						if (WaitForSingleObject(pwtd->cancelEvent, 0) == WAIT_OBJECT_0)
						{
							// Canceled from manager
							// Send the manager thread the canceled message
							::PostMessageW(pwtd->hwndManager, SRM_REGEX_CANCELED, GetCurrentThreadId(), 0);
							break;
						}

						ATL::CComPtr<ISmartRenameItem> spItem;
						if (SUCCEEDED(pwtd->spsrm->GetItemByIndex(u, &spItem)))
						{
							int id = -1;
							spItem->get_id(&id);

							bool isFolder = false;
							bool isSubFolderContent = false;
							spItem->get_isFolder(&isFolder);
							spItem->get_isSubFolderContent(&isSubFolderContent);
							if ((isFolder && (flags & SmartRenameFlags::ExcludeFolders)) || (!isFolder && (flags & SmartRenameFlags::ExcludeFiles)) || (isSubFolderContent && (flags & SmartRenameFlags::ExcludeSubfolders)))
							{
								// Exclude this item from renaming.  Ensure new name is cleared.
								spItem->put_newName(nullptr);

								// Send the manager thread the item processed message
								::PostMessageW(pwtd->hwndManager, SRM_REGEX_ITEM_UPDATED, GetCurrentThreadId(), id);

								continue;
							}

							PWSTR originalName = nullptr;
							if (SUCCEEDED(spItem->get_originalName(&originalName)))
							{
								PWSTR currentNewName = nullptr;
								spItem->get_newName(&currentNewName);

								wchar_t sourceName[MAX_PATH] = {0};
								if (flags & NameOnly)
								{
									StringCchCopy(sourceName, std::size(sourceName), fs::path(originalName).stem().c_str());
								}
								else if (flags & ExtensionOnly)
								{
									std::wstring extension = fs::path(originalName).extension().wstring();
									if (!extension.empty() && extension.front() == '.')
									{
										extension = extension.erase(0, 1);
									}
									StringCchCopy(sourceName, std::size(sourceName), extension.c_str());
								}
								else
								{
									StringCchCopy(sourceName, std::size(sourceName), originalName);
								}

								PWSTR newName = nullptr;
								// Failure here means we didn't match anything or had nothing to match
								// Call put_newName with null in that case to reset it
								spRenameRegEx->Replace(sourceName, &newName);

								wchar_t resultName[MAX_PATH] = {0};

								PWSTR newNameToUse = nullptr;

								// newName == nullptr likely means we have an empty search string.  We should leave newNameToUse
								// as nullptr so we clear the renamed column
								if (newName != nullptr)
								{
									newNameToUse = resultName;
									if (flags & NameOnly)
									{
										StringCchPrintf(resultName, std::size(resultName), L"%s%s", newName, fs::path(originalName).extension().c_str());
									}
									else if (flags & ExtensionOnly)
									{
										std::wstring extension = fs::path(originalName).extension().wstring();
										if (!extension.empty())
										{
											StringCchPrintf(resultName, std::size(resultName), L"%s.%s", fs::path(originalName).stem().c_str(), newName);
										}
										else
										{
											StringCchCopy(resultName, std::size(resultName), originalName);
										}
									}
									else
									{
										StringCchCopy(resultName, std::size(resultName), newName);
									}
								}

								// No change from originalName so set newName to
								// null so we clear it from our UI as well.
								if (lstrcmp(originalName, newNameToUse) == 0)
								{
									newNameToUse = nullptr;
								}

								wchar_t uniqueName[MAX_PATH] = {0};
								if (newNameToUse != nullptr && (flags & EnumerateItems))
								{
									unsigned long countUsed = 0;
									if (helpers::GetEnumeratedFileName(uniqueName, std::size(uniqueName), newNameToUse, nullptr, itemEnumIndex, &countUsed))
									{
										newNameToUse = uniqueName;
									}
									itemEnumIndex++;
								}

								spItem->put_newName(newNameToUse);

								// Was there a change?
								if (lstrcmp(currentNewName, newNameToUse) != 0)
								{
									// Send the manager thread the item processed message
									::PostMessageW(pwtd->hwndManager, SRM_REGEX_ITEM_UPDATED, GetCurrentThreadId(), id);
								}

								CoTaskMemFree(newName);
								CoTaskMemFree(currentNewName);
								CoTaskMemFree(originalName);
							}
						}
					}
				}
			}

			// Send the manager thread the completion message
			::PostMessageW(pwtd->hwndManager, SRM_REGEX_COMPLETE, GetCurrentThreadId(), 0);

			delete pwtd;
		}
		//CoUninitialize();
	}

	return 0;
}

void CSmartRenameManager::_CancelRegExWorkerThread()
{
	if (m_startRegExWorkerEvent)
	{
		SetEvent(m_startRegExWorkerEvent);
	}

	if (m_cancelRegExWorkerEvent)
	{
		SetEvent(m_cancelRegExWorkerEvent);
	}

	_WaitForRegExWorkerThread();
}

void CSmartRenameManager::_WaitForRegExWorkerThread()
{
	if (m_regExWorkerThreadHandle)
	{
		WaitForSingleObject(m_regExWorkerThreadHandle, INFINITE);
		CloseHandle(m_regExWorkerThreadHandle);
		m_regExWorkerThreadHandle = nullptr;
	}
}

void CSmartRenameManager::_Cancel()
{
	SetEvent(m_startFileOpWorkerEvent);
	_CancelRegExWorkerThread();
}

HRESULT CSmartRenameManager::_EnsureRegEx()
{
	HRESULT hr = S_OK;
	if (!m_spRegEx)
	{
		// Create the default regex handler
		hr = CSmartRenameRegEx::s_CreateInstance(&m_spRegEx);
		if (SUCCEEDED(hr))
		{
			hr = _InitRegEx();
			// Get the flags
			if (SUCCEEDED(hr))
			{
				m_spRegEx->get_flags(&m_flags);
			}
		}
	}
	return hr;
}

HRESULT CSmartRenameManager::_InitRegEx()
{
	HRESULT hr = E_FAIL;
	if (m_spRegEx)
	{
		hr = m_spRegEx->Advise(this, &m_regExAdviseCookie);
	}

	return hr;
}

void CSmartRenameManager::_ClearRegEx()
{
	if (m_spRegEx)
	{
		m_spRegEx->UnAdvise(m_regExAdviseCookie);
		m_regExAdviseCookie = 0;
	}
}

void CSmartRenameManager::_OnItemAdded(_In_ ISmartRenameItem* renameItem)
{
	CSRWSharedAutoLock lock(&m_lockEvents);

	for (std::vector<RENAME_MGR_EVENT>::iterator it = m_renameManagerEvents.begin(); it != m_renameManagerEvents.end(); ++it)
	{
		if (it->pEvents)
		{
			it->pEvents->OnItemAdded(renameItem);
		}
	}
}

void CSmartRenameManager::_OnUpdate(_In_ ISmartRenameItem* renameItem)
{
	CSRWSharedAutoLock lock(&m_lockEvents);

	for (std::vector<RENAME_MGR_EVENT>::iterator it = m_renameManagerEvents.begin(); it != m_renameManagerEvents.end(); ++it)
	{
		if (it->pEvents)
		{
			it->pEvents->OnUpdate(renameItem);
		}
	}
}

void CSmartRenameManager::_OnError(_In_ ISmartRenameItem* renameItem)
{
	CSRWSharedAutoLock lock(&m_lockEvents);

	for (std::vector<RENAME_MGR_EVENT>::iterator it = m_renameManagerEvents.begin(); it != m_renameManagerEvents.end(); ++it)
	{
		if (it->pEvents)
		{
			it->pEvents->OnError(renameItem);
		}
	}
}

void CSmartRenameManager::_OnRegExStarted(_In_ uint32_t threadId)
{
	CSRWSharedAutoLock lock(&m_lockEvents);

	for (std::vector<RENAME_MGR_EVENT>::iterator it = m_renameManagerEvents.begin(); it != m_renameManagerEvents.end(); ++it)
	{
		if (it->pEvents)
		{
			it->pEvents->OnRegExStarted(threadId);
		}
	}
}

void CSmartRenameManager::_OnRegExCanceled(_In_ uint32_t threadId)
{
	CSRWSharedAutoLock lock(&m_lockEvents);

	for (std::vector<RENAME_MGR_EVENT>::iterator it = m_renameManagerEvents.begin(); it != m_renameManagerEvents.end(); ++it)
	{
		if (it->pEvents)
		{
			it->pEvents->OnRegExCanceled(threadId);
		}
	}
}

void CSmartRenameManager::_OnRegExCompleted(_In_ uint32_t threadId)
{
	CSRWSharedAutoLock lock(&m_lockEvents);

	for (std::vector<RENAME_MGR_EVENT>::iterator it = m_renameManagerEvents.begin(); it != m_renameManagerEvents.end(); ++it)
	{
		if (it->pEvents)
		{
			it->pEvents->OnRegExCompleted(threadId);
		}
	}
}

void CSmartRenameManager::_OnRenameStarted()
{
	CSRWSharedAutoLock lock(&m_lockEvents);

	for (std::vector<RENAME_MGR_EVENT>::iterator it = m_renameManagerEvents.begin(); it != m_renameManagerEvents.end(); ++it)
	{
		if (it->pEvents)
		{
			it->pEvents->OnRenameStarted();
		}
	}
}

void CSmartRenameManager::_OnRenameCompleted()
{
	CSRWSharedAutoLock lock(&m_lockEvents);

	for (std::vector<RENAME_MGR_EVENT>::iterator it = m_renameManagerEvents.begin(); it != m_renameManagerEvents.end(); ++it)
	{
		if (it->pEvents)
		{
			it->pEvents->OnRenameCompleted();
		}
	}
}

void CSmartRenameManager::_ClearEventHandlers()
{
	CSRWExclusiveAutoLock lock(&m_lockEvents);

	// Cleanup event handlers
	for (std::vector<RENAME_MGR_EVENT>::iterator it = m_renameManagerEvents.begin(); it != m_renameManagerEvents.end(); ++it)
	{
		it->cookie = 0;
		if (it->pEvents)
		{
			it->pEvents->Release();
			it->pEvents = nullptr;
		}
	}

	m_renameManagerEvents.clear();
}

void CSmartRenameManager::_ClearSmartRenameItems()
{
	CSRWExclusiveAutoLock lock(&m_lockItems);

	// Cleanup smart rename items
	for (std::map<int, ISmartRenameItem*>::iterator it = m_renameItems.begin(); it != m_renameItems.end(); ++it)
	{
		ISmartRenameItem* pItem = it->second;
		pItem->Release();
	}

	m_renameItems.clear();
}

void CSmartRenameManager::_Cleanup()
{
	if (m_hwndMessage)
	{
		DestroyWindow(m_hwndMessage);
		m_hwndMessage = nullptr;
	}

	CloseHandle(m_startFileOpWorkerEvent);
	m_startFileOpWorkerEvent = nullptr;

	CloseHandle(m_startRegExWorkerEvent);
	m_startRegExWorkerEvent = nullptr;

	CloseHandle(m_cancelRegExWorkerEvent);
	m_cancelRegExWorkerEvent = nullptr;

	_ClearRegEx();
	_ClearEventHandlers();
	_ClearSmartRenameItems();
}
