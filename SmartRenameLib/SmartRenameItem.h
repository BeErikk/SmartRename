#pragma once

//#include "smartrename_pch.h"
//#include "SmartRenameInterfaces.h"
//#include "srwlock.h"

class CSmartRenameItem 
	: public ISmartRenameItem
	, public ISmartRenameItemFactory
{
public:
	// IUnknown
	IFACEMETHODIMP QueryInterface(_In_ REFIID iid, _Outptr_ void** resultInterface);
	IFACEMETHODIMP_(ULONG) AddRef();
	IFACEMETHODIMP_(ULONG) Release();

	// ISmartRenameItem
	IFACEMETHODIMP get_path(_Outptr_ PWSTR* path);
	IFACEMETHODIMP get_shellItem(_Outptr_ IShellItem** ppsi);
	IFACEMETHODIMP get_originalName(_Outptr_ PWSTR* originalName);
	IFACEMETHODIMP put_newName(_In_opt_ PCWSTR newName);
	IFACEMETHODIMP get_newName(_Outptr_ PWSTR* newName);
	IFACEMETHODIMP get_isFolder(_Out_ bool* isFolder);
	IFACEMETHODIMP get_isSubFolderContent(_Out_ bool* isSubFolderContent);
	IFACEMETHODIMP get_selected(_Out_ bool* selected);
	IFACEMETHODIMP put_selected(_In_ bool selected);
	IFACEMETHODIMP get_id(_Out_ int* id);
	IFACEMETHODIMP get_iconIndex(_Out_ int* iconIndex);
	IFACEMETHODIMP get_depth(_Out_ uint32_t* depth);
	IFACEMETHODIMP put_depth(_In_ int depth);
	IFACEMETHODIMP Reset();
	IFACEMETHODIMP ShouldRenameItem(_In_ uint32_t flags, _Out_ bool* shouldRename);

	// ISmartRenameItemFactory
	IFACEMETHODIMP Create(_In_ IShellItem* psi, _Outptr_ ISmartRenameItem** ppItem)
	{
		return CSmartRenameItem::s_CreateInstance(psi, IID_PPV_ARGS(ppItem));
	}

public:
	static HRESULT s_CreateInstance(_In_opt_ IShellItem* psi, _In_ REFIID iid, _Outptr_ void** resultInterface);

protected:
	static int s_id;
	CSmartRenameItem(void) noexcept
	{
		m_id = ++s_id;
	}
	virtual ~CSmartRenameItem(void) noexcept
	{
		//CoTaskMemFree(m_path);
		//CoTaskMemFree(m_newName);
		//CoTaskMemFree(m_originalName);
	}

	CSmartRenameItem(CSmartRenameItem const&) = delete;
	CSmartRenameItem& operator=(CSmartRenameItem const&) = delete;

	HRESULT _Init(_In_ IShellItem* psi);

	int m_id = -1;
	int m_iconIndex = -1;
	uint32_t m_depth = 0;
	HRESULT m_error = S_OK;
	wil::unique_cotaskmem_string m_path;
	wil::unique_cotaskmem_string m_originalName;
	wil::unique_cotaskmem_string m_newName;
	CSRWLock m_lock;
	std::atomic<long> m_refCount = 1;
	bool m_selected = true;
	bool m_isFolder = false;
	uint16_t _padding;
};
