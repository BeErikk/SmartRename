#pragma once

//#include "smartrename_pch.h"

enum SmartRenameFlags
{
	CaseSensitive = 0x1,
	MatchAllOccurrences = 0x2,
	UseRegularExpressions = 0x4,
	EnumerateItems = 0x8,
	ExcludeFiles = 0x10,
	ExcludeFolders = 0x20,
	ExcludeSubfolders = 0x40,
	NameOnly = 0x80,
	ExtensionOnly = 0x100
};

interface __declspec(novtable) __declspec(uuid("3ECBA62B-E0F0-4472-AA2E-DEE7A1AA46B9")) ISmartRenameRegExEvents : public IUnknown
{
public:
	IFACEMETHOD(OnSearchTermChanged)
	(_In_ PCWSTR searchTerm) = 0;
	IFACEMETHOD(OnReplaceTermChanged)
	(_In_ PCWSTR replaceTerm) = 0;
	IFACEMETHOD(OnFlagsChanged)
	(_In_ uint32_t flags) = 0;
};

interface __declspec(novtable) __declspec(uuid("E3ED45B5-9CE0-47E2-A595-67EB950B9B72")) ISmartRenameRegEx : public IUnknown
{
public:
	IFACEMETHOD(Advise)
	(_In_ ISmartRenameRegExEvents* regExEvents, _Out_ uint32_t* cookie) = 0;
	IFACEMETHOD(UnAdvise)
	(_In_ uint32_t cookie) = 0;
	IFACEMETHOD(get_searchTerm)
	(_Outptr_ PWSTR* searchTerm) = 0;
	IFACEMETHOD(put_searchTerm)
	(_In_ PCWSTR searchTerm) = 0;
	IFACEMETHOD(get_replaceTerm)
	(_Outptr_ PWSTR* replaceTerm) = 0;
	IFACEMETHOD(put_replaceTerm)
	(_In_ PCWSTR replaceTerm) = 0;
	IFACEMETHOD(get_flags)
	(_Out_ uint32_t* flags) = 0;
	IFACEMETHOD(put_flags)
	(_In_ uint32_t flags) = 0;
	IFACEMETHOD(Replace)
	(_In_ PCWSTR source, _Outptr_ PWSTR* result) = 0;
};

interface __declspec(novtable) __declspec(uuid("C7F59201-4DE1-4855-A3A2-26FC3279C8A5")) ISmartRenameItem : public IUnknown
{
public:
	IFACEMETHOD(get_path)
	(_Outptr_ PWSTR* path) = 0;
	IFACEMETHOD(get_shellItem)
	(_Outptr_ IShellItem** ppsi) = 0;
	IFACEMETHOD(get_originalName)
	(_Outptr_ PWSTR* originalName) = 0;
	IFACEMETHOD(get_newName)
	(_Outptr_ PWSTR* newName) = 0;
	IFACEMETHOD(put_newName)
	(_In_opt_ PCWSTR newName) = 0;
	IFACEMETHOD(get_isFolder)
	(_Out_ bool* isFolder) = 0;
	IFACEMETHOD(get_isSubFolderContent)
	(_Out_ bool* isSubFolderContent) = 0;
	IFACEMETHOD(get_selected)
	(_Out_ bool* selected) = 0;
	IFACEMETHOD(put_selected)
	(_In_ bool selected) = 0;
	IFACEMETHOD(get_id)
	(_Out_ int* id) = 0;
	IFACEMETHOD(get_iconIndex)
	(_Out_ int* iconIndex) = 0;
	IFACEMETHOD(get_depth)
	(_Out_ uint32_t* depth) = 0;
	IFACEMETHOD(put_depth)
	(_In_ int depth) = 0;
	IFACEMETHOD(ShouldRenameItem)
	(_In_ uint32_t flags, _Out_ bool* shouldRename) = 0;
	IFACEMETHOD(Reset)
	() = 0;
};

interface __declspec(novtable) __declspec(uuid("{26CBFFD9-13B3-424E-BAC9-D12B0539149C}")) ISmartRenameItemFactory : public IUnknown
{
public:
	IFACEMETHOD(Create)
	(_In_ IShellItem* psi, _COM_Outptr_ ISmartRenameItem** ppItem) = 0;
};

interface __declspec(novtable) __declspec(uuid("87FC43F9-7634-43D9-99A5-20876AFCE4AD")) ISmartRenameManagerEvents : public IUnknown
{
public:
	IFACEMETHOD(OnItemAdded)
	(_In_ ISmartRenameItem* renameItem) = 0;
	IFACEMETHOD(OnUpdate)
	(_In_ ISmartRenameItem* renameItem) = 0;
	IFACEMETHOD(OnError)
	(_In_ ISmartRenameItem* renameItem) = 0;
	IFACEMETHOD(OnRegExStarted)
	(_In_ uint32_t threadId) = 0;
	IFACEMETHOD(OnRegExCanceled)
	(_In_ uint32_t threadId) = 0;
	IFACEMETHOD(OnRegExCompleted)
	(_In_ uint32_t threadId) = 0;
	IFACEMETHOD(OnRenameStarted)
	() = 0;
	IFACEMETHOD(OnRenameCompleted)
	() = 0;
};

interface __declspec(novtable) __declspec(uuid("001BBD88-53D2-4FA6-95D2-F9A9FA4F9F70")) ISmartRenameManager : public IUnknown
{
public:
	IFACEMETHOD(Advise)
	(_In_ ISmartRenameManagerEvents* renameManagerEvent, _Out_ uint32_t* cookie) = 0;
	IFACEMETHOD(UnAdvise)
	(_In_ uint32_t cookie) = 0;
	IFACEMETHOD(Start)
	() = 0;
	IFACEMETHOD(Stop)
	() = 0;
	IFACEMETHOD(Reset)
	() = 0;
	IFACEMETHOD(Shutdown)
	() = 0;
	IFACEMETHOD(Rename)
	(_In_ HWND hwndParent) = 0;
	IFACEMETHOD(AddItem)
	(_In_ ISmartRenameItem* pItem) = 0;
	IFACEMETHOD(GetItemByIndex)
	(_In_ uint32_t index, _COM_Outptr_ ISmartRenameItem** ppItem) = 0;
	IFACEMETHOD(GetItemById)
	(_In_ int id, _COM_Outptr_ ISmartRenameItem** ppItem) = 0;
	IFACEMETHOD(GetItemCount)
	(_Out_ uint32_t* count) = 0;
	IFACEMETHOD(GetSelectedItemCount)
	(_Out_ uint32_t* count) = 0;
	IFACEMETHOD(GetRenameItemCount)
	(_Out_ uint32_t* count) = 0;
	IFACEMETHOD(get_flags)
	(_Out_ uint32_t* flags) = 0;
	IFACEMETHOD(put_flags)
	(_In_ uint32_t flags) = 0;
	IFACEMETHOD(get_renameRegEx)
	(_COM_Outptr_ ISmartRenameRegEx** ppRegEx) = 0;
	IFACEMETHOD(put_renameRegEx)
	(_In_ ISmartRenameRegEx* pRegEx) = 0;
	IFACEMETHOD(get_renameItemFactory)
	(_COM_Outptr_ ISmartRenameItemFactory** ppItemFactory) = 0;
	IFACEMETHOD(put_renameItemFactory)
	(_In_ ISmartRenameItemFactory* pItemFactory) = 0;
};

interface __declspec(novtable) __declspec(uuid("E6679DEB-460D-42C1-A7A8-E25897061C99")) ISmartRenameUI : public IUnknown
{
public:
	IFACEMETHOD(Show)
	(_In_opt_ HWND hwndParent) = 0;
	IFACEMETHOD(Close)
	() = 0;
	IFACEMETHOD(Update)
	() = 0;
};

interface __declspec(novtable) __declspec(uuid("04AAFABE-B76E-4E13-993A-B5941F52B139")) ISmartRenameMRU : public IUnknown
{
public:
	IFACEMETHOD(AddMRUString)
	(_In_ PCWSTR entry) = 0;
};

interface __declspec(novtable) __declspec(uuid("2EFBAB41-A841-47B5-898B-B1CFBF151855")) ISmartRenameEnum : public IUnknown
{
public:
	IFACEMETHOD(Start)
	() = 0;
	IFACEMETHOD(Cancel)
	() = 0;
};
