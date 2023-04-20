#pragma once

//#include "smartrename_pch.h"
//#include "smartrenameinterfaces.h"
//#include <vector>
//#include "srwlock.h"

class CSmartRenameEnum :
    public ISmartRenameEnum
{
public:
    // IUnknown
    IFACEMETHODIMP  QueryInterface(_In_ REFIID iid, _Outptr_ void** resultInterface);
    IFACEMETHODIMP_(ULONG) AddRef();
    IFACEMETHODIMP_(ULONG) Release();

    // ISmartRenameEnum
    IFACEMETHODIMP Start();
    IFACEMETHODIMP Cancel();

public:
    static HRESULT s_CreateInstance(_In_ IUnknown* punk, _In_ ISmartRenameManager* psrm, _In_ REFIID iid, _Outptr_ void** resultInterface);

protected:
    CSmartRenameEnum();
    virtual ~CSmartRenameEnum();

    HRESULT _Init(_In_ IUnknown* punk, _In_ ISmartRenameManager* psrm);
    HRESULT _ParseEnumItems(_In_ IEnumShellItems* pesi, _In_ int depth = 0);

    ATL::CComPtr<ISmartRenameManager> m_spsrm;
    ATL::CComPtr<IUnknown> m_spunk;
    bool m_canceled = false;
    long m_refCount = 0;
};
