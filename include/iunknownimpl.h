
template <QITAB* Tqit>
class IUnknownImpl
	: public IUnknown
{
public:
		IFACEMETHODIMP QueryInterface(_In_ REFIID riid, _COM_Outptr_ void** ppv)
		{
			return ::QISearch(this, Tqit, riid, ppv);
		}
		IFACEMETHODIMP_(ULONG) AddRef(void)
		{
			return ++m_refCount;
		}
		IFACEMETHODIMP_(ULONG) Release(void)
		{
			auto refCount = --m_refCount;
			if (refCount == 0)
			{
				delete this;
			}
			return refCount;
		}

private:
	std::atomic<ULONG> m_refCount = 1u;
};
