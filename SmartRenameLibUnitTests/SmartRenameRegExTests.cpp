#include "smartrename_pch.h"
#include "common.h"

#include <CppUnitTest.h>
#include "srwlock.h"
#include "smartrenameinterfaces.h"
#include "smartrenameregex.h"
#include "MockSmartRenameRegExEvents.h"

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

namespace SmartRenameRegExTests
{
    struct SearchReplaceExpected
    {
        PCWSTR search;
        PCWSTR replace;
        PCWSTR test;
        PCWSTR expected;
    };

    TEST_CLASS(SimpleTests)
    {
    public:
        TEST_METHOD(GeneralReplaceTest)
        {
            ATL::CComPtr<ISmartRenameRegEx> renameRegEx;
            Assert::IsTrue(CSmartRenameRegEx::s_CreateInstance(&renameRegEx) == S_OK);
            PWSTR result = nullptr;
            Assert::IsTrue(renameRegEx->put_searchTerm(L"foo") == S_OK);
            Assert::IsTrue(renameRegEx->put_replaceTerm(L"big") == S_OK);
            Assert::IsTrue(renameRegEx->Replace(L"foobar", &result) == S_OK);
            Assert::IsTrue(wcscmp(result, L"bigbar") == 0);
            CoTaskMemFree(result);
        }

        TEST_METHOD(ReplaceNoMatch)
        {
            ATL::CComPtr<ISmartRenameRegEx> renameRegEx;
            Assert::IsTrue(CSmartRenameRegEx::s_CreateInstance(&renameRegEx) == S_OK);
            PWSTR result = nullptr;
            Assert::IsTrue(renameRegEx->put_searchTerm(L"notfound") == S_OK);
            Assert::IsTrue(renameRegEx->put_replaceTerm(L"big") == S_OK);
            Assert::IsTrue(renameRegEx->Replace(L"foobar", &result) == S_OK);
            Assert::IsTrue(wcscmp(result, L"foobar") == 0);
            CoTaskMemFree(result);
        }

        TEST_METHOD(ReplaceNoSearchOrReplaceTerm)
        {
            ATL::CComPtr<ISmartRenameRegEx> renameRegEx;
            Assert::IsTrue(CSmartRenameRegEx::s_CreateInstance(&renameRegEx) == S_OK);
            PWSTR result = nullptr;
            Assert::IsTrue(renameRegEx->Replace(L"foobar", &result) != S_OK);
            Assert::IsTrue(result == nullptr);
            CoTaskMemFree(result);
        }

        TEST_METHOD(ReplaceNoReplaceTerm)
        {
            ATL::CComPtr<ISmartRenameRegEx> renameRegEx;
            Assert::IsTrue(CSmartRenameRegEx::s_CreateInstance(&renameRegEx) == S_OK);
            PWSTR result = nullptr;
            Assert::IsTrue(renameRegEx->put_searchTerm(L"foo") == S_OK);
            Assert::IsTrue(renameRegEx->Replace(L"foobar", &result) == S_OK);
            Assert::IsTrue(wcscmp(result, L"bar") == 0);
            CoTaskMemFree(result);
        }

        TEST_METHOD(ReplaceEmptyStringReplaceTerm)
        {
            ATL::CComPtr<ISmartRenameRegEx> renameRegEx;
            Assert::IsTrue(CSmartRenameRegEx::s_CreateInstance(&renameRegEx) == S_OK);
            PWSTR result = nullptr;
            Assert::IsTrue(renameRegEx->put_searchTerm(L"foo") == S_OK);
            Assert::IsTrue(renameRegEx->put_replaceTerm(L"") == S_OK);
            Assert::IsTrue(renameRegEx->Replace(L"foobar", &result) == S_OK);
            Assert::IsTrue(wcscmp(result, L"bar") == 0);
            CoTaskMemFree(result);
        }

        TEST_METHOD(VerifyDefaultFlags)
        {
            ATL::CComPtr<ISmartRenameRegEx> renameRegEx;
            Assert::IsTrue(CSmartRenameRegEx::s_CreateInstance(&renameRegEx) == S_OK);
            DWORD flags = 0;
            Assert::IsTrue(renameRegEx->get_flags(&flags) == S_OK);
            Assert::IsTrue(flags == MatchAllOccurrences);
        }

        TEST_METHOD(VerifyCaseSensitiveSearch)
        {
            ATL::CComPtr<ISmartRenameRegEx> renameRegEx;
            Assert::IsTrue(CSmartRenameRegEx::s_CreateInstance(&renameRegEx) == S_OK);
            DWORD flags = CaseSensitive;
            Assert::IsTrue(renameRegEx->put_flags(flags) == S_OK);

            SearchReplaceExpected sreTable[] =
            {
                { L"Foo", L"Foo", L"FooBar", L"FooBar" },
                { L"Foo", L"boo", L"FooBar", L"booBar" },
                { L"Foo", L"boo", L"foobar", L"foobar" },
                { L"123", L"654", L"123456", L"654456" },
            };

            for (int i = 0; i < ARRAYSIZE(sreTable); i++)
            {
                PWSTR result = nullptr;
                Assert::IsTrue(renameRegEx->put_searchTerm(sreTable[i].search) == S_OK);
                Assert::IsTrue(renameRegEx->put_replaceTerm(sreTable[i].replace) == S_OK);
                Assert::IsTrue(renameRegEx->Replace(sreTable[i].test, &result) == S_OK);
                Assert::IsTrue(wcscmp(result, sreTable[i].expected) == 0);
                CoTaskMemFree(result);
            }
        }

        TEST_METHOD(VerifyReplaceFirstOnly)
        {
            ATL::CComPtr<ISmartRenameRegEx> renameRegEx;
            Assert::IsTrue(CSmartRenameRegEx::s_CreateInstance(&renameRegEx) == S_OK);
            DWORD flags = 0;
            Assert::IsTrue(renameRegEx->put_flags(flags) == S_OK);

            SearchReplaceExpected sreTable[] =
            {
                { L"B", L"BB", L"ABA", L"ABBA" },
                { L"B", L"A", L"ABBBA", L"AABBA" },
                { L"B", L"BBB", L"ABABAB", L"ABBBABAB" },
            };

            for (int i = 0; i < ARRAYSIZE(sreTable); i++)
            {
                PWSTR result = nullptr;
                Assert::IsTrue(renameRegEx->put_searchTerm(sreTable[i].search) == S_OK);
                Assert::IsTrue(renameRegEx->put_replaceTerm(sreTable[i].replace) == S_OK);
                Assert::IsTrue(renameRegEx->Replace(sreTable[i].test, &result) == S_OK);
                Assert::IsTrue(wcscmp(result, sreTable[i].expected) == 0);
                CoTaskMemFree(result);
            }
        }

        TEST_METHOD(VerifyReplaceAll)
        {
            ATL::CComPtr<ISmartRenameRegEx> renameRegEx;
            Assert::IsTrue(CSmartRenameRegEx::s_CreateInstance(&renameRegEx) == S_OK);
            DWORD flags = MatchAllOccurrences;
            Assert::IsTrue(renameRegEx->put_flags(flags) == S_OK);

            SearchReplaceExpected sreTable[] =
            {
                { L"B", L"BB", L"ABA", L"ABBA" },
                { L"B", L"A", L"ABBBA", L"AAAAA" },
                { L"B", L"BBB", L"ABABAB", L"ABBBABBBABBB" },
            };

            for (int i = 0; i < ARRAYSIZE(sreTable); i++)
            {
                PWSTR result = nullptr;
                Assert::IsTrue(renameRegEx->put_searchTerm(sreTable[i].search) == S_OK);
                Assert::IsTrue(renameRegEx->put_replaceTerm(sreTable[i].replace) == S_OK);
                Assert::IsTrue(renameRegEx->Replace(sreTable[i].test, &result) == S_OK);
                Assert::IsTrue(wcscmp(result, sreTable[i].expected) == 0);
                CoTaskMemFree(result);
            }
        }

        TEST_METHOD(VerifyReplaceAllCaseInsensitive)
        {
            ATL::CComPtr<ISmartRenameRegEx> renameRegEx;
            Assert::IsTrue(CSmartRenameRegEx::s_CreateInstance(&renameRegEx) == S_OK);
            DWORD flags = MatchAllOccurrences | CaseSensitive;
            Assert::IsTrue(renameRegEx->put_flags(flags) == S_OK);

            SearchReplaceExpected sreTable[] =
            {
                { L"B", L"BB", L"ABA", L"ABBA" },
                { L"B", L"A", L"ABBBA", L"AAAAA" },
                { L"B", L"BBB", L"ABABAB", L"ABBBABBBABBB" },
                { L"b", L"BBB", L"AbABAb", L"ABBBABABBB" },
            };

            for (int i = 0; i < ARRAYSIZE(sreTable); i++)
            {
                PWSTR result = nullptr;
                Assert::IsTrue(renameRegEx->put_searchTerm(sreTable[i].search) == S_OK);
                Assert::IsTrue(renameRegEx->put_replaceTerm(sreTable[i].replace) == S_OK);
                Assert::IsTrue(renameRegEx->Replace(sreTable[i].test, &result) == S_OK);
                Assert::IsTrue(wcscmp(result, sreTable[i].expected) == 0);
                CoTaskMemFree(result);
            }
        }
    };

    TEST_CLASS(RegExEventsTests)
    {
    public:
        TEST_METHOD(VerifyEventsFire)
        {
            ATL::CComPtr<ISmartRenameRegEx> renameRegEx;
            Assert::IsTrue(CSmartRenameRegEx::s_CreateInstance(&renameRegEx) == S_OK);
            CMockSmartRenameRegExEvents* mockEvents = new CMockSmartRenameRegExEvents();
            ATL::CComPtr<ISmartRenameRegExEvents> regExEvents;
            Assert::IsTrue(mockEvents->QueryInterface(IID_PPV_ARGS(&regExEvents)) == S_OK);
            DWORD cookie = 0;
            Assert::IsTrue(renameRegEx->Advise(regExEvents, &cookie) == S_OK);
            DWORD flags = MatchAllOccurrences | UseRegularExpressions | CaseSensitive;
            Assert::IsTrue(renameRegEx->put_flags(flags) == S_OK);
            Assert::IsTrue(renameRegEx->put_searchTerm(L"FOO") == S_OK);
            Assert::IsTrue(renameRegEx->put_replaceTerm(L"BAR") == S_OK);
            Assert::IsTrue(lstrcmpi(L"FOO", mockEvents->m_searchTerm) == 0);
            Assert::IsTrue(lstrcmpi(L"BAR", mockEvents->m_replaceTerm) == 0);
            Assert::IsTrue(flags == mockEvents->m_flags);
            Assert::IsTrue(renameRegEx->UnAdvise(cookie) == S_OK);
            mockEvents->Release();
        }
    };

    TEST_CLASS(RegExTests)
    {
    public:
        TEST_METHOD(VerifyReplaceFirstOnlyUseRegEx)
        {
            ATL::CComPtr<ISmartRenameRegEx> renameRegEx;
            Assert::IsTrue(CSmartRenameRegEx::s_CreateInstance(&renameRegEx) == S_OK);
            DWORD flags = UseRegularExpressions;
            Assert::IsTrue(renameRegEx->put_flags(flags) == S_OK);

            SearchReplaceExpected sreTable[] =
            {
                { L"B", L"BB", L"ABA", L"ABBA" },
                { L"B", L"A", L"ABBBA", L"AABBA" },
                { L"B", L"BBB", L"ABABAB", L"ABBBABAB" },
            };

            for (int i = 0; i < ARRAYSIZE(sreTable); i++)
            {
                PWSTR result = nullptr;
                Assert::IsTrue(renameRegEx->put_searchTerm(sreTable[i].search) == S_OK);
                Assert::IsTrue(renameRegEx->put_replaceTerm(sreTable[i].replace) == S_OK);
                Assert::IsTrue(renameRegEx->Replace(sreTable[i].test, &result) == S_OK);
                Assert::IsTrue(wcscmp(result, sreTable[i].expected) == 0);
                CoTaskMemFree(result);
            }
        }

        TEST_METHOD(VerifyReplaceAllUseRegEx)
        {
            ATL::CComPtr<ISmartRenameRegEx> renameRegEx;
            Assert::IsTrue(CSmartRenameRegEx::s_CreateInstance(&renameRegEx) == S_OK);
            DWORD flags = MatchAllOccurrences | UseRegularExpressions;
            Assert::IsTrue(renameRegEx->put_flags(flags) == S_OK);

            SearchReplaceExpected sreTable[] =
            {
                { L"B", L"BB", L"ABA", L"ABBA" },
                { L"B", L"A", L"ABBBA", L"AAAAA" },
                { L"B", L"BBB", L"ABABAB", L"ABBBABBBABBB" },
            };

            for (int i = 0; i < ARRAYSIZE(sreTable); i++)
            {
                PWSTR result = nullptr;
                Assert::IsTrue(renameRegEx->put_searchTerm(sreTable[i].search) == S_OK);
                Assert::IsTrue(renameRegEx->put_replaceTerm(sreTable[i].replace) == S_OK);
                Assert::IsTrue(renameRegEx->Replace(sreTable[i].test, &result) == S_OK);
                Assert::IsTrue(wcscmp(result, sreTable[i].expected) == 0);
                CoTaskMemFree(result);
            }
        }

        TEST_METHOD(VerifyReplaceAllUseRegExCaseSensitive)
        {
            ATL::CComPtr<ISmartRenameRegEx> renameRegEx;
            Assert::IsTrue(CSmartRenameRegEx::s_CreateInstance(&renameRegEx) == S_OK);
            DWORD flags = MatchAllOccurrences | UseRegularExpressions | CaseSensitive;
            Assert::IsTrue(renameRegEx->put_flags(flags) == S_OK);

            SearchReplaceExpected sreTable[] =
            {
                { L"B", L"BB", L"ABA", L"ABBA" },
                { L"B", L"A", L"ABBBA", L"AAAAA" },
                { L"b", L"BBB", L"AbABAb", L"ABBBABABBB" },
            };

            for (int i = 0; i < ARRAYSIZE(sreTable); i++)
            {
                PWSTR result = nullptr;
                Assert::IsTrue(renameRegEx->put_searchTerm(sreTable[i].search) == S_OK);
                Assert::IsTrue(renameRegEx->put_replaceTerm(sreTable[i].replace) == S_OK);
                Assert::IsTrue(renameRegEx->Replace(sreTable[i].test, &result) == S_OK);
                Assert::IsTrue(wcscmp(result, sreTable[i].expected) == 0);
                CoTaskMemFree(result);
            }
        }

        TEST_METHOD(VerifyMatchAllWildcardUseRegEx)
        {
            ATL::CComPtr<ISmartRenameRegEx> renameRegEx;
            Assert::IsTrue(CSmartRenameRegEx::s_CreateInstance(&renameRegEx) == S_OK);
            DWORD flags = MatchAllOccurrences | UseRegularExpressions;
            Assert::IsTrue(renameRegEx->put_flags(flags) == S_OK);

            SearchReplaceExpected sreTable[] =
            {
                { L".*", L"Foo", L"AAAAAA", L"Foo" },
            };

            for (int i = 0; i < ARRAYSIZE(sreTable); i++)
            {
                PWSTR result = nullptr;
                Assert::IsTrue(renameRegEx->put_searchTerm(sreTable[i].search) == S_OK);
                Assert::IsTrue(renameRegEx->put_replaceTerm(sreTable[i].replace) == S_OK);
                Assert::IsTrue(renameRegEx->Replace(sreTable[i].test, &result) == S_OK);
                Assert::IsTrue(wcscmp(result, sreTable[i].expected) == 0);
                CoTaskMemFree(result);
            }
        }

        TEST_METHOD(VerifyReplaceFirstWildcardUseRegEx)
        {
            ATL::CComPtr<ISmartRenameRegEx> renameRegEx;
            Assert::IsTrue(CSmartRenameRegEx::s_CreateInstance(&renameRegEx) == S_OK);
            DWORD flags = UseRegularExpressions;
            Assert::IsTrue(renameRegEx->put_flags(flags) == S_OK);

            SearchReplaceExpected sreTable[] =
            {
                { L".", L"Foo", L"AAAAAA", L"FooAAAAA" },
            };

            for (int i = 0; i < ARRAYSIZE(sreTable); i++)
            {
                PWSTR result = nullptr;
                Assert::IsTrue(renameRegEx->put_searchTerm(sreTable[i].search) == S_OK);
                Assert::IsTrue(renameRegEx->put_replaceTerm(sreTable[i].replace) == S_OK);
                Assert::IsTrue(renameRegEx->Replace(sreTable[i].test, &result) == S_OK);
                Assert::IsTrue(wcscmp(result, sreTable[i].expected) == 0);
                CoTaskMemFree(result);
            }
        }

        TEST_METHOD(VerifyMatchBeginningTextUseRegEx)
        {
            ATL::CComPtr<ISmartRenameRegEx> renameRegEx;
            Assert::IsTrue(CSmartRenameRegEx::s_CreateInstance(&renameRegEx) == S_OK);
            DWORD flags = UseRegularExpressions;
            Assert::IsTrue(renameRegEx->put_flags(flags) == S_OK);

            SearchReplaceExpected sreTable[] =
            {
                { L"^Foo", L"Baa", L"FooBar", L"BaaBar" },
                { L"^Foo", L"Baa", L"BarFoo", L"BarFoo" },
                { L"^Foo", L"Baa", L"FooBarFoo", L"BaaBarFoo" },
            };

            for (int i = 0; i < ARRAYSIZE(sreTable); i++)
            {
                PWSTR result = nullptr;
                Assert::IsTrue(renameRegEx->put_searchTerm(sreTable[i].search) == S_OK);
                Assert::IsTrue(renameRegEx->put_replaceTerm(sreTable[i].replace) == S_OK);
                Assert::IsTrue(renameRegEx->Replace(sreTable[i].test, &result) == S_OK);
                Assert::IsTrue(wcscmp(result, sreTable[i].expected) == 0);
                CoTaskMemFree(result);
            }
        }

        TEST_METHOD(VerifyMatchEndingTextUseRegEx)
        {
            ATL::CComPtr<ISmartRenameRegEx> renameRegEx;
            Assert::IsTrue(CSmartRenameRegEx::s_CreateInstance(&renameRegEx) == S_OK);
            DWORD flags = UseRegularExpressions;
            Assert::IsTrue(renameRegEx->put_flags(flags) == S_OK);

            SearchReplaceExpected sreTable[] =
            {
                { L"Bar$", L"Baa", L"FooBar", L"FooBaa" },
                { L"Bar$", L"Baa", L"BarFoo", L"BarFoo" },
                { L"Bar$", L"Baa", L"BarFooBar", L"BarFooBaa" },
            };

            for (int i = 0; i < ARRAYSIZE(sreTable); i++)
            {
                PWSTR result = nullptr;
                Assert::IsTrue(renameRegEx->put_searchTerm(sreTable[i].search) == S_OK);
                Assert::IsTrue(renameRegEx->put_replaceTerm(sreTable[i].replace) == S_OK);
                Assert::IsTrue(renameRegEx->Replace(sreTable[i].test, &result) == S_OK);
                Assert::IsTrue(wcscmp(result, sreTable[i].expected) == 0);
                CoTaskMemFree(result);
            }
        }
    };
}
