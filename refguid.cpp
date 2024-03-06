// refguid - Generating/Converting/Searching GUIDs
// License: MIT

#include <windows.h>
#include <shlobj.h>
#include <shlwapi.h>
#include <initguid.h>

#include <string>
#include <vector>
#include <cstdio>
#include <cassert>
#include <strsafe.h>

void show_version(void)
{
    std::puts("refguid version 1.2 by katahiromz");
}

void usage(void)
{
    std::puts("Usage: refguid guid_name\n"
              "       refguid \"guid_string\"\n"
              "       refguid --search \"search_string\"\n"
              "       refguid --list\n"
              "       refguid --generate");
}

template <typename T_CHAR>
inline void mstr_trim(std::basic_string<T_CHAR>& str, const T_CHAR* spaces)
{
    typedef std::basic_string<T_CHAR> string_type;
    size_t i = str.find_first_not_of(spaces);
    size_t j = str.find_last_not_of(spaces);
    if ((i == string_type::npos) || (j == string_type::npos))
    {
        str.clear();
    }
    else
    {
        str = str.substr(i, j - i + 1);
    }
}

template <typename T_STR>
inline bool
mstr_replace_all(T_STR& str, const T_STR& from, const T_STR& to)
{
    bool ret = false;
    size_t i = 0;
    for (;;) {
        i = str.find(from, i);
        if (i == T_STR::npos)
            break;
        ret = true;
        str.replace(i, from.size(), to);
        i += to.size();
    }
    return ret;
}

template <typename T_STR>
inline bool
mstr_replace_all(T_STR& str,
                 const typename T_STR::value_type *from,
                 const typename T_STR::value_type *to)
{
    return mstr_replace_all(str, T_STR(from), T_STR(to));
}

template <typename T_STR_CONTAINER>
inline void
mstr_split(T_STR_CONTAINER& container,
    const typename T_STR_CONTAINER::value_type& str,
    const typename T_STR_CONTAINER::value_type& chars)
{
    container.clear();
    size_t i = 0, k = str.find_first_of(chars);
    while (k != T_STR_CONTAINER::value_type::npos)
    {
        container.push_back(str.substr(i, k - i));
        i = k + 1;
        k = str.find_first_of(chars, i);
    }
    container.push_back(str.substr(i));
}

std::wstring BytesTextFromCLSID(REFCLSID clsid)
{
    std::wstring ret;
    const BYTE *pb = (LPBYTE)&clsid;
    for (size_t ib = 0; ib < sizeof(clsid); ++ib)
    {
        if (ib)
            ret += L' ';
        WCHAR sz[3];
        StringCchPrintfW(sz, _countof(sz), L"%02X", pb[ib]);
        ret += sz;
    }
    return ret;
}

BOOL getGUIDFromStructStr(GUID& guid, std::wstring arg)
{
    mstr_trim(arg, L" \t");
    if (arg.size() && arg[arg.size() - 1] == L';')
        arg.erase(arg.size() - 1, 1);
    mstr_trim(arg, L" \t");
    if (arg.size() && arg[0] == L'{')
        arg.erase(0, 1);
    mstr_trim(arg, L" \t");
    if (arg.size() && arg[arg.size() - 1] == L'}')
        arg.erase(arg.size() - 1, 1);
    mstr_trim(arg, L" \t");
    if (arg.size() && arg[arg.size() - 1] == L'}')
        arg.erase(arg.size() - 1, 1);

    mstr_replace_all(arg, L" ", L"");
    mstr_replace_all(arg, L"\t", L"");

    auto ich0 = arg.find(L',');
    if (ich0 != arg.npos)
    {
        auto ich1 = arg.find(L',', ich0 + 1);
        if (ich1 != arg.npos)
        {
            auto ich2 = arg.find(L',', ich1 + 1);
            if (ich2 == arg.npos || arg[ich2 + 1] != L'{')
                return FALSE;

            arg.erase(ich2 + 1, 1);
        }
        else
        {
            return FALSE;
        }
    }
    else
    {
        return FALSE;
    }

    std::vector<std::wstring> items;
    mstr_split(items, arg, L",");

    if (items.size() != 11)
        return FALSE;

    for (auto& item : items)
    {
        mstr_trim(item, L" \t");
    }

    guid.Data1 = wcstoul(items[0].c_str(), NULL, 0);
    guid.Data2 = (USHORT)wcstoul(items[1].c_str(), NULL, 0);
    guid.Data3 = (USHORT)wcstoul(items[2].c_str(), NULL, 0);
    guid.Data4[0] = (BYTE)wcstoul(items[3].c_str(), NULL, 0);
    guid.Data4[1] = (BYTE)wcstoul(items[4].c_str(), NULL, 0);
    guid.Data4[2] = (BYTE)wcstoul(items[5].c_str(), NULL, 0);
    guid.Data4[3] = (BYTE)wcstoul(items[6].c_str(), NULL, 0);
    guid.Data4[4] = (BYTE)wcstoul(items[7].c_str(), NULL, 0);
    guid.Data4[5] = (BYTE)wcstoul(items[8].c_str(), NULL, 0);
    guid.Data4[6] = (BYTE)wcstoul(items[9].c_str(), NULL, 0);
    guid.Data4[7] = (BYTE)wcstoul(items[10].c_str(), NULL, 0);

    return TRUE;
}

BOOL getGUIDFromDefineGUID(GUID& guid, std::wstring arg, std::wstring *pstrName)
{
    mstr_trim(arg, L" \t");
    if (arg.find(L"DEFINE_GUID") == 0)
        arg.erase(0, 11);
    mstr_trim(arg, L" \t");
    if (arg.size() && arg[0] == L'(')
        arg.erase(0, 1);
    mstr_trim(arg, L" \t");
    if (arg.size() && arg[arg.size() - 1] == L';')
        arg.erase(arg.size() - 1, 1);
    mstr_trim(arg, L" \t");
    if (arg.size() && arg[arg.size() - 1] == L')')
        arg.erase(arg.size() - 1, 1);
    mstr_trim(arg, L" \t");
    mstr_replace_all(arg, L" ", L"");
    mstr_replace_all(arg, L"\t", L"");

    std::vector<std::wstring> items;
    mstr_split(items, arg, L",");

    if (items.size() != 12)
        return FALSE;

    for (auto& item : items)
    {
        mstr_trim(item, L" \t");
    }

    if (pstrName)
        *pstrName = items[0];

    guid.Data1 = wcstoul(items[1].c_str(), NULL, 0);
    guid.Data2 = (USHORT)wcstoul(items[2].c_str(), NULL, 0);
    guid.Data3 = (USHORT)wcstoul(items[3].c_str(), NULL, 0);
    guid.Data4[0] = (BYTE)wcstoul(items[4].c_str(), NULL, 0);
    guid.Data4[1] = (BYTE)wcstoul(items[5].c_str(), NULL, 0);
    guid.Data4[2] = (BYTE)wcstoul(items[6].c_str(), NULL, 0);
    guid.Data4[3] = (BYTE)wcstoul(items[7].c_str(), NULL, 0);
    guid.Data4[4] = (BYTE)wcstoul(items[8].c_str(), NULL, 0);
    guid.Data4[5] = (BYTE)wcstoul(items[9].c_str(), NULL, 0);
    guid.Data4[6] = (BYTE)wcstoul(items[10].c_str(), NULL, 0);
    guid.Data4[7] = (BYTE)wcstoul(items[11].c_str(), NULL, 0);

    return TRUE;
}

BOOL getGUIDFromBytesText(GUID& guid, std::wstring arg)
{
    mstr_replace_all(arg, L"0x", L"");

    CLSID ret;
    std::vector<BYTE> bytes;
    WCHAR sz[3];
    size_t ich = 0, ib = 0;
    for (auto ch : arg)
    {
        if (!iswxdigit(ch))
            continue;
        if ((ich % 2) == 0)
        {
            sz[0] = ch;
        }
        else
        {
            sz[1] = ch;
            sz[2] = 0;
            BYTE byte = (BYTE)wcstoul(sz, NULL, 16);
            bytes.push_back(byte);
            ib++;
        }
        ich++;
    }

    if (ib != sizeof(CLSID))
        return FALSE;

    CopyMemory(&guid, bytes.data(), sizeof(ret));
    return TRUE;
}

std::wstring getDefineGUIDFromGUID(REFGUID guid, LPCWSTR name)
{
    WCHAR sz[256];
    StringCchPrintfW(sz, _countof(sz),
        L"DEFINE_GUID(%ls, 0x%08X, 0x%04X, 0x%04X, 0x%02X, 0x%02X, 0x%02X, "
        L"0x%02X, 0x%02X, 0x%02X, 0x%02X, 0x%02X);", (name && name[0] ? name : L"<Name>"),
        guid.Data1, guid.Data2, guid.Data3,
        guid.Data4[0], guid.Data4[1], guid.Data4[2], guid.Data4[3],
        guid.Data4[4], guid.Data4[5], guid.Data4[6], guid.Data4[7]);
    return sz;
}

std::wstring getStructFromGUID(REFGUID guid)
{
    WCHAR sz[256];
    StringCchPrintfW(sz, _countof(sz),
        L"{ 0x%08X, 0x%04X, 0x%04X, { 0x%02X, 0x%02X, 0x%02X, "
        L"0x%02X, 0x%02X, 0x%02X, 0x%02X, 0x%02X } }",
        guid.Data1, guid.Data2, guid.Data3,
        guid.Data4[0], guid.Data4[1], guid.Data4[2], guid.Data4[3],
        guid.Data4[4], guid.Data4[5], guid.Data4[6], guid.Data4[7]);
    return sz;
}

std::wstring getAnalysisStringFromGUID(REFCLSID guid, LPCWSTR name)
{
    std::wstring ret;
    if (name[0] == 0)
        name = NULL;

    ret += L"#include <initguid.h>\n\n";

    auto define_guid = getDefineGUIDFromGUID(guid, name);
    ret += define_guid;
    ret += L"\n\n";

    WCHAR szCLSID[40];
    StringFromGUID2(guid, szCLSID, _countof(szCLSID));
    ret += L"GUID: ";
    ret += szCLSID;
    ret += L"\n\n";

    auto strBytes = BytesTextFromCLSID(guid);
    ret += L"Bytes: ";
    ret += strBytes;
    ret += L"\n\n";

    auto struct_str = getStructFromGUID(guid);
    ret += L"Struct: ";
    ret += struct_str;
    ret += L"\n";

    return ret;
}

template <typename T_DATA>
void RandomData(T_DATA& data)
{
    std::srand(::GetTickCount());
    LPBYTE pb = (LPBYTE)&data;
    for (size_t ib = 0; ib < sizeof(data); ++ib)
    {
        pb[ib] = BYTE(std::rand());
    }
}

typedef struct tagENTRY
{
    std::wstring name;
    GUID guid;
} ENTRY, *PENTRY;

BOOL readAllGUIDs(std::vector<ENTRY>& entries)
{
    WCHAR szPath[MAX_PATH];
    GetModuleFileNameW(NULL, szPath, _countof(szPath));
    PathRemoveFileSpecW(szPath);
    PathAppendW(szPath, L"refguid.dat");

    FILE *fp = _wfopen(szPath, L"r");
    if (!fp)
    {
        fprintf(stderr, "WARNING: Cannot load %ls\n", szPath);
        return FALSE;
    }

    CHAR buf[1024];
    WCHAR szLine[1024];
    while (fgets(buf, sizeof(buf), fp))
    {
        std::string str = buf;
        mstr_trim(str, " \t\r\n");

        MultiByteToWideChar(CP_ACP, 0, buf, -1, szLine, _countof(szLine));
        szLine[_countof(szLine) - 1] = 0;

        GUID guid;
        std::wstring name;
        if (getGUIDFromDefineGUID(guid, szLine, &name))
        {
            ENTRY entry = { name, guid };
            entries.push_back(entry);
        }
    }

    fclose(fp);

    return TRUE;
}

BOOL getByName(ENTRY& found, const std::vector<ENTRY>& entries, LPCWSTR pszName)
{
    std::vector<ENTRY> got;
    for (auto& entry : entries)
    {
        if (lstrcmpiW(entry.name.c_str(), pszName) == 0)
        {
            found = entry;
            return TRUE;
        }
    }

    return FALSE;
}

BOOL doFilterByGuid(std::vector<ENTRY>& entries, const GUID& guid)
{
    std::vector<ENTRY> got;
    for (auto& entry : entries)
    {
        if (IsEqualGUID(entry.guid, guid))
        {
            got.push_back(entry);
        }
    }

    entries = std::move(got);
    return !entries.empty();
}

BOOL doFilterByString(std::vector<ENTRY>& entries, std::wstring text)
{
    ::CharUpperW(&text[0]);

    std::vector<ENTRY> got;
    for (auto& entry : entries)
    {
        auto str = getDefineGUIDFromGUID(entry.guid, entry.name.c_str());
        ::CharUpperW(&str[0]);
        if (str.find(text) != str.npos)
        {
            got.push_back(entry);
        }
    }

    entries = std::move(got);
    return !entries.empty();
}

BOOL IsHexGuidLikely(std::wstring arg)
{
    mstr_trim(arg, L" \t\r\n");
    mstr_replace_all(arg, L"0x", L"");

    if (arg.size() < 32)
        return FALSE;

    for (auto& ch : arg)
    {
        if (iswxdigit(ch))
            continue;
        if (ch == L',')
            continue;
        if (ch == L' ' || ch == L'\t')
            continue;
        return FALSE;
    }

    return TRUE;
}

void refguid_unittest(void)
{
#ifndef NDEBUG
    GUID guid = IID_IShellLinkW;

    std::vector<ENTRY> entries;
    assert(readAllGUIDs(entries));

    ENTRY found;
    assert(getByName(found, entries, L"IID_IShellLinkW"));
    assert(found.name == L"IID_IShellLinkW");
    assert(IsEqualGUID(found.guid, IID_IShellLinkW));

    assert(readAllGUIDs(entries));
    assert(doFilterByGuid(entries, guid));
    assert(entries.size() > 1);
    assert(entries[0].name == L"IID_IShellLinkW");
    assert(IsEqualGUID(entries[0].guid, IID_IShellLinkW));

    auto define_guid = getDefineGUIDFromGUID(guid, NULL);
    assert(define_guid ==
        L"DEFINE_GUID(<Name>, 0x000214F9, 0x0000, 0x0000, 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46);");

    WCHAR szCLSID[40];
    StringFromGUID2(guid, szCLSID, _countof(szCLSID));
    std::wstring str = L"{000214F9-0000-0000-C000-000000000046}";
    assert(str == szCLSID);

    auto strBytes = BytesTextFromCLSID(guid);
    assert(strBytes == L"F9 14 02 00 00 00 00 00 C0 00 00 00 00 00 00 46");

    auto struct_str = getStructFromGUID(guid);
    assert(struct_str == L"{ 0x000214F9, 0x0000, 0x0000, { 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46 } }");

    std::fprintf(stderr, "refguid_unittest: Success!\n\n");
#endif
}

int _wmain(int argc, wchar_t **argv)
{
    if (argc <= 1 || lstrcmpiW(argv[1], L"--help") == 0)
    {
        usage();
        return 0;
    }

    if (lstrcmpiW(argv[1], L"--version") == 0)
    {
        show_version();
        return 0;
    }

    std::wstring arg = argv[1];
    mstr_trim(arg, L" \t\r\n");

    if (argc >= 2 && argv[1][0] != '-')
    {
        for (int iarg = 2; iarg < argc; ++iarg)
        {
            arg += L" ";
            arg += argv[iarg];
        }
    }
    mstr_trim(arg, L" \t\r\n");

    std::vector<ENTRY> entries;
    if (!readAllGUIDs(entries))
        return 1;

    GUID guid = GUID_NULL;
    std::wstring name = L"";
    BOOL bSuccess = FALSE, bSearch = FALSE;

    if (lstrcmpiW(arg.c_str(), L"--list") == 0)
    {
        for (auto& entry : entries)
        {
            auto define_guid = getDefineGUIDFromGUID(entry.guid, entry.name.c_str());
            std::printf("%ls\n", define_guid.c_str());
        }
    }
    else if (lstrcmpiW(arg.c_str(), L"--generate") == 0)
    {
        RandomData(guid);
    }
    else if (lstrcmpiW(arg.c_str(), L"--search") == 0)
    {
        if (!doFilterByString(entries, argv[2]))
        {
            std::puts("Not found\n");
            return 1;
        }
        bSuccess = bSearch = TRUE;
        if (entries.size() == 1)
        {
            bSearch = FALSE;
            guid = entries[0].guid;
            name = entries[0].name;
        }
    }
    else if (arg[0] == L'{' && arg.find('{', 1) == arg.npos)
    {
        bSuccess = (CLSIDFromString(arg.c_str(), &guid) == S_OK);
    }
    else if (arg[0] == L'{' && arg.find('{', 1) != arg.npos)
    {
        bSuccess = getGUIDFromStructStr(guid, arg);
    }
    else if (arg.find(L"DEFINE_GUID(") == 0)
    {
        bSuccess = getGUIDFromDefineGUID(guid, arg, &name);
    }
    else if (IsHexGuidLikely(arg))
    {
        bSuccess = getGUIDFromBytesText(guid, arg);
    }
    else
    {
        ENTRY entry;
        bSuccess = getByName(entry, entries, arg.c_str());
        if (bSuccess)
        {
            name = entry.name;
            guid = entry.guid;
        }
        else if (!doFilterByString(entries, arg))
        {
            std::puts("Not found\n");
            return 1;
        }
        else
        {
            bSuccess = bSearch = TRUE;
        }
    }

    std::fprintf(stderr,
        "###########################\n"
        "## refguid by katahiromz ##\n"
        "###########################\n");

#ifndef NDEBUG
    refguid_unittest();
#endif

    if (!bSuccess || entries.empty())
    {
        std::printf("Not found\n");
        return -1;
    }

    if (bSearch)
    {
        printf("Found %d entries.\n", (int)entries.size());
        for (auto& entry : entries)
        {
            auto define_guid = getDefineGUIDFromGUID(entry.guid, entry.name.c_str());
            std::printf("%ls\n", define_guid.c_str());
        }
        return 0;
    }
    else if (doFilterByGuid(entries, guid))
    {
        bool first = true;
        for (auto& entry : entries)
        {
            if (first)
                std::printf("%4s: %ls\n", "Name", entry.name.c_str());
            else
                std::printf("%4s: %ls\n", "", entry.name.c_str());
            first = false;
        }
        std::puts("");
        if (name.empty())
            name = entries[0].name;
    }

    auto str = getAnalysisStringFromGUID(guid, name.c_str());
    std::printf("%ls", str.c_str());
    str = str;

    return 0;
}

int main(void)
{
    INT argc;
    LPWSTR *argv = CommandLineToArgvW(GetCommandLineW(), &argc);
    INT ret = _wmain(argc, argv);
    LocalFree(argv);
    return ret;
}
