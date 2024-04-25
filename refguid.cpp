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
#include <cctype>
#include <cwchar>
#include <strsafe.h>

BOOL g_bSearch = FALSE;
BOOL g_bList = FALSE;
INT g_nGenerate = 0;

void show_version(void)
{
    std::puts("refguid version 1.9.1 by katahiromz");
}

void usage(void)
{
    std::printf(
        "Usage:\n"
        "  refguid --search \"STRING\"\n"
        "  refguid IID_IDeskBand\n"
        "  refguid \"{EB0FE172-1A3A-11D0-89B3-00A0C90A90AC}\"\n"
        "  refguid \"72 E1 0F EB 3A 1A D0 11 89 B3 00 A0 C9 0A 90 AC\"\n"
        "  refguid \"{ 0xEB0FE172, 0x1A3A, 0x11D0, { 0x89, 0xB3, 0x00, 0xA0, 0xC9, 0x0A,\n"
        "             0x90, 0xAC } }\"\n"
        "  refguid \"DEFINE_GUID(IID_IDeskBand, 0xEB0FE172, 0x1A3A, 0x11D0, 0x89, 0xB3,\n"
        "                       0x00, 0xA0, 0xC9, 0x0A, 0x90, 0xAC);\"\n"
        "  refguid --list\n"
        "  refguid --generate NUMBER\n"
        "  refguid --help\n"
        "  refguid --version\n"
        "\n"
        "You can specify multiple GUIDs.\n");
}

typedef enum RET
{
    RET_DONE = 0,
    RET_FAILED = -1,
    RET_SUCCESS = +1,
} RET;

typedef struct tagENTRY
{
    std::wstring name;
    GUID guid;
} ENTRY, *PENTRY;

typedef std::vector<ENTRY> ENTRIES;

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

std::wstring getBytesTextFromGUID(REFCLSID clsid)
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

BOOL getGUIDFromStructStr(GUID& guid, std::wstring str)
{
    mstr_trim(str, L" \t\r\n");
    if (str.size() && str[str.size() - 1] == L';')
        str.erase(str.size() - 1, 1);
    mstr_trim(str, L" \t\r\n");
    if (str.size() && str[0] == L'{')
        str.erase(0, 1);
    else
        return FALSE;
    mstr_trim(str, L" \t\r\n");
    if (str.size() && str[str.size() - 1] == L'}')
        str.erase(str.size() - 1, 1);
    else
        return FALSE;
    mstr_trim(str, L" \t\r\n");
    if (str.size() && str[str.size() - 1] == L'}')
        str.erase(str.size() - 1, 1);
    else
        return FALSE;

    mstr_replace_all(str, L" ", L"");
    mstr_replace_all(str, L"\t", L"");

    auto ich0 = str.find(L',');
    if (ich0 != str.npos)
    {
        auto ich1 = str.find(L',', ich0 + 1);
        if (ich1 != str.npos)
        {
            auto ich2 = str.find(L',', ich1 + 1);
            if (ich2 == str.npos || str[ich2 + 1] != L'{')
                return FALSE;

            str.erase(ich2 + 1, 1);
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
    mstr_split(items, str, L",");

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

BOOL getGUIDFromDefineGUID(GUID& guid, std::wstring str, std::wstring *pstrName = NULL)
{
    mstr_trim(str, L" \t\r\n");

    if (str.find(L"DEFINE_GUID") == 0)
        str.erase(0, 11);
    else
        return FALSE;

    mstr_trim(str, L" \t\r\n");

    if (str.size() && str[0] == L'(')
        str.erase(0, 1);
    else
        return FALSE;

    mstr_trim(str, L" \t\r\n");

    if (str.size() && str[str.size() - 1] == L';')
        str.erase(str.size() - 1, 1);

    mstr_trim(str, L" \t\r\n");

    if (str.size() && str[str.size() - 1] == L')')
        str.erase(str.size() - 1, 1);
    else
        return FALSE;

    mstr_trim(str, L" \t\r\n");
    mstr_replace_all(str, L" ", L"");
    mstr_replace_all(str, L"\t", L"");

    std::vector<std::wstring> items;
    mstr_split(items, str, L",");

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

BOOL getGUIDFromBytesText(GUID& guid, std::wstring str)
{
    mstr_replace_all(str, L"0x", L"");

    CLSID ret;
    std::vector<BYTE> bytes;
    WCHAR sz[3];
    size_t ich = 0, ib = 0;
    for (auto ch : str)
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

std::wstring getDumpFromGUID(REFGUID guid, LPCWSTR name)
{
    if (name && name[0] == 0)
        name = NULL;

    std::wstring ret;

    auto define_guid = getDefineGUIDFromGUID(guid, name);
    ret += define_guid;
    ret += L"\n\n";

    WCHAR szGUID[40];
    StringFromGUID2(guid, szGUID, _countof(szGUID));
    ret += L"GUID: ";
    ret += szGUID;
    ret += L"\n\n";

    auto strBytes = getBytesTextFromGUID(guid);
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

BOOL readGUIDs(ENTRIES& entries)
{
    WCHAR szPath[MAX_PATH];
    GetModuleFileNameW(NULL, szPath, _countof(szPath));
    PathRemoveFileSpecW(szPath);
    PathAppendW(szPath, L"refguid.dat");

    FILE *fp = _wfopen(szPath, L"r");
    if (!fp)
    {
        fprintf(stderr, "WARNING: Cannot read file %ls\n", szPath);
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

BOOL getEntryByName(ENTRY& found, const ENTRIES& entries, std::wstring name)
{
    ::CharUpperW(&name[0]);

    for (auto& entry : entries)
    {
        std::wstring entry_name = entry.name;
        ::CharUpperW(&entry_name[0]);

        if (lstrcmpiW(entry_name.c_str(), name.c_str()) == 0)
        {
            found = entry;
            return TRUE;
        }
    }

    return FALSE;
}

BOOL getEntriesByGuid(ENTRIES& got, const ENTRIES& entries, const GUID& guid)
{
    for (auto& entry : entries)
    {
        if (IsEqualGUID(entry.guid, guid))
        {
            got.push_back(entry);
        }
    }

    return !got.empty();
}

BOOL getEntriesByString(ENTRIES& found, const ENTRIES& entries, std::wstring text)
{
    ::CharUpperW(&text[0]);

    for (auto& entry : entries)
    {
        auto str = getDefineGUIDFromGUID(entry.guid, entry.name.c_str());
        ::CharUpperW(&str[0]);

        if (str.find(text) != str.npos)
        {
            found.push_back(entry);
            continue;
        }

        str = getBytesTextFromGUID(entry.guid);
        ::CharUpperW(&str[0]);

        if (str.find(text) != str.npos)
        {
            found.push_back(entry);
            continue;
        }
    }

    return !found.empty();
}

BOOL isHexGuid(std::wstring str)
{
    mstr_trim(str, L" \t\r\n");
    mstr_replace_all(str, L"0x", L"");
    mstr_replace_all(str, L",", L"");
    mstr_replace_all(str, L" ", L"");
    mstr_replace_all(str, L"\t", L"");

    if (str.size() != 32)
        return FALSE;

    for (auto& ch : str)
    {
        if (!iswxdigit(ch))
            return FALSE;
    }

    return TRUE;
}

BOOL isGUIDString(LPCWSTR pszGUID)
{
    GUID guid;
    return CLSIDFromString(pszGUID, &guid) == S_OK;
}

BOOL isGUIDStruct(LPCWSTR pszGUID)
{
    GUID guid;
    return getGUIDFromStructStr(guid, pszGUID);
}

BOOL isDefineGUID(LPCWSTR pszGUID)
{
    GUID guid;
    return getGUIDFromDefineGUID(guid, pszGUID, NULL);
}

void refguid_unittest(void)
{
#ifndef NDEBUG
    GUID guid = IID_IShellLinkW;

    ENTRIES entries;
    assert(readGUIDs(entries));

    ENTRY entry;
    assert(getEntryByName(entry, entries, L"IID_IShellLinkW"));
    assert(entry.name == L"IID_IShellLinkW");
    assert(IsEqualGUID(entry.guid, IID_IShellLinkW));

    ENTRIES got;
    assert(getEntriesByGuid(got, entries, guid));
    assert(got.size() > 1);
    assert(got[0].name == L"IID_IShellLinkW");
    assert(IsEqualGUID(got[0].guid, IID_IShellLinkW));

    auto define_guid = getDefineGUIDFromGUID(guid, NULL);
    assert(define_guid ==
        L"DEFINE_GUID(<Name>, 0x000214F9, 0x0000, 0x0000, 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46);");
    assert(isDefineGUID(define_guid.c_str()));

    WCHAR szGUID[40];
    StringFromGUID2(guid, szGUID, _countof(szGUID));
    std::wstring str = L"{000214F9-0000-0000-C000-000000000046}";
    assert(str == szGUID);
    assert(isGUIDString(str.c_str());

    auto strBytes = getBytesTextFromGUID(guid);
    assert(strBytes == L"F9 14 02 00 00 00 00 00 C0 00 00 00 00 00 00 46");
    assert(isHexGuid(str.c_str());

    auto struct_str = getStructFromGUID(guid);
    assert(struct_str == L"{ 0x000214F9, 0x0000, 0x0000, { 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46 } }");
    assert(isGUIDStruct(str.c_str());

    std::fprintf(stderr, "refguid_unittest: Success!\n\n");
#endif
}

RET refguid_do_guid(const ENTRIES& entries, REFGUID guid, std::wstring *pstrName = NULL)
{
    std::printf("\n--------------------\n");

    std::wstring name;
    if (pstrName == NULL)
    {
        BOOL found = FALSE;
        for (auto& entry : entries)
        {
            if (IsEqualGUID(guid, entry.guid))
            {
                found = TRUE;
                name = entry.name;
                std::printf("Name: %ls\n", name.c_str());
            }
        }

        if (found)
        {
            std::printf("\n");
            pstrName = &name;
        }
    }

    auto str = getDumpFromGUID(guid, (pstrName ? pstrName->c_str() : NULL));
    std::printf("%ls", str.c_str());
    name = name;
    return RET_SUCCESS;
}

RET refguid_do_arg(const ENTRIES& entries, std::wstring str)
{
    mstr_trim(str, L" \t\r\n");

    if (isGUIDString(str.c_str()))
    {
        GUID guid;
        if (CLSIDFromString(str.c_str(), &guid) == S_OK)
            return refguid_do_guid(entries, guid);
        return RET_FAILED;
    }

    if (isGUIDStruct(str.c_str()))
    {
        GUID guid;
        if (getGUIDFromStructStr(guid, str))
            return refguid_do_guid(entries, guid);
        return RET_FAILED;
    }

    if (isDefineGUID(str.c_str()))
    {
        GUID guid;
        std::wstring name;
        if (getGUIDFromDefineGUID(guid, str, &name))
            return refguid_do_guid(entries, guid, &name);
        return RET_FAILED;
    }

    if (isHexGuid(str))
    {
        GUID guid;
        if (getGUIDFromBytesText(guid, str))
            return refguid_do_guid(entries, guid);
        return RET_FAILED;
    }

    if (!g_bSearch)
    {
        ENTRY entry;
        if (getEntryByName(entry, entries, str))
            return refguid_do_guid(entries, entry.guid, &entry.name);
    }

    ENTRIES found;
    if (getEntriesByString(found, entries, str))
    {
        if (found.size() == 1)
            return refguid_do_guid(entries, found[0].guid, &found[0].name);
    }
    else
    {
        std::printf("Not found\n");
        return RET_FAILED;
    }

    if (found.size() > 1)
    {
        std::printf("Found %d entries.\n", (int)found.size());
    }

    for (auto& item : found)
    {
        refguid_do_guid(entries, item.guid, &item.name);
    }

    return RET_SUCCESS;
}

RET parse_option(std::wstring str)
{
    if (str == L"--help")
    {
        usage();
        return RET_DONE;
    }

    if (str == L"--version")
    {
        show_version();
        return RET_DONE;
    }

    if (str == L"--list")
    {
        g_bList = TRUE;
        return RET_SUCCESS;
    }

    if (str == L"--search")
    {
        g_bSearch = TRUE;
        return RET_SUCCESS;
    }

    fprintf(stderr, "Invalid option: %ls\n", str.c_str());
    return RET_FAILED;
}

BOOL IsIdentifier(LPCWSTR pszParam)
{
    if (pszParam[0] == 0)
        return FALSE;

    for (size_t ich = 0; pszParam[ich]; ++ich)
    {
        WCHAR ch = pszParam[ich];
        if (ich == 0 && !iswalpha(ch) && ch != L'_')
            return FALSE;
        if (ich > 0 && !iswalnum(ch) && ch != L'_')
            return FALSE;
    }
    return TRUE;
}

RET parse_cmd_line(std::vector<std::wstring>& args, int argc, wchar_t **argv)
{
    if (argc <= 1)
    {
        usage();
        return RET_FAILED;
    }

    std::wstring param;

    for (int iarg = 1; iarg < argc; ++iarg)
    {
        std::wstring str = argv[iarg];

        if (str[0] == L'-')
        {
            if (str == L"--generate")
            {
                if (argc <= iarg + 1)
                    g_nGenerate = 1;
                else
                    g_nGenerate = _wtoi(argv[iarg + 1]);

                if (g_nGenerate <= 0)
                {
                    fprintf(stderr, "ERROR: Zero or negative value specified for '--generate'\n");
                    return RET_FAILED;
                }

                param.clear();
                args.clear();
                return RET_SUCCESS;
            }

            switch (parse_option(str))
            {
            case RET_FAILED:
                return RET_FAILED;
            case RET_DONE:
                return RET_DONE;
            case RET_SUCCESS:
                continue;
            }
        }

        if (param.size())
            param += L' ';

        param += str;

        if (isGUIDString(param.c_str()) || isGUIDStruct(param.c_str()) ||
            isDefineGUID(param.c_str()) || isHexGuid(param.c_str()))
        {
            args.push_back(param);
            param.clear();
            continue;
        }

        if (IsIdentifier(param.c_str()) && param != L"DEFINE_GUID")
        {
            args.push_back(param);
            param.clear();
            continue;
        }
    }

    if (param.size())
        args.push_back(param);

    return RET_SUCCESS;
}

int _wmain(int argc, wchar_t **argv)
{
#ifndef NDEBUG
    refguid_unittest();
#endif

    std::vector<std::wstring> args;
    RET ret = parse_cmd_line(args, argc, argv);
    if (ret == RET_DONE)
        return 0;
    if (ret == RET_FAILED)
        return -1;

    ENTRIES entries;
    if (!readGUIDs(entries))
        return -2;

    if (entries.empty() && !g_bList && g_nGenerate == 0)
    {
        std::printf("ERROR: File refguid.dat was empty\n");
        return -3;
    }

    if (g_bList)
    {
        for (auto& entry : entries)
        {
            auto define_guid = getDefineGUIDFromGUID(entry.guid, entry.name.c_str());
            std::printf("%ls\n", define_guid.c_str());
        }
        return 0;
    }

    if (g_nGenerate > 0)
    {
        for (INT i = 0; i < g_nGenerate; ++i)
        {
            GUID guid;
            RandomData(guid);
            refguid_do_guid(entries, guid);
        }
        return 0;
    }

    for (auto& item : args)
    {
        if (refguid_do_arg(entries, item) == RET_FAILED)
            return -4;
    }

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
