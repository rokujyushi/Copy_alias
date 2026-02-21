#include <windows.h>
#include <string>
#include <cstring>
#include "plugin2.h"

static std::wstring Utf8ToUtf16(const char* utf8)
{
    if (!utf8) return L"";

    const int required = MultiByteToWideChar(CP_UTF8, 0, utf8, -1, nullptr, 0);
    if (required <= 0) return L"";

    std::wstring out;
    out.resize(static_cast<size_t>(required));
    MultiByteToWideChar(CP_UTF8, 0, utf8, -1, &out[0], required);
    if (!out.empty() && out.back() == L'\0') out.pop_back();
    return out;
}

static bool CopyUnicodeTextToClipboard(const std::wstring& text)
{
    if (text.empty()) return false;
    if (!OpenClipboard(nullptr)) return false;

    EmptyClipboard();

    const size_t bytes = (text.size() + 1) * sizeof(wchar_t);
    HGLOBAL hg = GlobalAlloc(GMEM_MOVEABLE, bytes);
    if (!hg)
    {
        CloseClipboard();
        return false;
    }

    void* ptr = GlobalLock(hg);
    if (!ptr)
    {
        GlobalFree(hg);
        CloseClipboard();
        return false;
    }

    memcpy(ptr, text.c_str(), bytes);
    GlobalUnlock(hg);

    if (!SetClipboardData(CF_UNICODETEXT, hg))
    {
        GlobalFree(hg);
        CloseClipboard();
        return false;
    }

    CloseClipboard();
    return true;
}

static void CopySelectedObjectAliases(EDIT_SECTION* edit)
{
    if (!edit) return;

    const int selected = edit->get_selected_object_num();
    if (selected <= 0)
    {
        // 仕様: 選択が無い場合は何もしない
        return;
    }

    std::wstring joined;
    bool hasAny = false;

    for (int i = 0; i < selected; ++i)
    {
        OBJECT_HANDLE obj = edit->get_selected_object(i);
        if (!obj) continue;

        const char* aliasUtf8 = edit->get_object_alias(obj);
        if (!aliasUtf8) continue;

        std::wstring alias = Utf8ToUtf16(aliasUtf8);
        if (alias.empty()) continue;

        if (hasAny)
        {
            joined += L"\r\n";
        }
        joined += alias;
        hasAny = true;
    }

    if (!hasAny)
    {
        return;
    }

    CopyUnicodeTextToClipboard(joined);
}

extern "C" __declspec(dllexport) void RegisterPlugin(HOST_APP_TABLE* host)
{
    if (!host) return;

    host->set_plugin_information(L"選択オブジェクトのエイリアスをコピー v1.0");
    host->register_object_menu(L"エイリアスをコピー", CopySelectedObjectAliases);
}

extern "C" __declspec(dllexport) bool InitializePlugin(DWORD)
{
    return true;
}

extern "C" __declspec(dllexport) void UninitializePlugin()
{
}
