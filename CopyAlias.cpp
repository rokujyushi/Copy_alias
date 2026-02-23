#ifndef NOMINMAX
#define NOMINMAX
#endif
#include <windows.h>
#include <commctrl.h>
#include <string>
#include <cstring>
#include <vector>
#include <map>
#include <sstream>
#include <algorithm>
#include <cwctype>
#include <cstdio>
#include "plugin2.h"
#include "logger2.h"

#define IDC_INI_TREE     2101
#define IDC_INI_APPLY    2102
#define IDC_INI_CLOSE    2103

static EDIT_HANDLE* g_edit = nullptr;
static HWND g_iniDialog = nullptr;
static LOG_HANDLE* g_logger = nullptr;

static void PositionDialogNearCursor(HWND hwnd);

static void LogInfo(const wchar_t* msg)
{
    if (g_logger && g_logger->info) g_logger->info(g_logger, msg);
    OutputDebugStringW(msg);
    OutputDebugStringW(L"\n");
}

static void LogWarn(const wchar_t* msg)
{
    if (g_logger && g_logger->warn) g_logger->warn(g_logger, msg);
    OutputDebugStringW(msg);
    OutputDebugStringW(L"\n");
}

static void LogError(const wchar_t* msg)
{
    if (g_logger && g_logger->error) g_logger->error(g_logger, msg);
    OutputDebugStringW(msg);
    OutputDebugStringW(L"\n");
}

struct PropertyEntry {
    std::wstring key;
    std::wstring value;
};

struct EffectEntry {
    int objectIndex = -1;
    std::wstring effectName;
    int occurrence = 0;
    std::vector<PropertyEntry> properties;
};

struct BlockEntry {
    std::vector<EffectEntry> effects;
};

struct ParsedClipboard {
    std::vector<BlockEntry> blocks;
};

struct ApplyItem {
    std::wstring effectName;
    int occurrence = 0;
    std::wstring effectSelector;
    std::wstring propertyKey;
    std::string valueUtf8;
};

struct ApplySummary {
    int targetObjects = 0;
    int attempted = 0;
    int applied = 0;
    int skippedEffectMismatch = 0;
    int failedSet = 0;
};

enum class NodeKind {
    Block,
    Effect,
    Property
};

struct NodeRef {
    NodeKind kind;
    int blockIndex;
    int effectIndex;
    int propertyIndex;
};

struct IniDialogState {
    ParsedClipboard parsed;
    std::vector<NodeRef> refs;
    std::vector<ApplyItem> selectedItems;
    ApplySummary summary;
    std::vector<OBJECT_HANDLE> targetObjects;
    HWND tree = nullptr;
    bool updatingChecks = false;
};

class DialogBuilder {
    std::vector<unsigned char> data;
public:
    void Align() { while (data.size() % 4 != 0) data.push_back(0); }
    void Add(const void* p, size_t s) { const unsigned char* b = (const unsigned char*)p; data.insert(data.end(), b, b + s); }
    void AddW(WORD w) { Add(&w, 2); }
    void AddD(DWORD d) { Add(&d, 4); }
    void AddS(LPCWSTR s) { if (!s) AddW(0); else Add(s, (wcslen(s) + 1) * 2); }

    void Begin(LPCWSTR title, int w, int h, DWORD style) {
        data.clear();
        Align();
        DWORD exStyle = WS_EX_CONTROLPARENT;
        AddD(style); AddD(exStyle);
        AddW(0);
        AddW(0); AddW(0); AddW((WORD)w); AddW((WORD)h);
        AddW(0); AddW(0); AddS(title); AddW(9); AddS(L"Yu Gothic UI");
    }

    void AddControl(LPCWSTR className, LPCWSTR text, WORD id, int x, int y, int w, int h, DWORD style) {
        Align();
        AddD(style | WS_CHILD | WS_VISIBLE);
        AddD(0);
        AddW((WORD)x); AddW((WORD)y); AddW((WORD)w); AddW((WORD)h);
        AddW(id);

        if (wcscmp(className, L"BUTTON") == 0) {
            AddW(0xFFFF); AddW(0x0080);
            AddS(text);
            AddW(0);
        }
        else if (wcscmp(className, L"STATIC") == 0) {
            AddW(0xFFFF); AddW(0x0082);
            AddS(text);
            AddW(0);
        }
        else {
            AddS(className);
            AddS(text);
            AddW(0);
        }

        if (data.size() >= 10) {
            WORD* pCount = (WORD*)&data[8];
            (*pCount)++;
        }
    }

    DLGTEMPLATE* Get() { return (DLGTEMPLATE*)data.data(); }
};

static std::wstring Trim(const std::wstring& s)
{
    size_t begin = 0;
    while (begin < s.size() && iswspace(s[begin])) ++begin;
    size_t end = s.size();
    while (end > begin && iswspace(s[end - 1])) --end;
    return s.substr(begin, end - begin);
}

static bool IEquals(const std::wstring& a, const wchar_t* b)
{
    return _wcsicmp(a.c_str(), b) == 0;
}

static bool StartsWithI(const std::wstring& a, const wchar_t* prefix)
{
    const size_t n = wcslen(prefix);
    if (a.size() < n) return false;
    return _wcsnicmp(a.c_str(), prefix, n) == 0;
}

static std::string Utf16ToUtf8(const std::wstring& w)
{
    if (w.empty()) return std::string();
    const int required = WideCharToMultiByte(CP_UTF8, 0, w.c_str(), -1, nullptr, 0, nullptr, nullptr);
    if (required <= 1) return std::string();

    std::string out;
    out.resize(static_cast<size_t>(required - 1));
    WideCharToMultiByte(CP_UTF8, 0, w.c_str(), -1, &out[0], required, nullptr, nullptr);
    return out;
}

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

static bool GetClipboardUnicodeText(std::wstring& out)
{
    out.clear();
    if (!IsClipboardFormatAvailable(CF_UNICODETEXT)) return false;
    if (!OpenClipboard(nullptr)) return false;

    HANDLE h = GetClipboardData(CF_UNICODETEXT);
    if (!h) {
        CloseClipboard();
        return false;
    }

    const wchar_t* p = static_cast<const wchar_t*>(GlobalLock(h));
    if (!p) {
        CloseClipboard();
        return false;
    }

    out = p;
    GlobalUnlock(h);
    CloseClipboard();
    return true;
}

static bool TryParseObjectSectionIndex(const std::wstring& sectionName, int& index)
{
    if (!StartsWithI(sectionName, L"Object.")) return false;
    const std::wstring num = sectionName.substr(7);
    if (num.empty()) return false;
    for (wchar_t c : num) {
        if (c < L'0' || c > L'9') return false;
    }
    index = _wtoi(num.c_str());
    return index >= 0;
}

static ParsedClipboard ParseClipboardIni(const std::wstring& text)
{
    ParsedClipboard parsed;
    if (text.empty()) return parsed;

    BlockEntry* currentBlock = nullptr;
    EffectEntry* currentEffect = nullptr;

    std::wistringstream ss(text);
    std::wstring line;
    bool firstLine = true;
    while (std::getline(ss, line)) {
        if (!line.empty() && line.back() == L'\r') line.pop_back();
        if (firstLine && !line.empty() && line.front() == 0xFEFF) {
            line.erase(line.begin());
        }
        firstLine = false;

        std::wstring t = Trim(line);
        if (t.empty()) continue;
        if (t[0] == L';' || t[0] == L'#') continue;

        if (t.size() >= 3 && t.front() == L'[' && t.back() == L']') {
            std::wstring section = Trim(t.substr(1, t.size() - 2));
            currentEffect = nullptr;

            if (IEquals(section, L"Object")) {
                parsed.blocks.push_back(BlockEntry{});
                currentBlock = &parsed.blocks.back();
                continue;
            }

            int objectIndex = -1;
            if (TryParseObjectSectionIndex(section, objectIndex)) {
                if (!currentBlock) {
                    parsed.blocks.push_back(BlockEntry{});
                    currentBlock = &parsed.blocks.back();
                }
                currentBlock->effects.push_back(EffectEntry{});
                currentEffect = &currentBlock->effects.back();
                currentEffect->objectIndex = objectIndex;
                continue;
            }

            continue;
        }

        size_t eq = t.find(L'=');
        if (eq == std::wstring::npos) continue;
        std::wstring key = Trim(t.substr(0, eq));
        std::wstring val = t.substr(eq + 1);

        // [Object] メタ情報は表示・反映対象外
        if (!currentEffect) continue;

        if (IEquals(key, L"effect.name")) {
            currentEffect->effectName = Trim(val);
        }
        else {
            PropertyEntry p;
            p.key = key;
            p.value = val;
            currentEffect->properties.push_back(std::move(p));
        }
    }

    ParsedClipboard filtered;
    for (auto& b : parsed.blocks) {
        BlockEntry outBlock;
        std::map<std::wstring, int> occ;

        for (auto& e : b.effects) {
            if (Trim(e.effectName).empty()) {
                // effect.name 欠落セクションは非表示
                continue;
            }
            if (e.properties.empty()) {
                continue;
            }

            EffectEntry out = e;
            out.occurrence = occ[out.effectName]++;
            outBlock.effects.push_back(std::move(out));
        }

        if (!outBlock.effects.empty()) {
            filtered.blocks.push_back(std::move(outBlock));
        }
    }

    return filtered;
}

static std::wstring BuildEffectSelector(const std::wstring& effectName, int occurrence)
{
    if (occurrence <= 0) return effectName;
    return effectName + L":" + std::to_wstring(occurrence);
}

static bool GetTreeCheckState(HWND tree, HTREEITEM item)
{
    TVITEMW tv{};
    tv.mask = TVIF_HANDLE | TVIF_STATE;
    tv.hItem = item;
    tv.stateMask = TVIS_STATEIMAGEMASK;
    if (!TreeView_GetItem(tree, &tv)) return false;
    return ((tv.state & TVIS_STATEIMAGEMASK) >> 12) == 2;
}

static void SetTreeCheckState(HWND tree, HTREEITEM item, bool checked)
{
    TVITEMW tv{};
    tv.mask = TVIF_HANDLE | TVIF_STATE;
    tv.hItem = item;
    tv.stateMask = TVIS_STATEIMAGEMASK;
    tv.state = INDEXTOSTATEIMAGEMASK(checked ? 2 : 1);
    TreeView_SetItem(tree, &tv);
}

static void SetChildrenCheckState(HWND tree, HTREEITEM parent, bool checked)
{
    for (HTREEITEM child = TreeView_GetChild(tree, parent); child; child = TreeView_GetNextSibling(tree, child)) {
        SetTreeCheckState(tree, child, checked);
        SetChildrenCheckState(tree, child, checked);
    }
}

static void UpdateParentCheckState(HWND tree, HTREEITEM item)
{
    HTREEITEM parent = TreeView_GetParent(tree, item);
    while (parent) {
        bool hasChild = false;
        bool allChecked = true;
        for (HTREEITEM child = TreeView_GetChild(tree, parent); child; child = TreeView_GetNextSibling(tree, child)) {
            hasChild = true;
            if (!GetTreeCheckState(tree, child)) {
                allChecked = false;
                break;
            }
        }
        if (hasChild) {
            SetTreeCheckState(tree, parent, allChecked);
        }
        parent = TreeView_GetParent(tree, parent);
    }
}

static size_t CountTotalTreeNodes(const ParsedClipboard& parsed)
{
    size_t total = 0;
    for (const auto& b : parsed.blocks) {
        total += 1;
        for (const auto& e : b.effects) {
            total += 1;
            total += e.properties.size();
        }
    }
    return total;
}

static void InsertTreeNode(HWND tree, HTREEITEM parent, const std::wstring& text, const NodeRef* ref, HTREEITEM& outItem)
{
    TVINSERTSTRUCTW ins{};
    ins.hParent = parent;
    ins.hInsertAfter = TVI_LAST;
    ins.item.mask = TVIF_TEXT | TVIF_PARAM;
    ins.item.pszText = const_cast<LPWSTR>(text.c_str());
    ins.item.lParam = reinterpret_cast<LPARAM>(ref);
    outItem = TreeView_InsertItem(tree, &ins);
    if (outItem) {
        SetTreeCheckState(tree, outItem, true);
    }
}

static void PopulateTreeFromParsed(IniDialogState* state)
{
    if (!state || !state->tree) return;
    TreeView_DeleteAllItems(state->tree);
    state->refs.clear();
    state->refs.reserve(CountTotalTreeNodes(state->parsed));

    for (int bi = 0; bi < (int)state->parsed.blocks.size(); ++bi) {
        const auto& block = state->parsed.blocks[bi];

        state->refs.push_back(NodeRef{ NodeKind::Block, bi, -1, -1 });
        NodeRef* bRef = &state->refs.back();
        HTREEITEM blockItem = nullptr;
        InsertTreeNode(state->tree, TVI_ROOT, L"ブロック " + std::to_wstring(bi + 1), bRef, blockItem);

        for (int ei = 0; ei < (int)block.effects.size(); ++ei) {
            const auto& eff = block.effects[ei];
            std::wstring effLabel = L"Object." + std::to_wstring(eff.objectIndex) + L" : " + eff.effectName;
            if (eff.occurrence > 0) {
                effLabel += L" (" + std::to_wstring(eff.occurrence + 1) + L"個目)";
            }

            state->refs.push_back(NodeRef{ NodeKind::Effect, bi, ei, -1 });
            NodeRef* eRef = &state->refs.back();
            HTREEITEM effItem = nullptr;
            InsertTreeNode(state->tree, blockItem, effLabel, eRef, effItem);

            for (int pi = 0; pi < (int)eff.properties.size(); ++pi) {
                const auto& prop = eff.properties[pi];
                std::wstring propLabel = prop.key + L"=" + prop.value;
                state->refs.push_back(NodeRef{ NodeKind::Property, bi, ei, pi });
                NodeRef* pRef = &state->refs.back();
                HTREEITEM propItem = nullptr;
                InsertTreeNode(state->tree, effItem, propLabel, pRef, propItem);
            }

            if (effItem) TreeView_Expand(state->tree, effItem, TVE_EXPAND);
        }

        if (blockItem) TreeView_Expand(state->tree, blockItem, TVE_EXPAND);
    }
}

static void CollectCheckedPropertyItemsRecursive(IniDialogState* state, HTREEITEM item)
{
    if (!state || !state->tree || !item) return;

    TVITEMW tv{};
    tv.mask = TVIF_HANDLE | TVIF_PARAM;
    tv.hItem = item;
    if (TreeView_GetItem(state->tree, &tv)) {
        NodeRef* ref = reinterpret_cast<NodeRef*>(tv.lParam);
        if (ref && ref->kind == NodeKind::Property && GetTreeCheckState(state->tree, item)) {
            if (ref->blockIndex >= 0 && ref->blockIndex < (int)state->parsed.blocks.size()) {
                const auto& block = state->parsed.blocks[ref->blockIndex];
                if (ref->effectIndex >= 0 && ref->effectIndex < (int)block.effects.size()) {
                    const auto& eff = block.effects[ref->effectIndex];
                    if (ref->propertyIndex >= 0 && ref->propertyIndex < (int)eff.properties.size()) {
                        const auto& prop = eff.properties[ref->propertyIndex];
                        ApplyItem it;
                        it.effectName = eff.effectName;
                        it.occurrence = eff.occurrence;
                        it.effectSelector = BuildEffectSelector(eff.effectName, eff.occurrence);
                        it.propertyKey = prop.key;
                        it.valueUtf8 = Utf16ToUtf8(prop.value);
                        state->selectedItems.push_back(std::move(it));
                    }
                }
            }
        }
    }

    for (HTREEITEM child = TreeView_GetChild(state->tree, item); child; child = TreeView_GetNextSibling(state->tree, child)) {
        CollectCheckedPropertyItemsRecursive(state, child);
    }
}

static void CollectCheckedPropertyItems(IniDialogState* state)
{
    if (!state || !state->tree) return;
    state->selectedItems.clear();
    for (HTREEITEM root = TreeView_GetRoot(state->tree); root; root = TreeView_GetNextSibling(state->tree, root)) {
        CollectCheckedPropertyItemsRecursive(state, root);
    }
}

static void ApplyItemsToSelectedObjects(void* param, EDIT_SECTION* edit)
{
    if (!param || !edit) return;
    IniDialogState* state = reinterpret_cast<IniDialogState*>(param);
    state->summary = ApplySummary{};

    state->summary.targetObjects = (int)state->targetObjects.size();
    if (state->summary.targetObjects <= 0) return;

    for (OBJECT_HANDLE obj : state->targetObjects) {
        if (!obj) continue;

        for (const auto& it : state->selectedItems) {
            state->summary.attempted++;

            const int count = edit->count_object_effect(obj, it.effectName.c_str());
            if (count <= it.occurrence) {
                state->summary.skippedEffectMismatch++;
                continue;
            }

            const bool ok = edit->set_object_item_value(
                obj,
                it.effectSelector.c_str(),
                it.propertyKey.c_str(),
                it.valueUtf8.c_str());

            if (ok) state->summary.applied++;
            else state->summary.failedSet++;
        }
    }
}

static INT_PTR CALLBACK IniApplyDlgProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp)
{
    switch (msg) {
    case WM_INITDIALOG: {
        IniDialogState* state = reinterpret_cast<IniDialogState*>(lp);
        SetWindowLongPtr(hwnd, GWLP_USERDATA, (LONG_PTR)state);
        g_iniDialog = hwnd;

        // DialogTemplateはDLU、子コントロールはピクセルで扱っているため
        // ここで最終サイズをピクセルで明示して見た目を安定させる
        // SetWindowPos(hwnd, nullptr, 0, 0, 400, 500, SWP_NOMOVE | SWP_NOZORDER);

        state->tree = CreateWindowExW(
            WS_EX_CLIENTEDGE,
            WC_TREEVIEWW,
            L"",
            WS_CHILD | WS_VISIBLE | WS_TABSTOP |
            TVS_CHECKBOXES | TVS_HASBUTTONS | TVS_HASLINES | TVS_LINESATROOT | TVS_SHOWSELALWAYS,
            CW_USEDEFAULT, CW_USEDEFAULT, 400, 500,
            hwnd,
            (HMENU)IDC_INI_TREE,
            GetModuleHandle(nullptr),
            nullptr);

        if (!state->tree) return TRUE;

        // ダイアログのクライアント領域サイズを取得（タイトルバー等を除いた内側）
        RECT rc{};
        GetClientRect(hwnd, &rc);
        const int clientW = rc.right - rc.left;
        const int clientH = rc.bottom - rc.top;

        // テンプレートで作成済みの既存コントロールを取得
        HWND hDesc = GetDlgItem(hwnd, 1);               // 説明ラベル
        HWND hApply = GetDlgItem(hwnd, IDC_INI_APPLY);  // 適用ボタン
        HWND hClose = GetDlgItem(hwnd, IDC_INI_CLOSE);  // 閉じるボタン

        // 各コントロールを「現在のダイアログサイズ」に合わせて再配置
        // （DialogTemplateはDLU、ここはpx基準なので初期化時に見た目を統一）
        if (hDesc) MoveWindow(hDesc, 10, 8, clientW - 20, 18, TRUE);                      // 上部説明
        MoveWindow(state->tree, 10, 30, clientW - 20, clientH - 70, TRUE);                // 中央ツリー（余白確保）
        if (hApply) MoveWindow(hApply, clientW - 190, clientH - 32, 90, 22, TRUE);        // 右下: 適用
        if (hClose) MoveWindow(hClose, clientW - 95, clientH - 32, 85, 22, TRUE);         // 右下: 閉じる

        PopulateTreeFromParsed(state);
        PositionDialogNearCursor(hwnd);
        return TRUE;
    }
    case WM_NOTIFY: {
        // WM_INITDIALOGでGWLP_USERDATAに保存した state を再取得
        IniDialogState* state = reinterpret_cast<IniDialogState*>(GetWindowLongPtr(hwnd, GWLP_USERDATA));
        if (!state || !state->tree) break;

        // 通知元がツリービュー以外なら無視
        NMHDR* hdr = reinterpret_cast<NMHDR*>(lp);
        if (!hdr || hdr->idFrom != IDC_INI_TREE) break;

        // チェック状態の変更通知のみを拾う
        if (hdr->code == TVN_ITEMCHANGEDW && !state->updatingChecks) {
            NMTVITEMCHANGE* ch = reinterpret_cast<NMTVITEMCHANGE*>(lp);
            if (!ch) break;

            // state image(チェックボックス)が変わっていない通知は無視
            const UINT delta = (ch->uStateOld ^ ch->uStateNew) & TVIS_STATEIMAGEMASK;
            if (delta == 0) break;

            // 親子連動でチェック状態を同期
            const bool checked = ((ch->uStateNew & TVIS_STATEIMAGEMASK) >> 12) == 2;
            // SetChildren/UpdateParent 内で再帰的に通知が来るためガードを立てる
            state->updatingChecks = true;
            SetChildrenCheckState(state->tree, ch->hItem, checked);
            UpdateParentCheckState(state->tree, ch->hItem);
            state->updatingChecks = false;
        }
        break;
    }
    case WM_COMMAND: {
        const int id = LOWORD(wp);
        if (id == IDC_INI_APPLY) {
            IniDialogState* state = reinterpret_cast<IniDialogState*>(GetWindowLongPtr(hwnd, GWLP_USERDATA));
            if (!state) return TRUE;

            CollectCheckedPropertyItems(state);
            if (state->selectedItems.empty()) {
                LogInfo(L"CopyAlias: 適用するプロパティが選択されていません。");
                return TRUE;
            }

            if (!g_edit) {
                LogError(L"CopyAlias: 編集ハンドルの取得に失敗しています。");
                return TRUE;
            }

            g_edit->call_edit_section_param(state, ApplyItemsToSelectedObjects);

            wchar_t msgBuf[512];
            if (state->summary.targetObjects <= 0) {
                swprintf_s(msgBuf, L"選択中オブジェクトがありません。\n処理は実行されませんでした。");
                LogInfo(msgBuf);
                DestroyWindow(hwnd);
                return TRUE;
            }

            swprintf_s(
                msgBuf,
                L"適用対象オブジェクト: %d\n"
                L"試行数: %d\n"
                L"適用成功: %d\n"
                L"エフェクト不一致スキップ: %d\n"
                L"設定失敗: %d",
                state->summary.targetObjects,
                state->summary.attempted,
                state->summary.applied,
                state->summary.skippedEffectMismatch,
                state->summary.failedSet);

            LogInfo(msgBuf);
            DestroyWindow(hwnd);
            return TRUE;
        }
        if (id == IDC_INI_CLOSE) {
            DestroyWindow(hwnd);
            return TRUE;
        }
        break;
    }
    case WM_CLOSE:
        DestroyWindow(hwnd);
        return TRUE;
    case WM_NCDESTROY: {
        IniDialogState* state = reinterpret_cast<IniDialogState*>(GetWindowLongPtr(hwnd, GWLP_USERDATA));
        SetWindowLongPtr(hwnd, GWLP_USERDATA, 0);
        if (state) delete state;
        if (g_iniDialog == hwnd) g_iniDialog = nullptr;
        return TRUE;
    }
    }
    return FALSE;
}

static void PositionDialogNearCursor(HWND hwnd)
{
    POINT pt{};
    if (!GetCursorPos(&pt)) return;

    RECT wr{};
    if (!GetWindowRect(hwnd, &wr)) return;

    const int w = wr.right - wr.left;
    const int h = wr.bottom - wr.top;

    HMONITOR mon = MonitorFromPoint(pt, MONITOR_DEFAULTTONEAREST);
    MONITORINFO mi{};
    mi.cbSize = sizeof(mi);
    if (!GetMonitorInfoW(mon, &mi)) return;

    const RECT& rc = mi.rcWork; // タスクバーを除く領域
    int x = pt.x + 0;
    int y = pt.y + 16;

    const int minX = (int)rc.left;
    const int minY = (int)rc.top;
    const int maxX = (int)rc.right - w;
    const int maxY = (int)rc.bottom - h;
    x = std::max(minX, std::min(x, maxX));
    y = std::max(minY, std::min(y, maxY));

    SetWindowPos(hwnd, nullptr, x, y, 0, 0, SWP_NOZORDER | SWP_NOSIZE | SWP_NOACTIVATE);
}

static void OpenIniApplyDialog(EDIT_SECTION* edit)
{
    if (g_iniDialog) {
        SetForegroundWindow(g_iniDialog);
        return;
    }

    if (!edit) {
        LogError(L"CopyAlias: 編集情報を取得できませんでした。");
        return;
    }

    std::vector<OBJECT_HANDLE> capturedTargets;
    const int selected = edit->get_selected_object_num();
    for (int i = 0; i < selected; ++i) {
        OBJECT_HANDLE obj = edit->get_selected_object(i);
        if (obj) capturedTargets.push_back(obj);
    }
    if (capturedTargets.empty()) {
        LogInfo(L"CopyAlias: 対象オブジェクトを選択してから実行してください。");
        return;
    }

    std::wstring clip;
    if (!GetClipboardUnicodeText(clip) || clip.empty()) {
        LogInfo(L"CopyAlias: クリップボードにテキストがありません。");
        return;
    }

    ParsedClipboard parsed = ParseClipboardIni(clip);
    if (parsed.blocks.empty()) {
        LogWarn(L"CopyAlias: INI形式を解析できませんでした。表示対象（effect.name付きセクション）がありません。");
        return;
    }

    auto* state = new IniDialogState();
    state->parsed = std::move(parsed);
    state->targetObjects = std::move(capturedTargets);

    DialogBuilder db;
    db.Begin(
        L"プロパティ選択適用",
        260,
        180,
        WS_CAPTION | WS_SYSMENU | WS_VISIBLE | DS_SETFONT | DS_MODALFRAME | WS_CLIPCHILDREN);
    db.AddControl(L"STATIC", L"反映したいプロパティにチェックを入れて [適用] を押してください。", 1, 6, 6, 220, 10, SS_LEFT);
    db.AddControl(L"BUTTON", L"適用", IDC_INI_APPLY, 150, 155, 40, 14, BS_PUSHBUTTON);
    db.AddControl(L"BUTTON", L"閉じる", IDC_INI_CLOSE, 196, 155, 40, 14, BS_PUSHBUTTON);

    HWND dlg = CreateDialogIndirectParamW(GetModuleHandle(nullptr), db.Get(), nullptr, IniApplyDlgProc, (LPARAM)state);
    if (!dlg) {
        delete state;
        LogError(L"CopyAlias: ダイアログの作成に失敗しました。");
        return;
    }

    ShowWindow(dlg, SW_SHOW);
    SetForegroundWindow(dlg);
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

    INITCOMMONCONTROLSEX icc{};
    icc.dwSize = sizeof(icc);
    icc.dwICC = ICC_TREEVIEW_CLASSES;
    InitCommonControlsEx(&icc);

    host->set_plugin_information(L"CopyAlias 1.0.0");
    host->register_object_menu(L"エイリアスをコピー", CopySelectedObjectAliases);
    host->register_object_menu(L"エイリアスの値をペースト", OpenIniApplyDialog);

    g_edit = host->create_edit_handle();
}

extern "C" __declspec(dllexport) bool InitializePlugin(DWORD)
{
    return true;
}

extern "C" __declspec(dllexport) void InitializeLogger(LOG_HANDLE* logger)
{
    g_logger = logger;
    LogInfo(L"CopyAlias: Logger initialized.");
}

extern "C" __declspec(dllexport) void UninitializePlugin()
{
}
