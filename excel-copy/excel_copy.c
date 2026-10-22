/*
 * ExcelCopy - 엑셀 열 → 파워포인트 슬라이드 복사/붙여넣기 매크로
 *
 * 준비: 파워포인트에서 붙여넣을 위치를 클릭 → 엑셀에서 원하는 행의 셀을 클릭
 * 실행: Ctrl+Shift
 *
 *   1. 엑셀에서 선택된 셀의 행 번호를 읽는다.
 *   2. 파워포인트에서 마지막으로 클릭한 위치를 기억해 둔다.
 *   3. 슬라이드마다 입력한 열 순서로, [그 행의 셀부터 아래로 데이터 끝까지] 복사하고
 *      파워포인트 클릭 위치에 붙여넣은 뒤 → 1번.
 *      두 번째 슬라이드부터는 PageDown 후 같은 위치를 클릭하고 시작한다.
 *
 * 열 입력 (상태 창의 입력칸, exe 옆 ExcelCopy_columns.txt 에 자동 저장)
 *   한 줄 = 한 슬라이드, 열은 쉼표나 띄어쓰기로 구분. 예)
 *     D, AA, AX
 *     K, AH, BE
 *     S, AP, BM
 *
 * 엑셀 복사는 엑셀에 직접 연결(COM)해서 처리하므로 엑셀 창으로 전환하지 않는다.
 * (범위 = 그 행의 셀부터 Ctrl+Shift+↓ 와 같은 끝까지)
 *
 * Esc : 실행 중 중지
 * F9  : 일시정지 / 재개
 * 종료: 창의 [종료] 버튼 또는 창 닫기
 *
 * 실행할 때마다 exe와 같은 폴더에 ExcelCopy_log.txt 로 단계별 기록을 남긴다.
 *
 * 빌드 (MinGW-w64):
 *   x86_64-w64-mingw32-gcc -O2 -s -static -municode -mwindows \
 *       excel_copy.c -o ExcelCopy.exe -lole32 -loleaut32 -luuid
 */
#define WIN32_LEAN_AND_MEAN
#ifndef UNICODE
#define UNICODE
#endif
#ifndef _UNICODE
#define _UNICODE
#endif
#include <windows.h>
#include <ole2.h>
#include <oleauto.h>
#include <stdarg.h>

#define PPT_RIGHT_COUNT 1   /* 붙여넣기 후 파워포인트에서 → 횟수 */

#define MAX_PAGES 30        /* 입력할 수 있는 최대 슬라이드(줄) 수 */
#define MAX_COLS  30        /* 한 슬라이드에 입력할 수 있는 최대 열 수 */

#define DEFAULT_COLUMNS L"D, AA, AX\r\nK, AH, BE\r\nS, AP, BM"

/* 슬라이드별로 복사할 엑셀 열 (입력칸에서 읽음) */
typedef struct {
    int pages;
    int cols[MAX_PAGES];
    WCHAR name[MAX_PAGES][MAX_COLS][4];
    int error;              /* 잘못 입력한 줄 번호 (0 = 없음) */
} ColumnConfig;

static ColumnConfig g_cfg;
static CRITICAL_SECTION g_cfgLock;

#define XL_DOWN (-4121) /* Excel 상수 xlDown */

#define MACRO_SIGNATURE ((ULONG_PTR)0x58434350) /* 매크로가 보낸 입력 표시 */

#define WM_APP_UPDATE  (WM_APP + 2)

#define ID_BTN_PAUSE   101
#define ID_BTN_EXIT    102
#define ID_EDIT_COLS   103

/* 실행 결과 */
enum { RES_NONE, RES_DONE, RES_NO_EXCEL, RES_NO_PPT, RES_SWITCH_FAIL, RES_COPY_FAIL, RES_ABORTED,
       RES_NO_PPT_POS, RES_NO_ROW, RES_BAD_COLS, RES_MAX };

static HWND  g_wnd;
static HHOOK g_kbHook, g_mouseHook;
static HANDLE g_triggerEvent;           /* 작업 스레드를 깨우는 신호 */
static HWINEVENTHOOK g_fgHook;
static HFONT g_font, g_fontBold;
static HWND  g_edit;                    /* 열 입력칸 */

static HWND g_lastPpt;                  /* 마지막으로 활성화됐던 파워포인트 창 */
static volatile LONG g_busy;            /* 매크로 실행 중 여부 */
static volatile LONG g_abort;           /* Esc로 중지 요청 */
static volatile LONG g_paused;          /* 일시정지 여부 */
static volatile LONG g_result;          /* 최근 실행 결과 (RES_*) */
static volatile LONG g_page;            /* 현재 슬라이드 (1부터) */
static volatile LONG g_col;             /* 현재 슬라이드에서 몇 번째 열인지 (1부터) */
static volatile LONG g_row;             /* 엑셀에서 읽은 행 번호 */
static POINT g_pptClick;                /* 사용자가 파워포인트에서 마지막으로 클릭한 위치 */
static volatile LONG g_hasPptClick;

/* Ctrl+Shift 입력 감지 상태 */
static BOOL g_ctrlDown, g_shiftDown;
static BOOL g_fired;                    /* 이번 Ctrl+Shift 누름에서 이미 실행했음 (자동반복 방지) */

/* ---------------------------------------------------------------- 기록 */

static HANDLE g_log = INVALID_HANDLE_VALUE;
static DWORD g_logStart;

static void LogOpen(void)
{
    WCHAR path[MAX_PATH];
    WCHAR *slash, *c;
    DWORD w;

    GetModuleFileNameW(NULL, path, MAX_PATH);
    slash = path;
    for (c = path; *c; c++) if (*c == L'\\') slash = c + 1;
    lstrcpyW(slash, L"ExcelCopy_log.txt");
    g_log = CreateFileW(path, GENERIC_WRITE, FILE_SHARE_READ, NULL, CREATE_ALWAYS,
                        FILE_ATTRIBUTE_NORMAL, NULL);
    if (g_log != INVALID_HANDLE_VALUE)
        WriteFile(g_log, "\xEF\xBB\xBF", 3, &w, NULL); /* UTF-8 BOM (메모장 한글 표시용) */
    g_logStart = GetTickCount();
}

static void LogClose(void)
{
    if (g_log != INVALID_HANDLE_VALUE) CloseHandle(g_log);
    g_log = INVALID_HANDLE_VALUE;
}

static void Log(const WCHAR *fmt, ...)
{
    WCHAR buf[600];
    char utf8[1800];
    va_list ap;
    int len, n;
    DWORD w;

    if (g_log == INVALID_HANDLE_VALUE) return;
    len = wsprintfW(buf, L"[%6lu ms] ", GetTickCount() - g_logStart);
    va_start(ap, fmt);
    wvsprintfW(buf + len, fmt, ap);
    va_end(ap);
    lstrcatW(buf, L"\r\n");
    n = WideCharToMultiByte(CP_UTF8, 0, buf, -1, utf8, sizeof(utf8), NULL, NULL);
    if (n > 1) WriteFile(g_log, utf8, n - 1, &w, NULL);
}

/* ---------------------------------------------------------------- 열 입력 */

/* "D, AA, AX" 줄들을 읽어서 cfg에 채움. 잘못된 열 이름이 있으면 cfg->error에 줄 번호 */
static void ParseColumns(const WCHAR *text, ColumnConfig *cfg)
{
    int line = 1;
    const WCHAR *p = text;

    ZeroMemory(cfg, sizeof(*cfg));
    while (*p && cfg->pages < MAX_PAGES) {
        int n = 0;
        /* 한 줄 읽기 */
        while (*p && *p != L'\n') {
            WCHAR tok[8];
            int len = 0, bad = 0;
            while (*p == L' ' || *p == L'\t' || *p == L',' || *p == L'\r' || *p == 0x3000)
                p++;
            if (!*p || *p == L'\n') break;
            while (*p && *p != L' ' && *p != L'\t' && *p != L',' && *p != L'\r' && *p != L'\n'
                   && *p != 0x3000) {
                WCHAR ch = *p++;
                if (ch >= L'a' && ch <= L'z') ch -= 32;
                if (ch < L'A' || ch > L'Z') bad = 1;
                if (len < 7) tok[len] = ch;
                len++;
            }
            if (bad || len > 3) {
                if (!cfg->error) cfg->error = line;
                continue;
            }
            if (n < MAX_COLS) {
                CopyMemory(cfg->name[cfg->pages][n], tok, len * sizeof(WCHAR));
                cfg->name[cfg->pages][n][len] = 0;
                n++;
            }
        }
        if (*p == L'\n') p++;
        if (n > 0) cfg->cols[cfg->pages++] = n; /* 빈 줄은 건너뜀 */
        line++;
    }
}

static void ColumnsFilePath(WCHAR *path)
{
    WCHAR *slash, *c;
    GetModuleFileNameW(NULL, path, MAX_PATH);
    slash = path;
    for (c = path; *c; c++) if (*c == L'\\') slash = c + 1;
    lstrcpyW(slash, L"ExcelCopy_columns.txt");
}

/* 저장된 열 입력을 읽음 (없으면 기본값) */
static void LoadColumnsText(WCHAR *text, int max)
{
    WCHAR path[MAX_PATH];
    char buf[4096];
    DWORD got = 0;
    HANDLE f;

    lstrcpynW(text, DEFAULT_COLUMNS, max);
    ColumnsFilePath(path);
    f = CreateFileW(path, GENERIC_READ, FILE_SHARE_READ, NULL, OPEN_EXISTING, 0, NULL);
    if (f == INVALID_HANDLE_VALUE) return;
    if (ReadFile(f, buf, sizeof(buf) - 1, &got, NULL) && got > 0) {
        int skip = (got >= 3 && !memcmp(buf, "\xEF\xBB\xBF", 3)) ? 3 : 0; /* 메모장 BOM */
        int n = MultiByteToWideChar(CP_UTF8, 0, buf + skip, (int)got - skip, text, max - 1);
        text[n > 0 ? n : 0] = 0;
    }
    CloseHandle(f);
}

static void SaveColumnsText(const WCHAR *text)
{
    WCHAR path[MAX_PATH];
    char buf[4096];
    DWORD w;
    HANDLE f;
    int n = WideCharToMultiByte(CP_UTF8, 0, text, -1, buf, sizeof(buf), NULL, NULL);

    if (n <= 0) return;
    ColumnsFilePath(path);
    f = CreateFileW(path, GENERIC_WRITE, 0, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
    if (f == INVALID_HANDLE_VALUE) return;
    WriteFile(f, buf, n - 1, &w, NULL);
    CloseHandle(f);
}

/* 입력칸 내용이 바뀌면 다시 읽고 저장 */
static void UpdateColumnsFromEdit(void)
{
    WCHAR text[2048];
    ColumnConfig cfg;

    GetWindowTextW(g_edit, text, 2048);
    ParseColumns(text, &cfg);
    EnterCriticalSection(&g_cfgLock);
    g_cfg = cfg;
    LeaveCriticalSection(&g_cfgLock);
    SaveColumnsText(text);
}

/* ---------------------------------------------------------------- 창 */

static BOOL IsPptWindow(HWND hwnd)
{
    WCHAR cls[32];
    if (!hwnd || !IsWindow(hwnd)) return FALSE;
    if (!GetClassNameW(hwnd, cls, 32)) return FALSE;
    return lstrcmpW(cls, L"PPTFrameClass") == 0;
}

static HWND FindPpt(void)
{
    HWND fg = GetForegroundWindow();
    if (IsPptWindow(fg)) return fg;
    if (IsPptWindow(g_lastPpt)) return g_lastPpt;
    return FindWindowW(L"PPTFrameClass", NULL); /* 화면 맨 위의 파워포인트 창 */
}

/* 대상 프로그램이 지금까지 받은 입력을 모두 처리할 때까지 기다림.
 * 프로그램이 바쁘면 응답이 늦으므로, 연속 2번 바로 응답할 때까지 확인한다.
 * (한 번만 확인하면 붙여넣기·슬라이드 전환 도중에 응답해서 너무 일찍 넘어갈 수 있음) */
static void WaitIdle(HWND hwnd)
{
    DWORD_PTR r;
    DWORD t;
    int calm = 0, i;

    for (i = 0; i < 300 && calm < 2; i++) { /* 최대 약 5초 */
        Sleep(15);
        t = GetTickCount();
        if (!SendMessageTimeoutW(hwnd, WM_NULL, 0, 0, SMTO_NORMAL, 5000, &r))
            break;
        if (GetTickCount() - t < 20) calm++;
        else calm = 0;
    }
}

static BOOL ActivateWindow(HWND hwnd)
{
    HWND fg = GetForegroundWindow();
    DWORD fgThread, myThread = GetCurrentThreadId();
    int i;

    if (fg == hwnd) return TRUE;
    if (IsIconic(hwnd)) ShowWindow(hwnd, SW_RESTORE);

    fgThread = fg ? GetWindowThreadProcessId(fg, NULL) : 0;
    if (fgThread && fgThread != myThread) AttachThreadInput(myThread, fgThread, TRUE);
    BringWindowToTop(hwnd);
    SetForegroundWindow(hwnd);
    if (fgThread && fgThread != myThread) AttachThreadInput(myThread, fgThread, FALSE);

    for (i = 0; i < 20 && GetForegroundWindow() != hwnd; i++)
        Sleep(5);
    if (GetForegroundWindow() != hwnd) {
        SwitchToThisWindow(hwnd, TRUE); /* 일반 방법이 막혔을 때 대체 */
        for (i = 0; i < 100 && GetForegroundWindow() != hwnd; i++)
            Sleep(5);
    }
    if (GetForegroundWindow() != hwnd) return FALSE;
    WaitIdle(hwnd);
    return TRUE;
}

/* ---------------------------------------------------------------- 키보드/마우스 입력 */

static void AddKey(INPUT *in, int *n, WORD vk, BOOL up)
{
    INPUT *k = &in[(*n)++];
    ZeroMemory(k, sizeof(*k));
    k->type = INPUT_KEYBOARD;
    k->ki.wVk = vk;
    k->ki.wScan = (WORD)MapVirtualKeyW(vk, MAPVK_VK_TO_VSC);
    if (vk == VK_RIGHT || vk == VK_NEXT) k->ki.dwFlags |= KEYEVENTF_EXTENDEDKEY;
    if (up) k->ki.dwFlags |= KEYEVENTF_KEYUP;
    k->ki.dwExtraInfo = MACRO_SIGNATURE;
}

static void AddTap(INPUT *in, int *n, WORD vk)
{
    AddKey(in, n, vk, FALSE);
    AddKey(in, n, vk, TRUE);
}

static void ClickAt(HWND target, POINT pt)
{
    INPUT in[2];
    POINT old;

    GetCursorPos(&old);
    SetCursorPos(pt.x, pt.y);
    ZeroMemory(in, sizeof(in));
    in[0].type = in[1].type = INPUT_MOUSE;
    in[0].mi.dwFlags = MOUSEEVENTF_LEFTDOWN;
    in[1].mi.dwFlags = MOUSEEVENTF_LEFTUP;
    in[0].mi.dwExtraInfo = in[1].mi.dwExtraInfo = MACRO_SIGNATURE;
    SendInput(2, in, sizeof(INPUT));
    WaitIdle(target);
    SetCursorPos(old.x, old.y);
}

/* ---------------------------------------------------------------- 엑셀 (COM) */

/* 이름으로 속성/메서드 호출. 인자는 뒤에서부터 순서대로 넘긴다 (COM 규칙).
 * 엑셀이 바쁘면(셀 편집 중 등) 거절되므로 잠시 재시도한다. */
static HRESULT Invoke(IDispatch *obj, WORD flags, const WCHAR *name, VARIANT *result,
                      int argc, ...)
{
    DISPID id;
    DISPPARAMS dp = { NULL, NULL, 0, 0 };
    VARIANT args[4];
    LPOLESTR nm = (LPOLESTR)name;
    HRESULT hr;
    va_list ap;
    int i, tries;

    if (result) VariantInit(result);
    if (!obj) return E_POINTER;
    hr = obj->lpVtbl->GetIDsOfNames(obj, &IID_NULL, &nm, 1, LOCALE_USER_DEFAULT, &id);
    if (FAILED(hr)) return hr;

    va_start(ap, argc);
    for (i = 0; i < argc && i < 4; i++) args[i] = va_arg(ap, VARIANT);
    va_end(ap);
    dp.cArgs = argc;
    dp.rgvarg = args;

    for (tries = 0; tries < 60; tries++) { /* 최대 약 3초 */
        hr = obj->lpVtbl->Invoke(obj, id, &IID_NULL, LOCALE_USER_DEFAULT, flags, &dp,
                                 result, NULL, NULL);
        if (hr != RPC_E_CALL_REJECTED && hr != RPC_E_SERVERCALL_RETRYLATER) break;
        Sleep(50);
    }
    return hr;
}

static IDispatch *GetDisp(IDispatch *obj, const WCHAR *name)
{
    VARIANT v;
    if (FAILED(Invoke(obj, DISPATCH_PROPERTYGET, name, &v, 0))) return NULL;
    if (v.vt != VT_DISPATCH) { VariantClear(&v); return NULL; }
    return v.pdispVal;
}

static void Release(IDispatch *d)
{
    if (d) d->lpVtbl->Release(d);
}

static IDispatch *ConnectExcel(void)
{
    CLSID clsid;
    IUnknown *unk = NULL;
    IDispatch *app = NULL;

    if (FAILED(CLSIDFromProgID(L"Excel.Application", &clsid))) return NULL;
    if (FAILED(GetActiveObject(&clsid, NULL, &unk)) || !unk) return NULL;
    unk->lpVtbl->QueryInterface(unk, &IID_IDispatch, (void **)&app);
    unk->lpVtbl->Release(unk);
    return app;
}

/* 엑셀에서 선택된 셀의 행 번호 (실패하면 0) */
static LONG ExcelActiveRow(IDispatch *app)
{
    IDispatch *cell = GetDisp(app, L"ActiveCell");
    VARIANT v;
    LONG row = 0;

    if (cell && SUCCEEDED(Invoke(cell, DISPATCH_PROPERTYGET, L"Row", &v, 0))) {
        if (SUCCEEDED(VariantChangeType(&v, &v, 0, VT_I4))) row = v.lVal;
        VariantClear(&v);
    }
    Release(cell);
    return row;
}

/* 엑셀: [열][행] 셀부터 아래로 데이터 끝까지(Ctrl+Shift+↓ 와 같은 범위) 복사 */
static BOOL ExcelCopyColumn(IDispatch *app, const WCHAR *col, LONG row)
{
    IDispatch *sheet = GetDisp(app, L"ActiveSheet");
    IDispatch *first = NULL, *last = NULL, *range = NULL;
    WCHAR addr[32], rangeAddr[64] = L"?";
    VARIANT v, a, b;
    DWORD seq = GetClipboardSequenceNumber();
    BOOL ok = FALSE;

    wsprintfW(addr, L"%s%ld", col, row);
    if (!sheet) goto done;

    VariantInit(&a);
    a.vt = VT_BSTR;
    a.bstrVal = SysAllocString(addr);
    Invoke(sheet, DISPATCH_PROPERTYGET, L"Range", &v, 1, a);
    VariantClear(&a);
    if (v.vt != VT_DISPATCH) goto done;
    first = v.pdispVal;

    VariantInit(&a);
    a.vt = VT_I4;
    a.lVal = XL_DOWN;
    Invoke(first, DISPATCH_PROPERTYGET, L"End", &v, 1, a);
    if (v.vt != VT_DISPATCH) goto done;
    last = v.pdispVal;

    VariantInit(&a); a.vt = VT_DISPATCH; a.pdispVal = first;
    VariantInit(&b); b.vt = VT_DISPATCH; b.pdispVal = last;
    Invoke(sheet, DISPATCH_PROPERTYGET, L"Range", &v, 2, b, a); /* Range(first, last) */
    if (v.vt != VT_DISPATCH) goto done;
    range = v.pdispVal;

    if (SUCCEEDED(Invoke(range, DISPATCH_PROPERTYGET, L"Address", &v, 0))) {
        if (v.vt == VT_BSTR && v.bstrVal) lstrcpynW(rangeAddr, v.bstrVal, 64);
        VariantClear(&v);
    }

    if (FAILED(Invoke(range, DISPATCH_METHOD, L"Copy", NULL, 0))) goto done;
    ok = GetClipboardSequenceNumber() != seq;

done:
    Log(L"  엑셀 %s 부터 복사 %s (범위 %s)", addr, ok ? L"완료" : L"실패", rangeAddr);
    Release(range);
    Release(last);
    Release(first);
    Release(sheet);
    return ok;
}

/* ---------------------------------------------------------------- 파워포인트 */

/* Ctrl+V → → PPT_RIGHT_COUNT번 */
static void PptPaste(HWND ppt)
{
    INPUT in[6 + PPT_RIGHT_COUNT * 2];
    int n = 0, i;

    AddKey(in, &n, VK_SHIFT, TRUE);     /* Ctrl+Shift를 누르고 있어도 Ctrl+Shift+V가 되지 않게 */
    AddKey(in, &n, VK_CONTROL, FALSE);
    AddTap(in, &n, 'V');
    AddKey(in, &n, VK_CONTROL, TRUE);
    SendInput(n, in, sizeof(INPUT));
    /* 붙여넣기가 끝날 때까지 (다음 복사가 클립보드를 덮어쓰지 않도록) */
    WaitIdle(ppt);
    Sleep(50);
    WaitIdle(ppt);

    n = 0;
    for (i = 0; i < PPT_RIGHT_COUNT; i++)
        AddTap(in, &n, VK_RIGHT);
    SendInput(n, in, sizeof(INPUT));
    WaitIdle(ppt);
    Log(L"  PPT 붙여넣기 + → %d번 완료", PPT_RIGHT_COUNT);
}

/* 선택 해제 → PageDown (다음 슬라이드) */
static void PptNextSlide(HWND ppt)
{
    INPUT in[8];
    int n = 0;

    /* 셀 편집 → 셀 선택 → 표 선택 → 선택 없음 (남는 Esc는 아무 영향 없음) */
    AddTap(in, &n, VK_ESCAPE);
    AddTap(in, &n, VK_ESCAPE);
    AddTap(in, &n, VK_ESCAPE);
    SendInput(n, in, sizeof(INPUT));
    WaitIdle(ppt);

    n = 0;
    AddTap(in, &n, VK_NEXT);
    SendInput(n, in, sizeof(INPUT));
    WaitIdle(ppt);
    Sleep(50); /* 새 슬라이드가 화면에 그려질 시간 */
    WaitIdle(ppt);
    Log(L"  PPT Esc ×3 → PageDown 완료");
}

/* ---------------------------------------------------------------- 매크로 */

static void SetProgress(int page, int col)
{
    InterlockedExchange(&g_page, page);
    InterlockedExchange(&g_col, col);
    PostMessageW(g_wnd, WM_APP_UPDATE, 0, 0);
}

static ColumnConfig g_run;              /* 이번 실행에 쓰는 열 목록 (실행 중 입력칸을 고쳐도 영향 없음) */

static LONG RunMacro(void)
{
    HWND ppt = FindPpt();
    IDispatch *app;
    LONG row, res = RES_DONE;
    int page, c;

    EnterCriticalSection(&g_cfgLock);
    g_run = g_cfg;
    LeaveCriticalSection(&g_cfgLock);
    Log(L"열 입력: 슬라이드 %d개%s", g_run.pages, g_run.error ? L" (잘못된 입력 있음)" : L"");
    if (g_run.error || g_run.pages == 0) return RES_BAD_COLS;

    Log(L"시작: PPT 창 = %s, PPT 클릭 위치 = %s (%d, %d)",
        ppt ? L"있음" : L"없음", g_hasPptClick ? L"기록됨" : L"없음",
        (int)g_pptClick.x, (int)g_pptClick.y);
    if (!ppt) return RES_NO_PPT;
    if (!g_hasPptClick) return RES_NO_PPT_POS;

    app = ConnectExcel();
    if (!app) {
        Log(L"엑셀 연결 실패");
        return RES_NO_EXCEL;
    }

    /* 1. 선택된 셀의 행 번호 */
    row = ExcelActiveRow(app);
    InterlockedExchange(&g_row, row);
    Log(L"엑셀 선택 행 = %ld", row);
    if (row <= 0) { Release(app); return RES_NO_ROW; }

    if (!ActivateWindow(ppt)) { Release(app); return RES_SWITCH_FAIL; }

    for (page = 1; page <= g_run.pages && res == RES_DONE; page++) {
        Log(L"[슬라이드 %d]", page);
        if (page > 1) {
            if (g_abort) { res = RES_ABORTED; break; }
            if (!ActivateWindow(ppt)) { res = RES_SWITCH_FAIL; break; }
            PptNextSlide(ppt);
        }
        /* 붙여넣을 위치 클릭 (첫 슬라이드도 다시 클릭해서 위치를 확실히 잡음) */
        ClickAt(ppt, g_pptClick);
        Log(L"  PPT 위치 클릭 (%d, %d)", (int)g_pptClick.x, (int)g_pptClick.y);

        for (c = 0; c < g_run.cols[page - 1]; c++) {
            SetProgress(page, c + 1);
            if (g_abort) { res = RES_ABORTED; break; }
            if (!ExcelCopyColumn(app, g_run.name[page - 1][c], row)) { res = RES_COPY_FAIL; break; }

            if (g_abort) { res = RES_ABORTED; break; }
            if (!ActivateWindow(ppt)) { res = RES_SWITCH_FAIL; break; }
            PptPaste(ppt);
        }
    }

    Release(app);
    return res;
}

/* 실제 동작은 미리 띄워 둔 작업 스레드에서 수행한다.
 * (키보드 후크가 있는 메인 스레드를 막으면 시스템 전체 입력이 느려진다) */
static DWORD WINAPI MacroThread(LPVOID param)
{
    (void)param;
    CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);
    for (;;) {
        WaitForSingleObject(g_triggerEvent, INFINITE);
        InterlockedExchange(&g_abort, 0);
        LogOpen();
        InterlockedExchange(&g_result, RunMacro());
        Log(L"끝: 결과 코드 = %d", (int)g_result);
        LogClose();
        SetProgress(0, 0);
        InterlockedExchange(&g_busy, 0);
        PostMessageW(g_wnd, WM_APP_UPDATE, 0, 0);
    }
    return 0;
}

/* ---------------------------------------------------------------- 후크 */

static LRESULT CALLBACK KeyboardProc(int nCode, WPARAM wParam, LPARAM lParam)
{
    if (nCode == HC_ACTION) {
        const KBDLLHOOKSTRUCT *kb = (const KBDLLHOOKSTRUCT *)lParam;
        BOOL down = (wParam == WM_KEYDOWN || wParam == WM_SYSKEYDOWN);
        BOOL isCtrl, isShift;
        DWORD vk = kb->vkCode;

        /* 프로그램이 직접 보낸 키는 그대로 통과 */
        if ((kb->flags & LLKHF_INJECTED) || kb->dwExtraInfo == MACRO_SIGNATURE)
            return CallNextHookEx(g_kbHook, nCode, wParam, lParam);

        if (vk == VK_F9) {
            if (down)
                PostMessageW(g_wnd, WM_COMMAND, ID_BTN_PAUSE, 0);
            return 1;
        }

        if (vk == VK_ESCAPE && g_busy) {
            if (down)
                InterlockedExchange(&g_abort, 1);
            return 1;
        }

        isCtrl  = (vk == VK_LCONTROL || vk == VK_RCONTROL || vk == VK_CONTROL);
        isShift = (vk == VK_LSHIFT || vk == VK_RSHIFT || vk == VK_SHIFT);
        if (!isCtrl && !isShift)
            return CallNextHookEx(g_kbHook, nCode, wParam, lParam);

        if (isCtrl) g_ctrlDown = down;
        else        g_shiftDown = down;

        if (!g_ctrlDown || !g_shiftDown) {
            g_fired = FALSE; /* 둘 중 하나라도 떼면 다시 실행 가능 */
        } else if (!g_fired) {
            /* Ctrl, Shift 둘 다 눌린 순간 바로 실행 */
            g_fired = TRUE;
            if (!g_paused && InterlockedCompareExchange(&g_busy, 1, 0) == 0) {
                SetEvent(g_triggerEvent);
                return CallNextHookEx(g_kbHook, nCode, wParam, lParam);
            }
        }

        /* 실행 중에는 Ctrl/Shift를 계속 누르고 있어서 생기는 자동반복 입력을 막는다.
         * (막지 않으면 → 키가 Ctrl+→ 로 바뀔 수 있음) */
        if (down && g_busy)
            return 1;
    }
    return CallNextHookEx(g_kbHook, nCode, wParam, lParam);
}

/* 사용자가 파워포인트 창 안을 클릭하면 그 위치를 기록 */
static LRESULT CALLBACK MouseProc(int nCode, WPARAM wParam, LPARAM lParam)
{
    if (nCode == HC_ACTION && wParam == WM_LBUTTONDOWN && !g_busy) {
        const MSLLHOOKSTRUCT *ms = (const MSLLHOOKSTRUCT *)lParam;
        if (!(ms->flags & LLMHF_INJECTED) && ms->dwExtraInfo != MACRO_SIGNATURE) {
            HWND root = GetAncestor(WindowFromPoint(ms->pt), GA_ROOT);
            if (IsPptWindow(root)) {
                g_pptClick = ms->pt;
                InterlockedExchange(&g_hasPptClick, 1);
                PostMessageW(g_wnd, WM_APP_UPDATE, 0, 0);
            }
        }
    }
    return CallNextHookEx(g_mouseHook, nCode, wParam, lParam);
}

static void CALLBACK ForegroundProc(HWINEVENTHOOK hook, DWORD event, HWND hwnd,
                                    LONG idObject, LONG idChild,
                                    DWORD thread, DWORD time)
{
    (void)hook; (void)event; (void)idObject; (void)idChild; (void)thread; (void)time;
    if (IsPptWindow(hwnd)) {
        g_lastPpt = hwnd;
        InvalidateRect(g_wnd, NULL, TRUE);
    }
}

/* ---------------------------------------------------------------- 화면 */

#define EDIT_Y   158  /* 열 입력칸 위치/높이 */
#define EDIT_H   110
#define BTN_Y    (EDIT_Y + EDIT_H + 30)

static void DrawLine(HDC dc, int *y, const WCHAR *text)
{
    RECT rc;
    SetRect(&rc, 12, *y, 420, *y + 20);
    DrawTextW(dc, text, -1, &rc, DT_LEFT | DT_SINGLELINE | DT_NOPREFIX | DT_END_ELLIPSIS);
    *y += 22;
}

static void Paint(HWND hwnd)
{
    static const WCHAR *results[RES_MAX] = {
        L"PPT 위치 클릭 → 엑셀 셀 클릭 → Ctrl+Shift",
        L"최근 결과:  모든 슬라이드 완료",
        L"최근 결과:  엑셀에 연결하지 못했습니다",
        L"최근 결과:  파워포인트 창을 찾지 못했습니다",
        L"최근 결과:  창 전환에 실패했습니다",
        L"최근 결과:  엑셀 복사에 실패했습니다 (셀 편집 중인지 확인)",
        L"최근 결과:  Esc로 중지했습니다",
        L"최근 결과:  PPT에서 붙여넣을 위치를 먼저 클릭하세요",
        L"최근 결과:  엑셀 선택 셀의 행 번호를 읽지 못했습니다",
        L"최근 결과:  열 입력을 확인하세요",
    };
    PAINTSTRUCT ps;
    HDC dc = BeginPaint(hwnd, &ps);
    WCHAR title[128];
    WCHAR line[256];
    HWND w;
    LONG res = g_result, page = g_page, col = g_col;
    ColumnConfig *cfg = &g_cfg;
    int y = 10;

    if (g_busy) cfg = &g_run; /* 실행 중에는 실행에 쓰는 목록으로 표시 */

    SetBkMode(dc, TRANSPARENT);

    SelectObject(dc, g_fontBold);
    if (g_paused) {
        SetTextColor(dc, RGB(200, 60, 40));
        lstrcpyW(line, L"● 일시정지됨  (F9: 재개)");
    } else if (g_busy) {
        SetTextColor(dc, RGB(30, 90, 200));
        if (page > 0 && page <= cfg->pages && col > 0 && col <= cfg->cols[page - 1])
            wsprintfW(line, L"● 실행 중  슬라이드 %d/%d · %s열  (Esc: 중지)",
                      (int)page, cfg->pages, cfg->name[page - 1][col - 1]);
        else
            lstrcpyW(line, L"● 실행 중  (Esc: 중지)");
    } else {
        SetTextColor(dc, RGB(30, 140, 60));
        lstrcpyW(line, L"● 대기 중 - Ctrl+Shift 누르면 실행  (F9: 일시정지)");
    }
    DrawLine(dc, &y, line);
    y += 4;

    SelectObject(dc, g_font);
    SetTextColor(dc, RGB(20, 20, 20));
    if (res < 0 || res >= RES_MAX) res = RES_NONE;
    DrawLine(dc, &y, results[res]);

    if (g_row > 0)
        wsprintfW(line, L"엑셀 행:  %ld", g_row);
    else
        lstrcpyW(line, L"엑셀 행:  실행할 때 선택된 셀에서 읽음");
    DrawLine(dc, &y, line);

    title[0] = 0;
    w = FindPpt();
    if (w) GetWindowTextW(w, title, 128);
    wsprintfW(line, L"파워포인트:  %s", w ? (title[0] ? title : L"PowerPoint") : L"(없음)");
    DrawLine(dc, &y, line);

    if (g_hasPptClick)
        wsprintfW(line, L"PPT 클릭 위치:  기록됨 (%d, %d)", (int)g_pptClick.x, (int)g_pptClick.y);
    else
        lstrcpyW(line, L"PPT 클릭 위치:  없음 - PPT에서 붙여넣을 칸을 클릭하세요");
    DrawLine(dc, &y, line);

    /* 열 입력 안내 */
    y += 6;
    SelectObject(dc, g_fontBold);
    DrawLine(dc, &y, L"복사할 열  (한 줄 = 한 슬라이드, 쉼표로 구분)");
    SelectObject(dc, g_font);
    y = EDIT_Y + EDIT_H + 4;
    if (cfg->error) {
        SetTextColor(dc, RGB(200, 60, 40));
        wsprintfW(line, L"%d번째 줄에 잘못된 열 이름이 있습니다 (영문 1~3글자, 예: D, AA)", cfg->error);
    } else if (cfg->pages == 0) {
        SetTextColor(dc, RGB(200, 60, 40));
        lstrcpyW(line, L"열을 입력하세요 (예: D, AA, AX)");
    } else {
        int i, total = 0;
        for (i = 0; i < cfg->pages; i++) total += cfg->cols[i];
        SetTextColor(dc, RGB(90, 90, 90));
        wsprintfW(line, L"슬라이드 %d개 · 열 %d개  (자동 저장됨)", cfg->pages, total);
    }
    DrawLine(dc, &y, line);

    EndPaint(hwnd, &ps);
}

static LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
    switch (msg) {
    case WM_CREATE: {
        WCHAR text[2048];
        HWND b1 = CreateWindowW(L"BUTTON", L"일시정지 (F9)", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
                                12, BTN_Y, 194, 30, hwnd, (HMENU)ID_BTN_PAUSE, NULL, NULL);
        HWND b2 = CreateWindowW(L"BUTTON", L"종료", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
                                218, BTN_Y, 194, 30, hwnd, (HMENU)ID_BTN_EXIT, NULL, NULL);
        SendMessageW(b1, WM_SETFONT, (WPARAM)g_font, TRUE);
        SendMessageW(b2, WM_SETFONT, (WPARAM)g_font, TRUE);

        LoadColumnsText(text, 2048);
        g_edit = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", text,
                                 WS_CHILD | WS_VISIBLE | WS_VSCROLL | ES_MULTILINE |
                                 ES_AUTOVSCROLL | ES_WANTRETURN | ES_UPPERCASE,
                                 12, EDIT_Y, 400, EDIT_H, hwnd, (HMENU)ID_EDIT_COLS, NULL, NULL);
        SendMessageW(g_edit, WM_SETFONT, (WPARAM)g_font, TRUE);
        UpdateColumnsFromEdit();
        SetTimer(hwnd, 1, 1000, NULL); /* 창 이름 표시 갱신 */
        return 0;
    }
    case WM_COMMAND:
        switch (LOWORD(wParam)) {
        case ID_BTN_PAUSE:
            InterlockedExchange(&g_paused, g_paused ? 0 : 1);
            SetDlgItemTextW(hwnd, ID_BTN_PAUSE, g_paused ? L"재개 (F9)" : L"일시정지 (F9)");
            InvalidateRect(hwnd, NULL, TRUE);
            return 0;
        case ID_BTN_EXIT:
            DestroyWindow(hwnd);
            return 0;
        case ID_EDIT_COLS:
            if (HIWORD(wParam) == EN_CHANGE) {
                UpdateColumnsFromEdit();
                InvalidateRect(hwnd, NULL, TRUE);
            }
            return 0;
        }
        break;
    case WM_APP_UPDATE:
    case WM_TIMER:
        InvalidateRect(hwnd, NULL, TRUE);
        return 0;
    case WM_MOUSEACTIVATE: {
        /* 열 입력칸을 클릭했을 때만 창을 활성화해서 글자를 입력할 수 있게 하고,
         * 나머지(버튼 등)는 클릭해도 작업 중인 창이 계속 활성 상태로 남도록 함 */
        POINT pt;
        GetCursorPos(&pt);
        ScreenToClient(hwnd, &pt);
        if (ChildWindowFromPoint(hwnd, pt) == g_edit)
            return MA_ACTIVATE;
        return MA_NOACTIVATE;
    }
    case WM_PAINT:
        Paint(hwnd);
        return 0;
    case WM_DESTROY:
        KillTimer(hwnd, 1);
        PostQuitMessage(0);
        return 0;
    }
    return DefWindowProcW(hwnd, msg, wParam, lParam);
}

int WINAPI wWinMain(HINSTANCE inst, HINSTANCE prev, PWSTR cmd, int show)
{
    WNDCLASSW wc;
    MSG msg;
    RECT rc = { 0, 0, 424, BTN_Y + 42 };
    DWORD style = WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_MINIMIZEBOX;
    DWORD exStyle = WS_EX_TOPMOST | WS_EX_APPWINDOW; /* 활성화 여부는 WM_MOUSEACTIVATE에서 결정 */
    HANDLE mutex;
    (void)prev; (void)cmd; (void)show;

    InitializeCriticalSection(&g_cfgLock);

    /* 중복 실행 방지 */
    mutex = CreateMutexW(NULL, TRUE, L"ExcelCopy_SingleInstance");
    if (GetLastError() == ERROR_ALREADY_EXISTS) {
        MessageBoxW(NULL, L"ExcelCopy가 이미 실행 중입니다.", L"ExcelCopy", MB_OK | MB_ICONINFORMATION);
        return 0;
    }

    /* 화면 배율(DPI)과 관계없이 클릭 좌표가 정확하도록 (모니터별 배율이 달라도) */
    {
        typedef BOOL (WINAPI *SetCtxFn)(HANDLE);
        SetCtxFn setCtx = (SetCtxFn)(void *)GetProcAddress(GetModuleHandleW(L"user32.dll"),
                                                           "SetProcessDpiAwarenessContext");
        if (!setCtx || !setCtx((HANDLE)-4)) /* PER_MONITOR_AWARE_V2 */
            SetProcessDPIAware();
    }

    if (IsPptWindow(GetForegroundWindow()))
        g_lastPpt = GetForegroundWindow();

    g_font = CreateFontW(-15, 0, 0, 0, FW_NORMAL, 0, 0, 0, DEFAULT_CHARSET, 0, 0,
                         CLEARTYPE_QUALITY, 0, L"Malgun Gothic");
    g_fontBold = CreateFontW(-15, 0, 0, 0, FW_BOLD, 0, 0, 0, DEFAULT_CHARSET, 0, 0,
                             CLEARTYPE_QUALITY, 0, L"Malgun Gothic");

    ZeroMemory(&wc, sizeof(wc));
    wc.lpfnWndProc = WndProc;
    wc.hInstance = inst;
    wc.hCursor = LoadCursor(NULL, IDC_ARROW);
    wc.hIcon = LoadIcon(NULL, IDI_APPLICATION);
    wc.hbrBackground = (HBRUSH)(COLOR_BTNFACE + 1);
    wc.lpszClassName = L"ExcelCopyWnd";
    RegisterClassW(&wc);

    AdjustWindowRectEx(&rc, style, FALSE, exStyle);
    g_wnd = CreateWindowExW(exStyle, wc.lpszClassName, L"ExcelCopy - 엑셀 열 → 파워포인트 붙여넣기", style,
                            CW_USEDEFAULT, CW_USEDEFAULT,
                            rc.right - rc.left, rc.bottom - rc.top,
                            NULL, NULL, inst, NULL);
    if (!g_wnd) return 1;

    g_triggerEvent = CreateEventW(NULL, FALSE, FALSE, NULL);
    if (g_triggerEvent) {
        HANDLE th = CreateThread(NULL, 0, MacroThread, NULL, 0, NULL);
        if (th) {
            SetThreadPriority(th, THREAD_PRIORITY_HIGHEST);
            CloseHandle(th);
        }
    }

    g_kbHook = SetWindowsHookExW(WH_KEYBOARD_LL, KeyboardProc, inst, 0);
    if (!g_kbHook) {
        MessageBoxW(NULL, L"키보드 후크 설치에 실패했습니다.", L"ExcelCopy", MB_OK | MB_ICONERROR);
        return 1;
    }
    g_mouseHook = SetWindowsHookExW(WH_MOUSE_LL, MouseProc, inst, 0);
    g_fgHook = SetWinEventHook(EVENT_SYSTEM_FOREGROUND, EVENT_SYSTEM_FOREGROUND, NULL,
                               ForegroundProc, 0, 0, WINEVENT_OUTOFCONTEXT);

    /* 창을 띄우되 포커스는 빼앗지 않음 */
    ShowWindow(g_wnd, SW_SHOWNOACTIVATE);
    UpdateWindow(g_wnd);

    while (GetMessageW(&msg, NULL, 0, 0) > 0) {
        TranslateMessage(&msg);
        DispatchMessageW(&msg);
    }

    if (g_fgHook) UnhookWinEvent(g_fgHook);
    if (g_mouseHook) UnhookWindowsHookEx(g_mouseHook);
    UnhookWindowsHookEx(g_kbHook);
    DeleteObject(g_font);
    DeleteObject(g_fontBold);
    if (mutex) CloseHandle(mutex);
    return 0;
}
