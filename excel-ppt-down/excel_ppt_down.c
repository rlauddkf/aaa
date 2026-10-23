/*
 * ExcelPptDown - 엑셀 열 복사 → 파워포인트 붙여넣기 (아래로) 반복 매크로
 *
 * 준비: 파워포인트에서 붙여넣을 위치를 클릭(기록됨) → 엑셀에서 시작할 셀을 클릭
 * 실행: Ctrl + 숫자패드 -
 *
 * 페이지마다 (페이지 수는 상태 창에서 입력):
 *   0. 파워포인트: 기록한 위치 클릭 (2페이지부터는 Esc → PageDown 후 클릭)
 *   페이지당 횟수(기본 6)만큼 반복:
 *     1. 엑셀:       선택된 셀부터 아래로 데이터 끝까지(Ctrl+Shift+↓ 범위) 복사
 *     2. 파워포인트: 현재 위치에 붙여넣기 (Ctrl+V) → ↓ 1번
 *     3. 엑셀:       → 1번 (다음 열)   ※ 맨 마지막 회차는 생략
 * 엑셀 시작 셀이 비어 있으면 거기서 멈춘다 (데이터 끝).
 *
 * 엑셀은 직접 연결(COM)해서 처리하므로 엑셀 창으로 전환하지 않는다.
 * (파워포인트 창이 계속 앞에 있어서 빠름)
 *
 * Esc : 실행 중 중지
 * 종료: 창의 [종료] 버튼 또는 창 닫기
 *
 * 실행할 때마다 exe와 같은 폴더에 ExcelPptDown_log.txt 로 단계별 기록을 남긴다.
 *
 * 빌드 (MinGW-w64):
 *   x86_64-w64-mingw32-gcc -O2 -s -static -municode -mwindows \
 *       excel_ppt_down.c -o ExcelPptDown.exe -lole32 -loleaut32 -luuid
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

#define DEFAULT_PER_PAGE 6   /* 페이지당 반복 횟수 기본값 */
#define DEFAULT_PAGES    3   /* 페이지 수 기본값 */
#define MAX_NUMBER       99

#define XL_DOWN (-4121)      /* Excel 상수 xlDown */

#define MACRO_SIGNATURE ((ULONG_PTR)0x58505044) /* 매크로가 보낸 키 표시 */

#define WM_APP_UPDATE  (WM_APP + 2)
#define ID_BTN_EXIT    102
#define ID_EDIT_PAGES  103
#define ID_EDIT_PER    104

/* 실행 결과 */
enum { RES_NONE, RES_DONE, RES_NO_EXCEL, RES_NO_PPT, RES_SWITCH_FAIL, RES_COPY_FAIL,
       RES_MOVE_FAIL, RES_ABORTED, RES_NO_PPT_POS, RES_EMPTY, RES_MAX };

static HWND  g_wnd;
static HHOOK g_kbHook, g_mouseHook;
static HANDLE g_triggerEvent;
static HWINEVENTHOOK g_fgHook;
static HFONT g_font, g_fontBold;

static HWND g_lastPpt;                  /* 마지막으로 활성화됐던 파워포인트 창 */
static volatile LONG g_busy;
static volatile LONG g_abort;
static volatile LONG g_result;
static volatile LONG g_round;           /* 현재 페이지 안의 회차 */
static volatile LONG g_page;            /* 현재 페이지 */
static volatile LONG g_pages = DEFAULT_PAGES;      /* 페이지 수 (입력칸) */
static volatile LONG g_perPage = DEFAULT_PER_PAGE; /* 페이지당 횟수 (입력칸) */
static HWND g_editPages, g_editPer;
static POINT g_pptClick;                /* 사용자가 파워포인트에서 마지막으로 클릭한 위치 */
static volatile LONG g_hasPptClick;
static BOOL g_minusDown;                /* 숫자패드 - 눌림 상태 (자동반복 방지) */

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
    lstrcpyW(slash, L"ExcelPptDown_log.txt");
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

/* ---------------------------------------------------------------- 창 / 입력 */

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
    return FindWindowW(L"PPTFrameClass", NULL);
}

/* 대상 프로그램이 지금까지 받은 입력을 모두 처리할 때까지 기다림
 * (연속 2번 바로 응답할 때까지 확인) */
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

    if (!hwnd) return FALSE;
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
        SwitchToThisWindow(hwnd, TRUE);
        for (i = 0; i < 100 && GetForegroundWindow() != hwnd; i++)
            Sleep(5);
    }
    if (GetForegroundWindow() != hwnd) return FALSE;
    WaitIdle(hwnd);
    return TRUE;
}

static void AddKey(INPUT *in, int *n, WORD vk, BOOL up)
{
    INPUT *k = &in[(*n)++];
    ZeroMemory(k, sizeof(*k));
    k->type = INPUT_KEYBOARD;
    k->ki.wVk = vk;
    k->ki.wScan = (WORD)MapVirtualKeyW(vk, MAPVK_VK_TO_VSC);
    if (vk == VK_DOWN || vk == VK_NEXT) k->ki.dwFlags |= KEYEVENTF_EXTENDEDKEY;
    if (up) k->ki.dwFlags |= KEYEVENTF_KEYUP;
    k->ki.dwExtraInfo = MACRO_SIGNATURE;
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

/* ---------------------------------------------------------------- 설정 (페이지 수 / 페이지당 횟수) */

static void SettingsPath(WCHAR *path)
{
    WCHAR *slash, *c;
    GetModuleFileNameW(NULL, path, MAX_PATH);
    slash = path;
    for (c = path; *c; c++) if (*c == L'\\') slash = c + 1;
    lstrcpyW(slash, L"ExcelPptDown_settings.txt");
}

static void LoadSettings(void)
{
    WCHAR path[MAX_PATH];
    char buf[64];
    DWORD got = 0;
    HANDLE f;
    int a = 0, b = 0;
    char *p;

    SettingsPath(path);
    f = CreateFileW(path, GENERIC_READ, FILE_SHARE_READ, NULL, OPEN_EXISTING, 0, NULL);
    if (f == INVALID_HANDLE_VALUE) return;
    if (ReadFile(f, buf, sizeof(buf) - 1, &got, NULL)) {
        buf[got] = 0;
        p = buf;
        while (*p >= '0' && *p <= '9') a = a * 10 + (*p++ - '0');
        while (*p && (*p < '0' || *p > '9')) p++;
        while (*p >= '0' && *p <= '9') b = b * 10 + (*p++ - '0');
        if (a >= 1 && a <= MAX_NUMBER) g_pages = a;
        if (b >= 1 && b <= MAX_NUMBER) g_perPage = b;
    }
    CloseHandle(f);
}

static void SaveSettings(void)
{
    WCHAR path[MAX_PATH];
    char buf[64];
    DWORD w;
    HANDLE f;
    int n = wsprintfA(buf, "%ld %ld\r\n", g_pages, g_perPage);

    SettingsPath(path);
    f = CreateFileW(path, GENERIC_WRITE, 0, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
    if (f == INVALID_HANDLE_VALUE) return;
    WriteFile(f, buf, n, &w, NULL);
    CloseHandle(f);
}

/* 입력칸 숫자 읽기 (1~99, 잘못되면 0) */
static LONG ReadNumber(HWND edit)
{
    WCHAR t[16];
    LONG n = 0;
    int i;
    GetWindowTextW(edit, t, 16);
    for (i = 0; t[i]; i++) {
        if (t[i] < L'0' || t[i] > L'9') return 0;
        n = n * 10 + (t[i] - L'0');
        if (n > MAX_NUMBER) return 0;
    }
    return n;
}

/* ---------------------------------------------------------------- 엑셀 (COM) */

/* 이름으로 속성/메서드 호출. 인자는 뒤에서부터 순서대로 넘긴다 (COM 규칙).
 * 엑셀이 바쁘면(셀 편집 중 등) 거절되므로 잠시 재시도한다. */
static HRESULT Call(IDispatch *obj, WORD flags, const WCHAR *name, VARIANT *result,
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

static VARIANT VInt(LONG n) { VARIANT v; VariantInit(&v); v.vt = VT_I4; v.lVal = n; return v; }

static void Release(IDispatch *d)
{
    if (d) d->lpVtbl->Release(d);
}

static IDispatch *Get(IDispatch *obj, const WCHAR *name, int argc, VARIANT a, VARIANT b)
{
    VARIANT v;
    HRESULT hr = argc == 0 ? Call(obj, DISPATCH_PROPERTYGET, name, &v, 0)
               : argc == 1 ? Call(obj, DISPATCH_PROPERTYGET, name, &v, 1, a)
                           : Call(obj, DISPATCH_PROPERTYGET, name, &v, 2, b, a);
    if (FAILED(hr)) return NULL;
    if (v.vt != VT_DISPATCH) { VariantClear(&v); return NULL; }
    return v.pdispVal;
}

static VARIANT VNone(void) { VARIANT v; VariantInit(&v); return v; }
#define GET0(obj, name)       Get(obj, name, 0, VNone(), VNone())
#define GET1(obj, name, a)    Get(obj, name, 1, a, VNone())
#define GET2(obj, name, a, b) Get(obj, name, 2, a, b)

static LONG GetInt(IDispatch *obj, const WCHAR *name)
{
    VARIANT v;
    LONG n = 0;
    if (SUCCEEDED(Call(obj, DISPATCH_PROPERTYGET, name, &v, 0))) {
        if (SUCCEEDED(VariantChangeType(&v, &v, 0, VT_I4))) n = v.lVal;
        VariantClear(&v);
    }
    return n;
}

static void GetAddress(IDispatch *range, WCHAR *out, int max)
{
    VARIANT v;
    lstrcpynW(out, L"?", max);
    if (range && SUCCEEDED(Call(range, DISPATCH_PROPERTYGET, L"Address", &v, 0))) {
        if (v.vt == VT_BSTR && v.bstrVal) lstrcpynW(out, v.bstrVal, max);
        VariantClear(&v);
    }
}

/* 선택된 셀이 비어 있는지 */
static BOOL ExcelActiveCellEmpty(IDispatch *app)
{
    IDispatch *cell = GET0(app, L"ActiveCell");
    VARIANT v;
    BOOL empty = FALSE;
    if (cell && SUCCEEDED(Call(cell, DISPATCH_PROPERTYGET, L"Value", &v, 0))) {
        empty = (v.vt == VT_EMPTY) || (v.vt == VT_BSTR && (!v.bstrVal || !v.bstrVal[0]));
        VariantClear(&v);
    }
    Release(cell);
    return empty;
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

/* 엑셀: 선택된 셀부터 아래로 데이터 끝까지(Ctrl+Shift+↓ 와 같은 범위) 복사 */
static BOOL ExcelCopyDown(IDispatch *app)
{
    IDispatch *sheet = GET0(app, L"ActiveSheet");
    IDispatch *first = GET0(app, L"ActiveCell");
    IDispatch *last = NULL, *range = NULL;
    VARIANT a, b;
    WCHAR addr[64] = L"?";
    DWORD seq = GetClipboardSequenceNumber();
    BOOL ok = FALSE;

    if (!sheet || !first) goto done;
    last = GET1(first, L"End", VInt(XL_DOWN));
    if (!last) goto done;
    VariantInit(&a); a.vt = VT_DISPATCH; a.pdispVal = first;
    VariantInit(&b); b.vt = VT_DISPATCH; b.pdispVal = last;
    range = GET2(sheet, L"Range", a, b); /* Range(first, last) */
    if (!range) goto done;
    GetAddress(range, addr, 64);
    if (FAILED(Call(range, DISPATCH_METHOD, L"Copy", NULL, 0))) goto done;
    ok = GetClipboardSequenceNumber() != seq;

done:
    Log(L"  엑셀 복사 %s (범위 %s)", ok ? L"완료" : L"실패", addr);
    Release(range);
    Release(last);
    Release(first);
    Release(sheet);
    return ok;
}

/* 엑셀: 선택 셀을 오른쪽으로 한 칸 (→ 키처럼 숨긴 열은 건너뜀) */
static BOOL ExcelMoveRight(IDispatch *app)
{
    IDispatch *cell = GET0(app, L"ActiveCell");
    IDispatch *next = NULL;
    WCHAR addr[64] = L"?";
    BOOL ok = FALSE;
    int i;

    for (i = 0; cell && i < 200; i++) { /* 숨긴 열이 이어져 있으면 계속 오른쪽으로 */
        IDispatch *col;
        next = GET2(cell, L"Offset", VInt(0), VInt(1));
        Release(cell);
        cell = next;
        if (!cell) break;
        col = GET0(cell, L"EntireColumn");
        if (!col || !GetInt(col, L"Hidden")) { Release(col); break; }
        Release(col);
    }
    if (cell && SUCCEEDED(Call(cell, DISPATCH_METHOD, L"Select", NULL, 0)))
        ok = TRUE;
    GetAddress(cell, addr, 64);
    Log(L"  엑셀 → 1칸 %s (선택 셀 %s)", ok ? L"완료" : L"실패", addr);
    Release(cell);
    return ok;
}

/* ---------------------------------------------------------------- 파워포인트 */

/* Ctrl+V → ↓ 1번 */
static void PptPasteDown(HWND ppt)
{
    INPUT in[6];
    int n = 0;

    AddKey(in, &n, VK_CONTROL, FALSE);
    AddKey(in, &n, 'V', FALSE);
    AddKey(in, &n, 'V', TRUE);
    AddKey(in, &n, VK_CONTROL, TRUE);
    SendInput(n, in, sizeof(INPUT));
    /* 붙여넣기가 끝날 때까지 (다음 복사가 클립보드를 덮어쓰지 않도록) */
    WaitIdle(ppt);
    Sleep(50);
    WaitIdle(ppt);

    n = 0;
    AddKey(in, &n, VK_DOWN, FALSE);
    AddKey(in, &n, VK_DOWN, TRUE);
    SendInput(n, in, sizeof(INPUT));
    WaitIdle(ppt);
    Log(L"  PPT 붙여넣기 → ↓ 1번 완료");
}

/* 선택 해제 → PageDown (다음 슬라이드) */
static void PptNextSlide(HWND ppt)
{
    INPUT in[6];
    int n = 0, i;

    /* 셀 편집 → 셀 선택 → 표 선택 → 선택 없음 (남는 Esc는 아무 영향 없음) */
    for (i = 0; i < 3; i++) {
        AddKey(in, &n, VK_ESCAPE, FALSE);
        AddKey(in, &n, VK_ESCAPE, TRUE);
    }
    SendInput(n, in, sizeof(INPUT));
    WaitIdle(ppt);

    n = 0;
    AddKey(in, &n, VK_NEXT, FALSE);
    AddKey(in, &n, VK_NEXT, TRUE);
    SendInput(n, in, sizeof(INPUT));
    WaitIdle(ppt);
    Sleep(50); /* 새 슬라이드가 화면에 그려질 시간 */
    WaitIdle(ppt);
    Log(L"  PPT Esc ×3 → PageDown 완료");
}

/* ---------------------------------------------------------------- 매크로 */

static void SetProgress(int page, int r)
{
    InterlockedExchange(&g_page, page);
    InterlockedExchange(&g_round, r);
    PostMessageW(g_wnd, WM_APP_UPDATE, 0, 0);
}

static LONG RunMacro(void)
{
    HWND ppt = FindPpt();
    IDispatch *app;
    LONG res = RES_DONE, pages = g_pages, per = g_perPage;
    int page, r;

    Log(L"시작: 페이지 %ld × %ld회, 파워포인트 창 %s, PPT 클릭 위치 %s (%d, %d)",
        pages, per, ppt ? L"있음" : L"없음", g_hasPptClick ? L"기록됨" : L"없음",
        (int)g_pptClick.x, (int)g_pptClick.y);
    if (!ppt) return RES_NO_PPT;
    if (!g_hasPptClick) return RES_NO_PPT_POS;
    app = ConnectExcel();
    Log(L"엑셀 연결 %s", app ? L"성공" : L"실패");
    if (!app) return RES_NO_EXCEL;

    if (!ActivateWindow(ppt)) { Release(app); return RES_SWITCH_FAIL; }

    for (page = 1; page <= pages && res == RES_DONE; page++) {
        Log(L"[페이지 %d]", page);
        if (page > 1) {
            if (g_abort) { res = RES_ABORTED; break; }
            if (!ActivateWindow(ppt)) { res = RES_SWITCH_FAIL; break; }
            PptNextSlide(ppt);
        }
        /* 기록한 위치 찍기 (첫 페이지도 다시 찍어서 위치를 확실히 잡음) */
        ClickAt(ppt, g_pptClick);
        Log(L"  PPT 위치 클릭 (%d, %d)", (int)g_pptClick.x, (int)g_pptClick.y);

        for (r = 1; r <= per; r++) {
            SetProgress(page, r);
            Log(L" [페이지 %d - %d회차]", page, r);
            if (g_abort) { res = RES_ABORTED; break; }
            if (ExcelActiveCellEmpty(app)) {
                Log(L"  엑셀 시작 셀이 비어 있어서 멈춤");
                res = RES_EMPTY;
                break;
            }
            if (!ExcelCopyDown(app)) { res = RES_COPY_FAIL; break; }

            if (g_abort) { res = RES_ABORTED; break; }
            if (!ActivateWindow(ppt)) { res = RES_SWITCH_FAIL; break; }
            PptPasteDown(ppt);

            if (page == pages && r == per) break; /* 맨 마지막은 엑셀 → 생략 */
            if (!ExcelMoveRight(app)) { res = RES_MOVE_FAIL; break; }
        }
    }

    Release(app);
    return res;
}

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
        DWORD vk = kb->vkCode;

        /* 프로그램이 직접 보낸 키는 그대로 통과 */
        if ((kb->flags & LLKHF_INJECTED) || kb->dwExtraInfo == MACRO_SIGNATURE)
            return CallNextHookEx(g_kbHook, nCode, wParam, lParam);

        if (vk == VK_ESCAPE && g_busy) {
            if (down)
                InterlockedExchange(&g_abort, 1);
            return 1;
        }

        if (vk == VK_SUBTRACT) {
            if (!down) {
                if (g_minusDown) { g_minusDown = FALSE; return 1; }
            } else if (g_minusDown) {
                return 1; /* 누르고 있는 동안의 자동반복 */
            } else if (GetAsyncKeyState(VK_CONTROL) & 0x8000) {
                /* Ctrl + 숫자패드 - : 실행 (원래 키 입력은 막음) */
                g_minusDown = TRUE;
                if (InterlockedCompareExchange(&g_busy, 1, 0) == 0) {
                    SetEvent(g_triggerEvent);
                    PostMessageW(g_wnd, WM_APP_UPDATE, 0, 0);
                }
                return 1;
            }
        }

        /* 실행 중에는 Ctrl 을 계속 누르고 있어서 생기는 자동반복 입력을 막는다
         * (막지 않으면 ↓ 가 Ctrl+↓ 로 바뀔 수 있음) */
        if (down && g_busy && (vk == VK_LCONTROL || vk == VK_RCONTROL || vk == VK_CONTROL))
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

#define SET_Y  106             /* 입력칸 줄 위치 */
#define BTN_Y  (SET_Y + 38)

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
        L"PPT 위치 클릭 → 엑셀 시작 셀 클릭 → Ctrl + 숫자패드 -",
        L"최근 결과:  모든 페이지 완료",
        L"최근 결과:  엑셀에 연결하지 못했습니다",
        L"최근 결과:  파워포인트 창을 찾지 못했습니다",
        L"최근 결과:  파워포인트 창으로 전환하지 못했습니다",
        L"최근 결과:  엑셀 복사에 실패했습니다 (셀 편집 중인지 확인)",
        L"최근 결과:  엑셀에서 오른쪽으로 이동하지 못했습니다",
        L"최근 결과:  Esc로 중지했습니다",
        L"최근 결과:  PPT에서 붙여넣을 위치를 먼저 클릭하세요",
        L"최근 결과:  엑셀 시작 셀이 비어 있어서 멈췄습니다 (데이터 끝)",
    };
    PAINTSTRUCT ps;
    HDC dc = BeginPaint(hwnd, &ps);
    WCHAR title[128];
    WCHAR line[256];
    HWND w;
    RECT rc;
    LONG res = g_result;
    int y = 10;

    SetBkMode(dc, TRANSPARENT);

    SelectObject(dc, g_fontBold);
    if (g_busy) {
        SetTextColor(dc, RGB(30, 90, 200));
        wsprintfW(line, L"● 실행 중  페이지 %d/%ld · %d/%ld회  (Esc: 중지)",
                  (int)g_page, g_pages, (int)g_round, g_perPage);
    } else {
        SetTextColor(dc, RGB(30, 140, 60));
        lstrcpyW(line, L"● 대기 중 - Ctrl + 숫자패드 - 누르면 실행");
    }
    DrawLine(dc, &y, line);
    y += 4;

    SelectObject(dc, g_font);
    SetTextColor(dc, RGB(20, 20, 20));
    if (res < 0 || res >= RES_MAX) res = RES_NONE;
    DrawLine(dc, &y, results[res]);

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

    /* 입력칸 이름표 */
    y = SET_Y + 4;
    SetRect(&rc, 12, y, 90, y + 20);
    DrawTextW(dc, L"페이지 수", -1, &rc, DT_LEFT | DT_SINGLELINE);
    SetRect(&rc, 150, y, 260, y + 20);
    DrawTextW(dc, L"페이지당 횟수", -1, &rc, DT_LEFT | DT_SINGLELINE);

    EndPaint(hwnd, &ps);
}

static LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
    switch (msg) {
    case WM_CREATE: {
        WCHAR t[16];
        HWND b = CreateWindowW(L"BUTTON", L"종료", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
                               292, BTN_Y, 120, 30, hwnd, (HMENU)ID_BTN_EXIT, NULL, NULL);
        SendMessageW(b, WM_SETFONT, (WPARAM)g_font, TRUE);

        LoadSettings();
        wsprintfW(t, L"%ld", g_pages);
        g_editPages = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", t,
                                      WS_CHILD | WS_VISIBLE | ES_NUMBER | ES_CENTER,
                                      84, SET_Y, 50, 26, hwnd, (HMENU)ID_EDIT_PAGES, NULL, NULL);
        wsprintfW(t, L"%ld", g_perPage);
        g_editPer = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", t,
                                    WS_CHILD | WS_VISIBLE | ES_NUMBER | ES_CENTER,
                                    250, SET_Y, 50, 26, hwnd, (HMENU)ID_EDIT_PER, NULL, NULL);
        SendMessageW(g_editPages, EM_LIMITTEXT, 2, 0);
        SendMessageW(g_editPer, EM_LIMITTEXT, 2, 0);
        SendMessageW(g_editPages, WM_SETFONT, (WPARAM)g_font, TRUE);
        SendMessageW(g_editPer, WM_SETFONT, (WPARAM)g_font, TRUE);
        SetTimer(hwnd, 1, 1000, NULL); /* 창 이름 표시 갱신 */
        return 0;
    }
    case WM_COMMAND:
        if (LOWORD(wParam) == ID_BTN_EXIT) {
            DestroyWindow(hwnd);
            return 0;
        }
        if ((LOWORD(wParam) == ID_EDIT_PAGES || LOWORD(wParam) == ID_EDIT_PER) &&
            HIWORD(wParam) == EN_CHANGE) {
            LONG pages = ReadNumber(g_editPages), per = ReadNumber(g_editPer);
            if (pages > 0) InterlockedExchange(&g_pages, pages);
            if (per > 0) InterlockedExchange(&g_perPage, per);
            SaveSettings();
            InvalidateRect(hwnd, NULL, TRUE);
            return 0;
        }
        break;
    case WM_APP_UPDATE:
    case WM_TIMER:
        InvalidateRect(hwnd, NULL, TRUE);
        return 0;
    case WM_MOUSEACTIVATE: {
        /* 입력칸을 클릭했을 때만 창을 활성화해서 숫자를 입력할 수 있게 하고,
         * 나머지는 클릭해도 작업 중인 창이 계속 활성 상태로 남도록 함 */
        POINT pt;
        HWND child;
        GetCursorPos(&pt);
        ScreenToClient(hwnd, &pt);
        child = ChildWindowFromPoint(hwnd, pt);
        if (child == g_editPages || child == g_editPer)
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

    /* 중복 실행 방지 */
    mutex = CreateMutexW(NULL, TRUE, L"ExcelPptDown_SingleInstance");
    if (GetLastError() == ERROR_ALREADY_EXISTS) {
        MessageBoxW(NULL, L"ExcelPptDown이 이미 실행 중입니다.", L"ExcelPptDown",
                    MB_OK | MB_ICONINFORMATION);
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
    wc.lpszClassName = L"ExcelPptDownWnd";
    RegisterClassW(&wc);

    AdjustWindowRectEx(&rc, style, FALSE, exStyle);
    g_wnd = CreateWindowExW(exStyle, wc.lpszClassName, L"ExcelPptDown - 엑셀 → PPT 아래로 붙여넣기", style,
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
        MessageBoxW(NULL, L"키보드 후크 설치에 실패했습니다.", L"ExcelPptDown", MB_OK | MB_ICONERROR);
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
