/*
 * PptCopy - 파워포인트 → 파워포인트 복사/붙여넣기 + 서식 매크로
 *
 * Ctrl + 숫자패드 + 를 누르면:
 *   1. 열려 있는 파워포인트 중 이름이 "프레젠테이션1" 인 파일의 창으로 전환해서
 *      현재 선택된 부분을 복사 (Ctrl+C)
 *   2. 이름이 "OPT" 인 파일의 창으로 전환해서 현재 위치에 붙여넣기 (Ctrl+V)
 *   3. 리본 [테이블 디자인] 탭의 버튼을 직접 누름
 *        음영 ▼ → 채우기 없음
 *        테두리 ▼ → 모든 테두리
 *      버튼은 UI 자동화로 이름을 찾아서 누른다 (한국어/영어 오피스 모두).
 *      버튼을 찾지 못하면 같은 서식을 파워포인트에 직접 명령해서 적용한다.
 *
 * 실행할 때마다 exe와 같은 폴더에 PptCopy_log.txt 로 단계별 기록을 남긴다.
 *
 * 파워포인트에 직접 연결(COM)해서 파일을 이름으로 찾고 서식을 적용한다.
 * 복사/붙여넣기는 직접 누르는 것과 똑같이 Ctrl+C / Ctrl+V 로 한다.
 * 파일 이름은 확장자 없이 비교한다 (OPT.pptx → OPT).
 *
 * 종료: 창의 [종료] 버튼 또는 창 닫기
 *
 * 빌드 (MinGW-w64):
 *   x86_64-w64-mingw32-gcc -O2 -s -static -municode -mwindows \
 *       ppt_copy.c -o PptCopy.exe -lole32 -loleaut32 -luuid
 */
#define WIN32_LEAN_AND_MEAN
#define COBJMACROS
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
#include <uiautomation.h>

#define SRC_NAME L"프레젠테이션1"   /* 복사할 파워포인트 이름 */
#define DST_NAME L"OPT"            /* 붙여넣을 파워포인트 이름 */

#define BORDER_WEIGHT 1.0f         /* 테두리 두께 (pt) */
#define BORDER_COLOR  0x000000     /* 테두리 색 (검은색) */

#define MACRO_SIGNATURE ((ULONG_PTR)0x50505443) /* 매크로가 보낸 키 표시 */

#define WM_APP_UPDATE  (WM_APP + 2)
#define ID_BTN_EXIT    102

/* PowerPoint 상수 */
#define PP_SEL_SHAPES 2
#define PP_SEL_TEXT   3
#define MSO_TRUE      (-1)
#define MSO_FALSE     0

/* 실행 결과 */
enum { RES_NONE, RES_DONE, RES_NO_PPT, RES_NO_SRC, RES_NO_DST, RES_NO_SELECTION,
       RES_SWITCH_FAIL, RES_NOTHING_PASTED, RES_MAX };

static HWND  g_wnd;
static HHOOK g_kbHook;
static HANDLE g_triggerEvent;
static HFONT g_font, g_fontBold;

static volatile LONG g_busy;
static volatile LONG g_result;
static volatile LONG g_formatted;       /* 최근 실행에서 직접 서식을 적용한 표/도형 수 */
static volatile LONG g_uiPressed;       /* 리본 버튼을 누른 결과 (1=채우기 없음, 2=모든 테두리) */
static BOOL g_plusDown;                 /* 숫자패드 + 눌림 상태 (자동반복 방지) */

/* ---------------------------------------------------------------- COM 도우미 */

/* 이름으로 속성/메서드 호출. 인자는 뒤에서부터 순서대로 넘긴다 (COM 규칙).
 * 파워포인트가 바쁘면 거절되므로 잠시 재시도한다. */
static HRESULT Call(IDispatch *obj, WORD flags, const WCHAR *name, VARIANT *result,
                    int argc, ...)
{
    static DISPID putId = DISPID_PROPERTYPUT;
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
    if (flags & DISPATCH_PROPERTYPUT) {
        dp.cNamedArgs = 1;
        dp.rgdispidNamedArgs = &putId;
    }

    for (tries = 0; tries < 60; tries++) { /* 최대 약 3초 */
        hr = obj->lpVtbl->Invoke(obj, id, &IID_NULL, LOCALE_USER_DEFAULT, flags, &dp,
                                 result, NULL, NULL);
        if (hr != RPC_E_CALL_REJECTED && hr != RPC_E_SERVERCALL_RETRYLATER) break;
        Sleep(50);
    }
    return hr;
}

static VARIANT VInt(LONG n)   { VARIANT v; VariantInit(&v); v.vt = VT_I4; v.lVal = n; return v; }
static VARIANT VFloat(float f){ VARIANT v; VariantInit(&v); v.vt = VT_R4; v.fltVal = f; return v; }

static void Release(IDispatch *d)
{
    if (d) d->lpVtbl->Release(d);
}

/* obj.name (속성 또는 인자 있는 속성/메서드) → 객체 */
static IDispatch *Get(IDispatch *obj, const WCHAR *name)
{
    VARIANT v;
    if (FAILED(Call(obj, DISPATCH_PROPERTYGET | DISPATCH_METHOD, name, &v, 0))) return NULL;
    if (v.vt != VT_DISPATCH) { VariantClear(&v); return NULL; }
    return v.pdispVal;
}

static IDispatch *GetItem(IDispatch *obj, const WCHAR *name, LONG index)
{
    VARIANT v;
    if (FAILED(Call(obj, DISPATCH_PROPERTYGET | DISPATCH_METHOD, name, &v, 1, VInt(index))))
        return NULL;
    if (v.vt != VT_DISPATCH) { VariantClear(&v); return NULL; }
    return v.pdispVal;
}

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

static void Put(IDispatch *obj, const WCHAR *name, VARIANT value)
{
    Call(obj, DISPATCH_PROPERTYPUT, name, NULL, 1, value);
}

/* obj.child.name = value */
static void PutChild(IDispatch *obj, const WCHAR *child, const WCHAR *name, VARIANT value)
{
    IDispatch *c = Get(obj, child);
    if (c) { Put(c, name, value); Release(c); }
}

/* ---------------------------------------------------------------- 파워포인트 */

static IDispatch *ConnectPpt(void)
{
    CLSID clsid;
    IUnknown *unk = NULL;
    IDispatch *app = NULL;

    if (FAILED(CLSIDFromProgID(L"PowerPoint.Application", &clsid))) return NULL;
    if (FAILED(GetActiveObject(&clsid, NULL, &unk)) || !unk) return NULL;
    unk->lpVtbl->QueryInterface(unk, &IID_IDispatch, (void **)&app);
    unk->lpVtbl->Release(unk);
    return app;
}

/* 파일 이름 비교 (대소문자 무시, 확장자 무시) */
static BOOL NameMatches(const WCHAR *name, const WCHAR *target)
{
    WCHAR base[MAX_PATH];
    WCHAR *dot = NULL, *c;

    if (!lstrcmpiW(name, target)) return TRUE;
    lstrcpynW(base, name, MAX_PATH);
    for (c = base; *c; c++) if (*c == L'.') dot = c;
    if (dot) *dot = 0;
    return !lstrcmpiW(base, target);
}

/* 이름으로 열린 프레젠테이션 찾기 */
static IDispatch *FindPresentation(IDispatch *app, const WCHAR *target)
{
    IDispatch *list = Get(app, L"Presentations");
    IDispatch *found = NULL;
    LONG i, count = GetInt(list, L"Count");

    for (i = 1; i <= count && !found; i++) {
        IDispatch *pres = GetItem(list, L"Item", i);
        VARIANT v;
        if (pres && SUCCEEDED(Call(pres, DISPATCH_PROPERTYGET, L"Name", &v, 0))) {
            if (v.vt == VT_BSTR && v.bstrVal && NameMatches(v.bstrVal, target)) {
                found = pres;
                pres = NULL;
            }
            VariantClear(&v);
        }
        Release(pres);
    }
    Release(list);
    return found;
}

/* 표의 셀 하나에 채우기 없음 + 테두리 */
static void FormatCell(IDispatch *cell, BOOL top, BOOL left, BOOL bottom, BOOL right)
{
    IDispatch *shape = Get(cell, L"Shape");
    IDispatch *borders = Get(cell, L"Borders");
    BOOL sides[4];
    LONG b;

    sides[0] = top; sides[1] = left; sides[2] = bottom; sides[3] = right;
    if (shape) {
        PutChild(shape, L"Fill", L"Visible", VInt(MSO_FALSE));
        Release(shape);
    }
    for (b = 1; b <= 4 && borders; b++) { /* 1=위 2=왼쪽 3=아래 4=오른쪽 */
        IDispatch *line;
        if (!sides[b - 1]) continue;
        line = GetItem(borders, L"Item", b);
        if (!line) continue;
        PutChild(line, L"ForeColor", L"RGB", VInt(BORDER_COLOR));
        Put(line, L"Weight", VFloat(BORDER_WEIGHT));
        Put(line, L"Visible", VInt(MSO_TRUE));
        Release(line);
    }
    Release(borders);
}

/* 표 전체 또는 선택된 셀에 서식 적용 */
static void FormatTable(IDispatch *table, BOOL onlySelected)
{
    IDispatch *rowsObj = Get(table, L"Rows");
    IDispatch *colsObj = Get(table, L"Columns");
    LONG rows = GetInt(rowsObj, L"Count");
    LONG cols = GetInt(colsObj, L"Count");
    LONG r, c;
    int applied = 0;

    Release(rowsObj);
    Release(colsObj);

    for (r = 1; r <= rows; r++) {
        for (c = 1; c <= cols; c++) {
            VARIANT v;
            IDispatch *cell;
            if (FAILED(Call(table, DISPATCH_METHOD, L"Cell", &v, 2, VInt(c), VInt(r))))
                continue;
            if (v.vt != VT_DISPATCH) { VariantClear(&v); continue; }
            cell = v.pdispVal;
            if (onlySelected && !GetInt(cell, L"Selected")) { Release(cell); continue; }
            if (onlySelected)
                FormatCell(cell, TRUE, TRUE, TRUE, TRUE);
            else /* 표 전체: 이웃 셀과 겹치는 테두리는 한 번만 */
                FormatCell(cell, TRUE, TRUE, r == rows, c == cols);
            applied++;
            Release(cell);
        }
    }

    /* 선택된 셀을 알 수 없으면 표 전체에 적용 */
    if (onlySelected && applied == 0)
        FormatTable(table, FALSE);
}

static void FormatShape(IDispatch *shape, BOOL onlySelectedCells)
{
    if (GetInt(shape, L"HasTable")) {
        IDispatch *table = Get(shape, L"Table");
        if (table) { FormatTable(table, onlySelectedCells); Release(table); }
    } else {
        IDispatch *line = Get(shape, L"Line");
        PutChild(shape, L"Fill", L"Visible", VInt(MSO_FALSE));
        if (line) {
            PutChild(line, L"ForeColor", L"RGB", VInt(BORDER_COLOR));
            Put(line, L"Weight", VFloat(BORDER_WEIGHT));
            Put(line, L"Visible", VInt(MSO_TRUE));
            Release(line);
        }
    }
    InterlockedIncrement(&g_formatted);
}

/* 붙여넣은 뒤 선택된 부분(= 붙여넣은 것)에 서식 적용. 적용한 개수를 돌려줌 */
static int FormatSelection(IDispatch *window)
{
    IDispatch *sel = Get(window, L"Selection");
    IDispatch *range = NULL;
    LONG type = GetInt(sel, L"Type");
    LONG i, count;

    if (type == PP_SEL_SHAPES || type == PP_SEL_TEXT)
        range = Get(sel, L"ShapeRange");
    count = GetInt(range, L"Count");
    for (i = 1; i <= count; i++) {
        IDispatch *shape = GetItem(range, L"Item", i);
        if (shape) {
            /* 표 안에 붙여넣은 경우(텍스트 선택)는 선택된 셀에만 적용 */
            FormatShape(shape, type == PP_SEL_TEXT);
            Release(shape);
        }
    }
    Release(range);
    Release(sel);
    return (int)count;
}

/* ---------------------------------------------------------------- 창 / 입력 */

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

/* 창 제목이 파일 이름(확장자 없이)으로 시작하는 파워포인트 창.
 * 창 제목 예: "OPT.pptx - PowerPoint", "OPT - PowerPoint" */
static HWND FindPptWindowByTitle(const WCHAR *fileName)
{
    WCHAR base[MAX_PATH];
    WCHAR *dot = NULL, *c;
    HWND w = NULL;
    int len;

    lstrcpynW(base, fileName, MAX_PATH);
    for (c = base; *c; c++) if (*c == L'.') dot = c;
    if (dot) *dot = 0;
    len = lstrlenW(base);

    while ((w = FindWindowExW(NULL, w, L"PPTFrameClass", NULL)) != NULL) {
        WCHAR title[256];
        int tlen = GetWindowTextW(w, title, 256);
        if (tlen >= len &&
            CompareStringW(LOCALE_USER_DEFAULT, NORM_IGNORECASE, title, len, base, len) == CSTR_EQUAL &&
            (title[len] == 0 || title[len] == L'.' || title[len] == L' '))
            return w;
    }
    return NULL;
}

/* 프레젠테이션의 창을 앞으로 가져옴. 성공하면 창 핸들 */
static HWND ShowPresentation(IDispatch *app, IDispatch *pres, IDispatch **outWin)
{
    IDispatch *wins = Get(pres, L"Windows");
    IDispatch *win = GetItem(wins, L"Item", 1);
    WCHAR name[MAX_PATH] = L"";
    VARIANT v;
    HWND hwnd;

    Release(wins);
    *outWin = win;
    if (!win) return NULL;
    Call(win, DISPATCH_METHOD, L"Activate", NULL, 0);
    if (SUCCEEDED(Call(pres, DISPATCH_PROPERTYGET, L"Name", &v, 0))) {
        if (v.vt == VT_BSTR && v.bstrVal) lstrcpynW(name, v.bstrVal, MAX_PATH);
        VariantClear(&v);
    }
    hwnd = name[0] ? FindPptWindowByTitle(name) : NULL;
    if (!hwnd) hwnd = (HWND)(LONG_PTR)GetInt(app, L"HWND");
    return ActivateWindow(hwnd) ? hwnd : NULL;
}

static void AddKey(INPUT *in, int *n, WORD vk, BOOL up)
{
    INPUT *k = &in[(*n)++];
    ZeroMemory(k, sizeof(*k));
    k->type = INPUT_KEYBOARD;
    k->ki.wVk = vk;
    k->ki.wScan = (WORD)MapVirtualKeyW(vk, MAPVK_VK_TO_VSC);
    if (up) k->ki.dwFlags |= KEYEVENTF_KEYUP;
    k->ki.dwExtraInfo = MACRO_SIGNATURE;
}

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
    lstrcpyW(slash, L"PptCopy_log.txt");
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

/* ---------------------------------------------------------------- 리본 버튼 누르기 (UI 자동화) */

/* 버튼/메뉴 이름 (한국어, 영어). NULL 로 끝남 */
static const WCHAR *const NAMES_TAB[]     = { L"테이블 디자인", L"표 디자인", L"Table Design", NULL };
static const WCHAR *const NAMES_TAB_OLD[] = { L"디자인", L"Design", NULL }; /* 2016 이전: 표 도구 > 디자인 */
static const WCHAR *const NAMES_SHADING[] = { L"음영", L"Shading", NULL };
static const WCHAR *const NAMES_NOFILL[]  = { L"채우기 없음", L"No Fill", NULL };
static const WCHAR *const NAMES_BORDERS[] = { L"테두리", L"Borders", NULL };
static const WCHAR *const NAMES_ALLBORD[] = { L"모든 테두리", L"All Borders", NULL };

static IUIAutomation *g_uia;

static IUIAutomationCondition *IntCond(PROPERTYID prop, LONG value)
{
    IUIAutomationCondition *c = NULL;
    VARIANT v;
    VariantInit(&v);
    v.vt = VT_I4;
    v.lVal = value;
    IUIAutomation_CreatePropertyCondition(g_uia, prop, v, &c);
    return c;
}

#define CLICKABLE (-1) /* 버튼/분할 버튼/메뉴 항목/목록 항목 (슬라이드 안의 글자와 구분) */

static IUIAutomationCondition *ClickableCond(void)
{
    static const LONG types[] = { 50000 /* Button */, 50031 /* SplitButton */,
                                  50011 /* MenuItem */, 50007 /* ListItem */,
                                  50002 /* CheckBox */ };
    IUIAutomationCondition *cond = NULL;
    int i;
    for (i = 0; i < (int)(sizeof(types) / sizeof(types[0])); i++) {
        IUIAutomationCondition *c = IntCond(UIA_ControlTypePropertyId, types[i]), *both = NULL;
        if (!c) continue;
        if (!cond) { cond = c; continue; }
        IUIAutomation_CreateOrCondition(g_uia, cond, c, &both);
        IUIAutomationCondition_Release(cond);
        IUIAutomationCondition_Release(c);
        cond = both;
    }
    return cond;
}

/* 이름 목록 중 하나와 같은 요소 조건 (+ 컨트롤 종류, 0이면 무시, CLICKABLE 이면 버튼류) */
static IUIAutomationCondition *NameCond(const WCHAR *const *names, LONG controlType)
{
    IUIAutomationCondition *cond = NULL;
    int i;

    for (i = 0; names[i]; i++) {
        IUIAutomationCondition *c = NULL, *both = NULL;
        VARIANT v;
        VariantInit(&v);
        v.vt = VT_BSTR;
        v.bstrVal = SysAllocString(names[i]);
        IUIAutomation_CreatePropertyCondition(g_uia, UIA_NamePropertyId, v, &c);
        VariantClear(&v);
        if (!c) continue;
        if (!cond) {
            cond = c;
        } else {
            IUIAutomation_CreateOrCondition(g_uia, cond, c, &both);
            IUIAutomationCondition_Release(cond);
            IUIAutomationCondition_Release(c);
            cond = both;
        }
    }
    if (cond && controlType) {
        IUIAutomationCondition *both = NULL;
        IUIAutomationCondition *t = controlType == CLICKABLE ? ClickableCond()
                                                             : IntCond(UIA_ControlTypePropertyId, controlType);
        if (t) {
            IUIAutomation_CreateAndCondition(g_uia, cond, t, &both);
            IUIAutomationCondition_Release(t);
            IUIAutomationCondition_Release(cond);
            cond = both;
        }
    }
    return cond;
}

static IUIAutomationElement *FindIn(IUIAutomationElement *root, const WCHAR *const *names,
                                    LONG controlType)
{
    IUIAutomationElement *el = NULL;
    IUIAutomationCondition *cond;

    if (!root) return NULL;
    cond = NameCond(names, controlType);
    if (!cond) return NULL;
    IUIAutomationElement_FindFirst(root, TreeScope_Descendants, cond, &el);
    IUIAutomationCondition_Release(cond);
    return el;
}

/* 드롭다운(▼)을 펼칠 수 있는 버튼 우선으로 찾기.
 * 분할 버튼의 왼쪽 부분을 누르면 메뉴 대신 마지막에 쓴 서식이 바로 적용되므로 피한다. */
static IUIAutomationElement *FindDropdown(IUIAutomationElement *root, const WCHAR *const *names)
{
    IUIAutomationElement *el = NULL;
    IUIAutomationCondition *cond = NameCond(names, CLICKABLE), *exp = NULL, *both = NULL;
    VARIANT v;

    if (!root || !cond) return NULL;
    VariantInit(&v);
    v.vt = VT_BOOL;
    v.boolVal = VARIANT_TRUE;
    IUIAutomation_CreatePropertyCondition(g_uia, UIA_IsExpandCollapsePatternAvailablePropertyId, v, &exp);
    if (exp) {
        IUIAutomation_CreateAndCondition(g_uia, cond, exp, &both);
        IUIAutomationCondition_Release(exp);
    }
    if (both) {
        IUIAutomationElement_FindFirst(root, TreeScope_Descendants, both, &el);
        IUIAutomationCondition_Release(both);
    }
    if (!el)
        IUIAutomationElement_FindFirst(root, TreeScope_Descendants, cond, &el);
    IUIAutomationCondition_Release(cond);
    return el;
}

/* 파워포인트 프로세스의 모든 최상위 창(본 창 + 펼친 메뉴 창)에서 찾기 */
static IUIAutomationElement *FindInProcess(DWORD pid, const WCHAR *const *names, LONG controlType)
{
    IUIAutomationElement *root = NULL, *found = NULL;
    IUIAutomationElementArray *wins = NULL;
    IUIAutomationCondition *pidCond = IntCond(UIA_ProcessIdPropertyId, (LONG)pid);
    int i, count = 0;

    IUIAutomation_GetRootElement(g_uia, &root);
    if (root && pidCond)
        IUIAutomationElement_FindAll(root, TreeScope_Children, pidCond, &wins);
    if (wins) IUIAutomationElementArray_get_Length(wins, &count);
    for (i = 0; i < count && !found; i++) {
        IUIAutomationElement *w = NULL;
        IUIAutomationElementArray_GetElement(wins, i, &w);
        found = FindIn(w, names, controlType);
        if (w) IUIAutomationElement_Release(w);
    }
    if (wins) IUIAutomationElementArray_Release(wins);
    if (pidCond) IUIAutomationCondition_Release(pidCond);
    if (root) IUIAutomationElement_Release(root);
    return found;
}

static void ClickCenter(IUIAutomationElement *el, BOOL rightEdge)
{
    RECT r;
    INPUT in[2];
    POINT old;

    if (FAILED(IUIAutomationElement_get_CurrentBoundingRectangle(el, &r)) || r.right <= r.left)
        return;
    GetCursorPos(&old);
    /* 분할 버튼은 오른쪽(▼) 끝을 눌러야 메뉴가 열림 */
    SetCursorPos(rightEdge ? r.right - 4 : (r.left + r.right) / 2, (r.top + r.bottom) / 2);
    ZeroMemory(in, sizeof(in));
    in[0].type = in[1].type = INPUT_MOUSE;
    in[0].mi.dwFlags = MOUSEEVENTF_LEFTDOWN;
    in[1].mi.dwFlags = MOUSEEVENTF_LEFTUP;
    in[0].mi.dwExtraInfo = in[1].mi.dwExtraInfo = MACRO_SIGNATURE;
    SendInput(2, in, sizeof(INPUT));
    Sleep(30);
    SetCursorPos(old.x, old.y);
}

/* 요소 누르기. expand=TRUE 면 드롭다운 펼치기를 먼저 시도 */
static BOOL Press(IUIAutomationElement *el, BOOL expand)
{
    IUnknown *pat = NULL;

    if (expand &&
        SUCCEEDED(IUIAutomationElement_GetCurrentPatternAs(el, UIA_ExpandCollapsePatternId,
                  &IID_IUIAutomationExpandCollapsePattern, (void **)&pat)) && pat) {
        HRESULT hr = IUIAutomationExpandCollapsePattern_Expand((IUIAutomationExpandCollapsePattern *)pat);
        IUnknown_Release(pat);
        if (SUCCEEDED(hr)) return TRUE;
    }
    pat = NULL;
    if (SUCCEEDED(IUIAutomationElement_GetCurrentPatternAs(el, UIA_InvokePatternId,
                  &IID_IUIAutomationInvokePattern, (void **)&pat)) && pat) {
        HRESULT hr = IUIAutomationInvokePattern_Invoke((IUIAutomationInvokePattern *)pat);
        IUnknown_Release(pat);
        if (SUCCEEDED(hr)) return TRUE;
    }
    pat = NULL;
    if (SUCCEEDED(IUIAutomationElement_GetCurrentPatternAs(el, UIA_SelectionItemPatternId,
                  &IID_IUIAutomationSelectionItemPattern, (void **)&pat)) && pat) {
        HRESULT hr = IUIAutomationSelectionItemPattern_Select((IUIAutomationSelectionItemPattern *)pat);
        IUnknown_Release(pat);
        if (SUCCEEDED(hr)) return TRUE;
    }
    pat = NULL;
    if (SUCCEEDED(IUIAutomationElement_GetCurrentPatternAs(el, UIA_LegacyIAccessiblePatternId,
                  &IID_IUIAutomationLegacyIAccessiblePattern, (void **)&pat)) && pat) {
        HRESULT hr = IUIAutomationLegacyIAccessiblePattern_DoDefaultAction(
                         (IUIAutomationLegacyIAccessiblePattern *)pat);
        IUnknown_Release(pat);
        if (SUCCEEDED(hr)) return TRUE;
    }
    ClickCenter(el, expand);
    return TRUE;
}

/* 리본에서 테이블 디자인 탭을 열고 음영/테두리 버튼이 보이게 함 */
static IUIAutomationElement *FindRibbonButton(HWND hwnd, const WCHAR *const *names)
{
    IUIAutomationElement *win = NULL, *btn = NULL, *tab;
    int i;

    IUIAutomation_ElementFromHandle(g_uia, hwnd, &win);
    if (!win) return NULL;

    btn = FindDropdown(win, names);
    if (!btn) {
        tab = FindIn(win, NAMES_TAB, UIA_TabItemControlTypeId);
        if (tab) {
            Log(L"  리본: 테이블 디자인 탭 선택");
            Press(tab, FALSE);
            IUIAutomationElement_Release(tab);
            for (i = 0; i < 20 && !btn; i++) { Sleep(25); btn = FindDropdown(win, names); }
        }
    }
    if (!btn) {
        /* 예전 버전: '디자인' 탭이 두 개(일반, 표 도구)면 뒤쪽이 표 도구 */
        IUIAutomationElementArray *tabs = NULL;
        IUIAutomationCondition *cond = NameCond(NAMES_TAB_OLD, UIA_TabItemControlTypeId);
        int count = 0;
        if (cond) {
            IUIAutomationElement_FindAll(win, TreeScope_Descendants, cond, &tabs);
            IUIAutomationCondition_Release(cond);
        }
        if (tabs) IUIAutomationElementArray_get_Length(tabs, &count);
        if (count > 0) {
            IUIAutomationElement *last = NULL;
            IUIAutomationElementArray_GetElement(tabs, count - 1, &last);
            if (last) {
                Log(L"  리본: 디자인 탭 선택 (표 도구)");
                Press(last, FALSE);
                IUIAutomationElement_Release(last);
                for (i = 0; i < 20 && !btn; i++) { Sleep(25); btn = FindDropdown(win, names); }
            }
        }
        if (tabs) IUIAutomationElementArray_Release(tabs);
    }
    IUIAutomationElement_Release(win);
    return btn;
}

/* 리본의 드롭다운 버튼을 펼치고 메뉴 항목을 누름 (예: 음영 → 채우기 없음) */
static BOOL PressRibbonMenu(HWND hwnd, const WCHAR *const *button, const WCHAR *const *item)
{
    IUIAutomationElement *btn, *menuItem = NULL;
    DWORD pid = 0;
    int i;

    if (!g_uia) return FALSE;
    GetWindowThreadProcessId(hwnd, &pid);

    btn = FindRibbonButton(hwnd, button);
    if (!btn) {
        Log(L"  리본: '%s' 버튼을 찾지 못함", button[0]);
        return FALSE;
    }
    Press(btn, TRUE);
    IUIAutomationElement_Release(btn);

    for (i = 0; i < 40 && !menuItem; i++) { /* 메뉴가 열릴 때까지 최대 약 1초 */
        Sleep(25);
        menuItem = FindInProcess(pid, item, CLICKABLE);
    }
    if (!menuItem) {
        Log(L"  리본: '%s' 메뉴 항목을 찾지 못함", item[0]);
        {
            INPUT in[2];
            int n = 0;
            AddKey(in, &n, VK_ESCAPE, FALSE); /* 열린 메뉴 닫기 */
            AddKey(in, &n, VK_ESCAPE, TRUE);
            SendInput(n, in, sizeof(INPUT));
        }
        return FALSE;
    }
    Press(menuItem, FALSE);
    IUIAutomationElement_Release(menuItem);
    WaitIdle(hwnd);
    Log(L"  리본: %s → %s 누름", button[0], item[0]);
    return TRUE;
}

/* ---------------------------------------------------------------- 매크로 */

static void SendCtrlKey(WORD key)
{
    INPUT in[4];
    int n = 0;
    AddKey(in, &n, VK_CONTROL, FALSE);
    AddKey(in, &n, key, FALSE);
    AddKey(in, &n, key, TRUE);
    AddKey(in, &n, VK_CONTROL, TRUE);
    SendInput(n, in, sizeof(INPUT));
}

static LONG RunMacro(void)
{
    IDispatch *app, *src = NULL, *dst = NULL, *srcWin = NULL, *dstWin = NULL, *sel = NULL;
    HWND srcHwnd, dstHwnd;
    DWORD seq;
    LONG res = RES_DONE;
    int i;

    InterlockedExchange(&g_formatted, 0);

    InterlockedExchange(&g_uiPressed, 0);
    app = ConnectPpt();
    Log(L"시작: 파워포인트 연결 %s", app ? L"성공" : L"실패");
    if (!app) return RES_NO_PPT;

    src = FindPresentation(app, SRC_NAME);
    dst = FindPresentation(app, DST_NAME);
    Log(L"'%s' %s, '%s' %s", SRC_NAME, src ? L"찾음" : L"없음", DST_NAME, dst ? L"찾음" : L"없음");
    if (!src) { res = RES_NO_SRC; goto done; }
    if (!dst) { res = RES_NO_DST; goto done; }

    /* 1. 프레젠테이션1 창에서 선택된 부분 복사 */
    srcHwnd = ShowPresentation(app, src, &srcWin);
    if (!srcHwnd) { res = RES_SWITCH_FAIL; goto done; }
    sel = Get(srcWin, L"Selection");
    if (!sel || GetInt(sel, L"Type") == 0) { res = RES_NO_SELECTION; goto done; }

    seq = GetClipboardSequenceNumber();
    SendCtrlKey('C');
    for (i = 0; i < 600 && GetClipboardSequenceNumber() == seq; i++) /* 최대 3초 */
        Sleep(5);
    if (GetClipboardSequenceNumber() == seq) { res = RES_NO_SELECTION; goto done; }
    WaitIdle(srcHwnd);

    /* 2. OPT 창으로 전환해서 현재 위치에 붙여넣기 */
    dstHwnd = ShowPresentation(app, dst, &dstWin);
    if (!dstHwnd) { res = RES_SWITCH_FAIL; goto done; }
    SendCtrlKey('V');
    WaitIdle(dstHwnd);
    Sleep(50);
    WaitIdle(dstHwnd);

    Log(L"붙여넣기 완료");

    /* 3. 리본 버튼 누르기: 음영 → 채우기 없음, 테두리 → 모든 테두리 */
    {
        BOOL fill = PressRibbonMenu(dstHwnd, NAMES_SHADING, NAMES_NOFILL);
        BOOL border = PressRibbonMenu(dstHwnd, NAMES_BORDERS, NAMES_ALLBORD);
        InterlockedExchange(&g_uiPressed, (fill ? 1 : 0) + (border ? 2 : 0));
        if (!fill || !border) {
            /* 버튼을 못 누른 경우 같은 서식을 직접 적용 */
            Log(L"버튼을 누르지 못해 서식을 직접 적용");
            if (FormatSelection(dstWin) == 0) res = RES_NOTHING_PASTED;
        }
    }

done:
    Release(sel);
    Release(srcWin);
    Release(dstWin);
    Release(src);
    Release(dst);
    Release(app);
    return res;
}

static DWORD WINAPI MacroThread(LPVOID param)
{
    (void)param;
    CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);
    CoCreateInstance(&CLSID_CUIAutomation, NULL, CLSCTX_INPROC_SERVER, &IID_IUIAutomation,
                     (void **)&g_uia);
    for (;;) {
        WaitForSingleObject(g_triggerEvent, INFINITE);
        LogOpen();
        InterlockedExchange(&g_result, RunMacro());
        Log(L"끝: 결과 코드 = %d, 리본 버튼 = %d", (int)g_result, (int)g_uiPressed);
        LogClose();
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

        if (vk == VK_ADD) {
            if (!down) {
                if (g_plusDown) { g_plusDown = FALSE; return 1; }
            } else if (g_plusDown) {
                return 1; /* 누르고 있는 동안의 자동반복 */
            } else if (GetAsyncKeyState(VK_CONTROL) & 0x8000) {
                /* Ctrl + 숫자패드 + : 실행 (원래 키 입력은 막음) */
                g_plusDown = TRUE;
                if (InterlockedCompareExchange(&g_busy, 1, 0) == 0) {
                    SetEvent(g_triggerEvent);
                    PostMessageW(g_wnd, WM_APP_UPDATE, 0, 0);
                }
                return 1;
            }
        }

        /* 실행 중에는 Ctrl 을 계속 누르고 있어서 생기는 자동반복 입력을 막는다 */
        if (down && g_busy && (vk == VK_LCONTROL || vk == VK_RCONTROL || vk == VK_CONTROL))
            return 1;
    }
    return CallNextHookEx(g_kbHook, nCode, wParam, lParam);
}

/* ---------------------------------------------------------------- 화면 */

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
        L"프레젠테이션1에서 선택 → Ctrl + 숫자패드 +",
        L"최근 결과:  붙여넣기 → 채우기 없음 → 모든 테두리 완료",
        L"최근 결과:  파워포인트에 연결하지 못했습니다",
        L"최근 결과:  '" SRC_NAME L"' 파일이 열려 있지 않습니다",
        L"최근 결과:  '" DST_NAME L"' 파일이 열려 있지 않습니다",
        L"최근 결과:  '" SRC_NAME L"'에서 선택된 부분이 없습니다",
        L"최근 결과:  '" DST_NAME L"' 창으로 전환하지 못했습니다",
        L"최근 결과:  붙여넣은 내용을 찾지 못해 서식을 적용하지 못했습니다",
    };
    PAINTSTRUCT ps;
    HDC dc = BeginPaint(hwnd, &ps);
    WCHAR line[256];
    LONG res = g_result;
    int y = 10;

    SetBkMode(dc, TRANSPARENT);

    SelectObject(dc, g_fontBold);
    if (g_busy) {
        SetTextColor(dc, RGB(30, 90, 200));
        lstrcpyW(line, L"● 실행 중...");
    } else {
        SetTextColor(dc, RGB(30, 140, 60));
        lstrcpyW(line, L"● 대기 중 - Ctrl + 숫자패드 + 누르면 실행");
    }
    DrawLine(dc, &y, line);
    y += 4;

    SelectObject(dc, g_font);
    SetTextColor(dc, RGB(20, 20, 20));
    if (res < 0 || res >= RES_MAX) res = RES_NONE;
    DrawLine(dc, &y, results[res]);
    if (res == RES_DONE) {
        if (g_uiPressed == 3)
            lstrcpyW(line, L"리본 버튼:  채우기 없음, 모든 테두리 누름");
        else
            wsprintfW(line, L"리본 버튼을 찾지 못해 직접 적용함 (표/도형 %d개)", (int)g_formatted);
        DrawLine(dc, &y, line);
    } else {
        DrawLine(dc, &y, L"복사:  " SRC_NAME L"  →  붙여넣기:  " DST_NAME);
    }

    EndPaint(hwnd, &ps);
}

static LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
    switch (msg) {
    case WM_CREATE: {
        HWND b = CreateWindowW(L"BUTTON", L"종료", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
                               292, 84, 120, 30, hwnd, (HMENU)ID_BTN_EXIT, NULL, NULL);
        SendMessageW(b, WM_SETFONT, (WPARAM)g_font, TRUE);
        return 0;
    }
    case WM_COMMAND:
        if (LOWORD(wParam) == ID_BTN_EXIT) {
            DestroyWindow(hwnd);
            return 0;
        }
        break;
    case WM_APP_UPDATE:
        InvalidateRect(hwnd, NULL, TRUE);
        return 0;
    case WM_MOUSEACTIVATE:
        /* 창을 클릭해도 활성화되지 않게 해서 작업 중인 창이 계속 활성 상태로 남도록 함 */
        return MA_NOACTIVATE;
    case WM_PAINT:
        Paint(hwnd);
        return 0;
    case WM_DESTROY:
        PostQuitMessage(0);
        return 0;
    }
    return DefWindowProcW(hwnd, msg, wParam, lParam);
}

int WINAPI wWinMain(HINSTANCE inst, HINSTANCE prev, PWSTR cmd, int show)
{
    WNDCLASSW wc;
    MSG msg;
    RECT rc = { 0, 0, 424, 126 };
    DWORD style = WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_MINIMIZEBOX;
    DWORD exStyle = WS_EX_TOPMOST | WS_EX_NOACTIVATE | WS_EX_APPWINDOW;
    HANDLE mutex;
    (void)prev; (void)cmd; (void)show;

    /* 중복 실행 방지 */
    mutex = CreateMutexW(NULL, TRUE, L"PptCopy_SingleInstance");
    if (GetLastError() == ERROR_ALREADY_EXISTS) {
        MessageBoxW(NULL, L"PptCopy가 이미 실행 중입니다.", L"PptCopy", MB_OK | MB_ICONINFORMATION);
        return 0;
    }

    /* 화면 배율(DPI)과 관계없이 버튼 위치 좌표가 정확하도록 */
    {
        typedef BOOL (WINAPI *SetCtxFn)(HANDLE);
        SetCtxFn setCtx = (SetCtxFn)(void *)GetProcAddress(GetModuleHandleW(L"user32.dll"),
                                                           "SetProcessDpiAwarenessContext");
        if (!setCtx || !setCtx((HANDLE)-4)) /* PER_MONITOR_AWARE_V2 */
            SetProcessDPIAware();
    }

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
    wc.lpszClassName = L"PptCopyWnd";
    RegisterClassW(&wc);

    AdjustWindowRectEx(&rc, style, FALSE, exStyle);
    g_wnd = CreateWindowExW(exStyle, wc.lpszClassName, L"PptCopy - 프레젠테이션1 → OPT 붙여넣기", style,
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
        MessageBoxW(NULL, L"키보드 후크 설치에 실패했습니다.", L"PptCopy", MB_OK | MB_ICONERROR);
        return 1;
    }

    /* 창을 띄우되 포커스는 빼앗지 않음 */
    ShowWindow(g_wnd, SW_SHOWNOACTIVATE);
    UpdateWindow(g_wnd);

    while (GetMessageW(&msg, NULL, 0, 0) > 0) {
        TranslateMessage(&msg);
        DispatchMessageW(&msg);
    }

    UnhookWindowsHookEx(g_kbHook);
    DeleteObject(g_font);
    DeleteObject(g_fontBold);
    if (mutex) CloseHandle(mutex);
    return 0;
}
