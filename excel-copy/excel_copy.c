/*
 * ExcelCopy - 엑셀 → 파워포인트 반복 복사/붙여넣기 매크로
 *
 * Ctrl+Shift 가 함께 눌리는 즉시 아래 동작을 LOOP_COUNT(3)번 반복한다.
 *   1. 엑셀:       Ctrl+Shift+↓ (아래로 데이터 끝까지 선택) → Ctrl+C (복사)
 *   2. 파워포인트: 창 전환 → Ctrl+V (붙여넣기) → → 1번
 *   3. 엑셀:       창 전환 → → 19번 (마지막 회차는 생략, 파워포인트에서 끝남)
 *
 * 각 단계는 고정 대기 대신 "복사 완료(클립보드 변경)", "창 전환 완료",
 * "프로그램이 입력 처리를 마침"을 확인하는 즉시 다음으로 넘어간다.
 *
 * Esc : 실행 중 중지
 * F9  : 일시정지 / 재개
 * 종료: 창의 [종료] 버튼 또는 창 닫기
 *
 * 빌드 (MinGW-w64):
 *   x86_64-w64-mingw32-gcc -O2 -s -static -municode -mwindows \
 *       excel_copy.c -o ExcelCopy.exe
 */
#define WIN32_LEAN_AND_MEAN
#ifndef UNICODE
#define UNICODE
#endif
#ifndef _UNICODE
#define _UNICODE
#endif
#include <windows.h>

#define LOOP_COUNT        3   /* 전체 반복 횟수 */
#define PPT_RIGHT_COUNT   1   /* 붙여넣기 후 파워포인트에서 → 횟수 */
#define EXCEL_RIGHT_COUNT 19  /* 엑셀로 돌아와서 → 횟수 */

#define MACRO_SIGNATURE ((ULONG_PTR)0x58434350) /* 매크로가 보낸 키 표시 */

#define WM_APP_UPDATE  (WM_APP + 2)

#define ID_BTN_PAUSE   101
#define ID_BTN_EXIT    102

/* 실행 결과 */
enum { RES_NONE, RES_DONE, RES_NO_EXCEL, RES_NO_PPT, RES_SWITCH_FAIL, RES_COPY_FAIL, RES_ABORTED };

static HWND  g_wnd;
static HHOOK g_kbHook;
static HANDLE g_triggerEvent;           /* 작업 스레드를 깨우는 신호 */
static HWINEVENTHOOK g_fgHook;
static HFONT g_font, g_fontBold;

static HWND g_lastExcel;                /* 마지막으로 활성화됐던 엑셀 창 */
static HWND g_lastPpt;                  /* 마지막으로 활성화됐던 파워포인트 창 */
static volatile LONG g_busy;            /* 매크로 실행 중 여부 */
static volatile LONG g_abort;           /* Esc로 중지 요청 */
static volatile LONG g_paused;          /* 일시정지 여부 */
static volatile LONG g_result;          /* 최근 실행 결과 (RES_*) */
static volatile LONG g_round;           /* 현재 반복 회차 (1~LOOP_COUNT) */

/* Ctrl+Shift 입력 감지 상태 */
static BOOL g_ctrlDown, g_shiftDown;
static BOOL g_fired;                    /* 이번 Ctrl+Shift 누름에서 이미 실행했음 (자동반복 방지) */

static BOOL HasClass(HWND hwnd, const WCHAR *name)
{
    WCHAR cls[32];
    if (!hwnd || !IsWindow(hwnd)) return FALSE;
    if (!GetClassNameW(hwnd, cls, 32)) return FALSE;
    return lstrcmpW(cls, name) == 0;
}

static BOOL IsExcelWindow(HWND hwnd) { return HasClass(hwnd, L"XLMAIN"); }
static BOOL IsPptWindow(HWND hwnd)   { return HasClass(hwnd, L"PPTFrameClass"); }

static HWND FindExcel(void)
{
    HWND fg = GetForegroundWindow();
    if (IsExcelWindow(fg)) return fg;
    if (IsExcelWindow(g_lastExcel)) return g_lastExcel;
    return FindWindowW(L"XLMAIN", NULL); /* 화면 맨 위의 엑셀 창 */
}

static HWND FindPpt(void)
{
    HWND fg = GetForegroundWindow();
    if (IsPptWindow(fg)) return fg;
    if (IsPptWindow(g_lastPpt)) return g_lastPpt;
    return FindWindowW(L"PPTFrameClass", NULL); /* 화면 맨 위의 파워포인트 창 */
}

/* 대상 프로그램이 지금까지 받은 입력을 모두 처리할 때까지 기다림 */
static void WaitIdle(HWND hwnd)
{
    DWORD_PTR r;
    Sleep(15); /* 보낸 키가 대상 프로그램의 입력 큐에 들어갈 시간 */
    SendMessageTimeoutW(hwnd, WM_NULL, 0, 0, SMTO_NORMAL, 5000, &r);
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

    /* 창 전환이 끝나는 즉시 진행 */
    for (i = 0; i < 20 && GetForegroundWindow() != hwnd; i++)
        Sleep(5);
    if (GetForegroundWindow() != hwnd) {
        SwitchToThisWindow(hwnd, TRUE); /* 일반 방법이 막혔을 때 대체 */
        for (i = 0; i < 100 && GetForegroundWindow() != hwnd; i++)
            Sleep(5);
    }
    if (GetForegroundWindow() != hwnd) return FALSE;
    WaitIdle(hwnd); /* 포커스 처리가 끝날 때까지 */
    return TRUE;
}

static void AddKey(INPUT *in, int *n, WORD vk, BOOL up)
{
    INPUT *k = &in[(*n)++];
    ZeroMemory(k, sizeof(*k));
    k->type = INPUT_KEYBOARD;
    k->ki.wVk = vk;
    k->ki.wScan = (WORD)MapVirtualKeyW(vk, MAPVK_VK_TO_VSC);
    if (vk == VK_DOWN || vk == VK_RIGHT) k->ki.dwFlags |= KEYEVENTF_EXTENDEDKEY;
    if (up) k->ki.dwFlags |= KEYEVENTF_KEYUP;
    k->ki.dwExtraInfo = MACRO_SIGNATURE;
}

static void AddTap(INPUT *in, int *n, WORD vk)
{
    AddKey(in, n, vk, FALSE);
    AddKey(in, n, vk, TRUE);
}

/* 엑셀: Ctrl+Shift+↓ → Ctrl+C. 클립보드가 바뀔 때까지 기다린다. */
static BOOL ExcelSelectCopy(HWND excel)
{
    INPUT in[8];
    int n = 0, i;
    DWORD seq = GetClipboardSequenceNumber();

    /* 사용자가 Ctrl+Shift를 누르고 있을 수 있으므로 복사 전에 Shift를 떼고,
     * 끝나면 Ctrl도 떼어 키가 눌린 채로 남지 않게 한다. */
    AddKey(in, &n, VK_CONTROL, FALSE);  /* Ctrl+Shift+↓ */
    AddKey(in, &n, VK_SHIFT, FALSE);
    AddTap(in, &n, VK_DOWN);
    AddKey(in, &n, VK_SHIFT, TRUE);
    AddTap(in, &n, 'C');                /* Ctrl+C */
    AddKey(in, &n, VK_CONTROL, TRUE);
    SendInput(n, in, sizeof(INPUT));

    for (i = 0; i < 600 && GetClipboardSequenceNumber() == seq; i++) /* 최대 3초 */
        Sleep(5);
    WaitIdle(excel);
    return GetClipboardSequenceNumber() != seq;
}

/* 파워포인트: Ctrl+V → → PPT_RIGHT_COUNT번 */
static void PptPaste(HWND ppt)
{
    INPUT in[4 + PPT_RIGHT_COUNT * 2];
    int n = 0, i;

    AddKey(in, &n, VK_CONTROL, FALSE);
    AddTap(in, &n, 'V');
    AddKey(in, &n, VK_CONTROL, TRUE);
    SendInput(n, in, sizeof(INPUT));
    /* 붙여넣기가 끝날 때까지 (붙여넣기 중에도 메시지를 처리할 수 있어 한 번 더 확인) */
    WaitIdle(ppt);
    Sleep(50);
    WaitIdle(ppt);

    n = 0;
    for (i = 0; i < PPT_RIGHT_COUNT; i++)
        AddTap(in, &n, VK_RIGHT);
    SendInput(n, in, sizeof(INPUT));
    WaitIdle(ppt);
}

/* 엑셀: → EXCEL_RIGHT_COUNT번 */
static void ExcelMoveRight(HWND excel)
{
    INPUT in[EXCEL_RIGHT_COUNT * 2];
    int n = 0, i;

    for (i = 0; i < EXCEL_RIGHT_COUNT; i++)
        AddTap(in, &n, VK_RIGHT);
    SendInput(n, in, sizeof(INPUT));
    WaitIdle(excel);
}

static void SetRound(int r)
{
    InterlockedExchange(&g_round, r);
    PostMessageW(g_wnd, WM_APP_UPDATE, 0, 0);
}

static LONG RunMacro(void)
{
    HWND excel = FindExcel();
    HWND ppt = FindPpt();
    int r;

    if (!excel) return RES_NO_EXCEL;
    if (!ppt) return RES_NO_PPT;

    for (r = 1; r <= LOOP_COUNT; r++) {
        SetRound(r);

        if (g_abort) return RES_ABORTED;
        if (!ActivateWindow(excel)) return RES_SWITCH_FAIL;
        if (!ExcelSelectCopy(excel)) return RES_COPY_FAIL;

        if (g_abort) return RES_ABORTED;
        if (!ActivateWindow(ppt)) return RES_SWITCH_FAIL;
        PptPaste(ppt);

        if (r == LOOP_COUNT) break; /* 마지막 회차는 엑셀로 돌아가지 않음 */

        if (g_abort) return RES_ABORTED;
        if (!ActivateWindow(excel)) return RES_SWITCH_FAIL;
        ExcelMoveRight(excel);
    }
    return RES_DONE;
}

/* 실제 키 전송은 미리 띄워 둔 작업 스레드에서 수행한다.
 * (키보드 후크가 있는 메인 스레드를 막으면 시스템 전체 입력이 느려진다) */
static DWORD WINAPI MacroThread(LPVOID param)
{
    (void)param;
    for (;;) {
        WaitForSingleObject(g_triggerEvent, INFINITE);
        InterlockedExchange(&g_abort, 0);
        InterlockedExchange(&g_result, RunMacro());
        InterlockedExchange(&g_round, 0);
        InterlockedExchange(&g_busy, 0);
        PostMessageW(g_wnd, WM_APP_UPDATE, 0, 0);
    }
    return 0;
}

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

static void CALLBACK ForegroundProc(HWINEVENTHOOK hook, DWORD event, HWND hwnd,
                                    LONG idObject, LONG idChild,
                                    DWORD thread, DWORD time)
{
    (void)hook; (void)event; (void)idObject; (void)idChild; (void)thread; (void)time;
    if (IsExcelWindow(hwnd))
        g_lastExcel = hwnd;
    else if (IsPptWindow(hwnd))
        g_lastPpt = hwnd;
    else
        return;
    InvalidateRect(g_wnd, NULL, TRUE);
}

static void DrawLine(HDC dc, int *y, const WCHAR *text)
{
    RECT rc;
    SetRect(&rc, 12, *y, 420, *y + 20);
    DrawTextW(dc, text, -1, &rc, DT_LEFT | DT_SINGLELINE | DT_NOPREFIX | DT_END_ELLIPSIS);
    *y += 22;
}

static void Paint(HWND hwnd)
{
    static const WCHAR *results[] = {
        L"동작:  복사 → PPT 붙여넣기 → 엑셀 →19번 ×3회 (마지막 →19 생략)",
        L"최근 결과:  3회 반복 완료",
        L"최근 결과:  엑셀 창을 찾지 못했습니다",
        L"최근 결과:  파워포인트 창을 찾지 못했습니다",
        L"최근 결과:  창 전환에 실패했습니다",
        L"최근 결과:  엑셀 복사에 실패했습니다",
        L"최근 결과:  Esc로 중지했습니다",
    };
    PAINTSTRUCT ps;
    HDC dc = BeginPaint(hwnd, &ps);
    WCHAR title[128];
    WCHAR line[256];
    HWND w;
    LONG res = g_result;
    int y = 10;

    SetBkMode(dc, TRANSPARENT);

    SelectObject(dc, g_fontBold);
    if (g_paused) {
        SetTextColor(dc, RGB(200, 60, 40));
        lstrcpyW(line, L"● 일시정지됨  (F9: 재개)");
    } else if (g_busy) {
        SetTextColor(dc, RGB(30, 90, 200));
        wsprintfW(line, L"● 실행 중 %d/%d  (Esc: 중지)", (int)g_round, LOOP_COUNT);
    } else {
        SetTextColor(dc, RGB(30, 140, 60));
        lstrcpyW(line, L"● 대기 중 - Ctrl+Shift 누르면 실행  (F9: 일시정지)");
    }
    DrawLine(dc, &y, line);
    y += 4;

    SelectObject(dc, g_font);
    SetTextColor(dc, RGB(20, 20, 20));
    if (res < 0 || res > RES_ABORTED) res = RES_NONE;
    DrawLine(dc, &y, results[res]);

    title[0] = 0;
    w = FindExcel();
    if (w) GetWindowTextW(w, title, 128);
    wsprintfW(line, L"엑셀:  %s", w ? (title[0] ? title : L"Excel") : L"(없음)");
    DrawLine(dc, &y, line);

    title[0] = 0;
    w = FindPpt();
    if (w) GetWindowTextW(w, title, 128);
    wsprintfW(line, L"파워포인트:  %s", w ? (title[0] ? title : L"PowerPoint") : L"(없음)");
    DrawLine(dc, &y, line);

    EndPaint(hwnd, &ps);
}

static LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
    switch (msg) {
    case WM_CREATE: {
        HWND b1 = CreateWindowW(L"BUTTON", L"일시정지 (F9)", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
                                12, 114, 194, 30, hwnd, (HMENU)ID_BTN_PAUSE, NULL, NULL);
        HWND b2 = CreateWindowW(L"BUTTON", L"종료", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
                                218, 114, 194, 30, hwnd, (HMENU)ID_BTN_EXIT, NULL, NULL);
        SendMessageW(b1, WM_SETFONT, (WPARAM)g_font, TRUE);
        SendMessageW(b2, WM_SETFONT, (WPARAM)g_font, TRUE);
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
        }
        break;
    case WM_APP_UPDATE:
    case WM_TIMER:
        InvalidateRect(hwnd, NULL, TRUE);
        return 0;
    case WM_MOUSEACTIVATE:
        /* 창을 클릭해도 활성화되지 않게 해서 작업 중인 창이 계속 활성 상태로 남도록 함 */
        return MA_NOACTIVATE;
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
    RECT rc = { 0, 0, 424, 156 };
    DWORD style = WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_MINIMIZEBOX;
    DWORD exStyle = WS_EX_TOPMOST | WS_EX_NOACTIVATE | WS_EX_APPWINDOW;
    HANDLE mutex;
    HWND fg;
    (void)prev; (void)cmd; (void)show;

    /* 중복 실행 방지 */
    mutex = CreateMutexW(NULL, TRUE, L"ExcelCopy_SingleInstance");
    if (GetLastError() == ERROR_ALREADY_EXISTS) {
        MessageBoxW(NULL, L"ExcelCopy가 이미 실행 중입니다.", L"ExcelCopy", MB_OK | MB_ICONINFORMATION);
        return 0;
    }

    fg = GetForegroundWindow();
    if (IsExcelWindow(fg)) g_lastExcel = fg;
    if (IsPptWindow(fg))   g_lastPpt = fg;

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
    g_wnd = CreateWindowExW(exStyle, wc.lpszClassName, L"ExcelCopy - 엑셀 → 파워포인트 반복 붙여넣기", style,
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
    UnhookWindowsHookEx(g_kbHook);
    DeleteObject(g_font);
    DeleteObject(g_fontBold);
    if (mutex) CloseHandle(mutex);
    return 0;
}
