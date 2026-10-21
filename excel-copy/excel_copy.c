/*
 * ExcelCopy - 엑셀 아래로 선택 후 복사 매크로
 *
 * Ctrl+Shift 를 눌렀다 떼면 (다른 키 없이) 엑셀 창에 아래 동작을 보낸다.
 *   Ctrl+Shift+↓  (현재 셀부터 아래로 데이터 끝까지 선택)
 *   Ctrl+C        (선택 영역 복사)
 *
 * 엑셀이 활성 창이 아니면 마지막으로 사용한 엑셀 창을 앞으로 가져온 뒤 동작한다.
 *
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

#define MACRO_SIGNATURE ((ULONG_PTR)0x58434350) /* 매크로가 보낸 키 표시 */

#define WM_APP_TRIGGER (WM_APP + 1)
#define WM_APP_UPDATE  (WM_APP + 2)

#define ID_BTN_PAUSE   101
#define ID_BTN_EXIT    102

static HWND  g_wnd;
static HHOOK g_kbHook, g_mouseHook;
static HWINEVENTHOOK g_fgHook;
static HFONT g_font, g_fontBold;

static HWND g_lastExcel;                /* 마지막으로 활성화됐던 엑셀 창 */
static volatile LONG g_busy;            /* 키 전송 중 여부 */
static volatile LONG g_paused;          /* 일시정지 여부 */
static volatile LONG g_result;          /* 0=대기, 1=복사 완료, 2=엑셀 없음 */

/* Ctrl+Shift 단독 입력 감지 상태 */
static BOOL g_ctrlDown, g_shiftDown;
static BOOL g_armed;                    /* Ctrl+Shift가 함께 눌렸고 다른 키가 없었음 */

static BOOL IsExcelWindow(HWND hwnd)
{
    WCHAR cls[32];
    if (!hwnd || !IsWindow(hwnd)) return FALSE;
    if (!GetClassNameW(hwnd, cls, 32)) return FALSE;
    return lstrcmpW(cls, L"XLMAIN") == 0;
}

static HWND FindExcel(void)
{
    HWND fg = GetForegroundWindow();
    if (IsExcelWindow(fg)) return fg;
    if (IsExcelWindow(g_lastExcel)) return g_lastExcel;
    return FindWindowW(L"XLMAIN", NULL); /* 화면 맨 위의 엑셀 창 */
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

    /* 창 전환이 끝나는 즉시 진행 (최대 300ms 대기) */
    for (i = 0; i < 30 && GetForegroundWindow() != hwnd; i++)
        Sleep(10);
    if (GetForegroundWindow() != hwnd) return FALSE;
    Sleep(30); /* 엑셀이 포커스를 받을 시간 */
    return TRUE;
}

static void AddKey(INPUT *in, int *n, WORD vk, BOOL up)
{
    INPUT *k = &in[(*n)++];
    ZeroMemory(k, sizeof(*k));
    k->type = INPUT_KEYBOARD;
    k->ki.wVk = vk;
    k->ki.wScan = (WORD)MapVirtualKeyW(vk, MAPVK_VK_TO_VSC);
    if (vk == VK_DOWN) k->ki.dwFlags |= KEYEVENTF_EXTENDEDKEY;
    if (up) k->ki.dwFlags |= KEYEVENTF_KEYUP;
    k->ki.dwExtraInfo = MACRO_SIGNATURE;
}

/* 실제 키 전송은 별도 스레드에서 수행한다.
 * (키보드 후크가 있는 메인 스레드를 막으면 시스템 전체 입력이 느려진다) */
static DWORD WINAPI MacroThread(LPVOID param)
{
    HWND excel = FindExcel();
    (void)param;

    if (!excel || !ActivateWindow(excel)) {
        InterlockedExchange(&g_result, 2);
    } else {
        /* 모든 키를 한 번에 전송 (중간 대기 없음) */
        INPUT in[8];
        int n = 0;
        AddKey(in, &n, VK_CONTROL, FALSE);  /* Ctrl+Shift+↓ */
        AddKey(in, &n, VK_SHIFT, FALSE);
        AddKey(in, &n, VK_DOWN, FALSE);
        AddKey(in, &n, VK_DOWN, TRUE);
        AddKey(in, &n, VK_SHIFT, TRUE);
        AddKey(in, &n, 'C', FALSE);         /* Ctrl+C */
        AddKey(in, &n, 'C', TRUE);
        AddKey(in, &n, VK_CONTROL, TRUE);
        SendInput(n, in, sizeof(INPUT));
        InterlockedExchange(&g_result, 1);
    }

    InterlockedExchange(&g_busy, 0);
    PostMessageW(g_wnd, WM_APP_UPDATE, 0, 0);
    return 0;
}

static LRESULT CALLBACK KeyboardProc(int nCode, WPARAM wParam, LPARAM lParam)
{
    if (nCode == HC_ACTION) {
        const KBDLLHOOKSTRUCT *kb = (const KBDLLHOOKSTRUCT *)lParam;
        BOOL down = (wParam == WM_KEYDOWN || wParam == WM_SYSKEYDOWN);
        DWORD vk = kb->vkCode;

        /* 프로그램이 직접 보낸 키는 그대로 통과 */
        if ((kb->flags & LLKHF_INJECTED) || kb->dwExtraInfo == MACRO_SIGNATURE)
            return CallNextHookEx(g_kbHook, nCode, wParam, lParam);

        if (vk == VK_F9) {
            if (down)
                PostMessageW(g_wnd, WM_COMMAND, ID_BTN_PAUSE, 0);
            return 1;
        }

        if (vk == VK_LCONTROL || vk == VK_RCONTROL || vk == VK_CONTROL) {
            g_ctrlDown = down;
        } else if (vk == VK_LSHIFT || vk == VK_RSHIFT || vk == VK_SHIFT) {
            g_shiftDown = down;
        } else if (down) {
            g_armed = FALSE; /* Ctrl+Shift+다른키 조합이면 취소 */
        }

        if (down && g_ctrlDown && g_shiftDown)
            g_armed = TRUE;

        /* Ctrl, Shift 둘 다 떼었을 때 실행 */
        if (!down && !g_ctrlDown && !g_shiftDown && g_armed) {
            g_armed = FALSE;
            if (!g_paused)
                PostMessageW(g_wnd, WM_APP_TRIGGER, 0, 0);
        }
    }
    return CallNextHookEx(g_kbHook, nCode, wParam, lParam);
}

/* Ctrl+Shift+클릭은 매크로로 보지 않음 */
static LRESULT CALLBACK MouseProc(int nCode, WPARAM wParam, LPARAM lParam)
{
    if (nCode == HC_ACTION &&
        (wParam == WM_LBUTTONDOWN || wParam == WM_RBUTTONDOWN ||
         wParam == WM_MBUTTONDOWN || wParam == WM_MOUSEWHEEL))
        g_armed = FALSE;
    return CallNextHookEx(g_mouseHook, nCode, wParam, lParam);
}

static void CALLBACK ForegroundProc(HWINEVENTHOOK hook, DWORD event, HWND hwnd,
                                    LONG idObject, LONG idChild,
                                    DWORD thread, DWORD time)
{
    (void)hook; (void)event; (void)idObject; (void)idChild; (void)thread; (void)time;
    if (IsExcelWindow(hwnd)) {
        g_lastExcel = hwnd;
        InvalidateRect(g_wnd, NULL, TRUE);
    }
}

static void Paint(HWND hwnd)
{
    PAINTSTRUCT ps;
    HDC dc = BeginPaint(hwnd, &ps);
    RECT rc;
    WCHAR title[128] = L"";
    WCHAR line[256];
    HWND excel;
    int y = 10;

    SetBkMode(dc, TRANSPARENT);

    SelectObject(dc, g_fontBold);
    if (g_paused) {
        SetTextColor(dc, RGB(200, 60, 40));
        lstrcpyW(line, L"● 일시정지됨  (F9: 재개)");
    } else {
        SetTextColor(dc, RGB(30, 140, 60));
        lstrcpyW(line, L"● 대기 중 - Ctrl+Shift 눌렀다 떼기  (F9: 일시정지)");
    }
    SetRect(&rc, 12, y, 420, y + 22);
    DrawTextW(dc, line, -1, &rc, DT_LEFT | DT_SINGLELINE | DT_NOPREFIX | DT_END_ELLIPSIS);
    y += 26;

    SelectObject(dc, g_font);
    SetTextColor(dc, RGB(20, 20, 20));
    if (g_busy)
        lstrcpyW(line, L"실행 중...");
    else if (g_result == 1)
        lstrcpyW(line, L"최근 결과:  Ctrl+Shift+↓ 선택 후 복사 완료");
    else if (g_result == 2)
        lstrcpyW(line, L"최근 결과:  엑셀 창을 찾지 못했습니다");
    else
        lstrcpyW(line, L"동작:  Ctrl+Shift+↓ 선택 → Ctrl+C 복사");
    SetRect(&rc, 12, y, 420, y + 20);
    DrawTextW(dc, line, -1, &rc, DT_LEFT | DT_SINGLELINE | DT_NOPREFIX);
    y += 22;

    excel = FindExcel();
    if (excel)
        GetWindowTextW(excel, title, 128);
    wsprintfW(line, L"엑셀 창:  %s", excel ? (title[0] ? title : L"Excel") : L"(없음)");
    SetRect(&rc, 12, y, 420, y + 20);
    DrawTextW(dc, line, -1, &rc, DT_LEFT | DT_SINGLELINE | DT_NOPREFIX | DT_END_ELLIPSIS);

    EndPaint(hwnd, &ps);
}

static LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
    switch (msg) {
    case WM_CREATE: {
        HWND b1 = CreateWindowW(L"BUTTON", L"일시정지 (F9)", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
                                12, 92, 194, 30, hwnd, (HMENU)ID_BTN_PAUSE, NULL, NULL);
        HWND b2 = CreateWindowW(L"BUTTON", L"종료", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
                                218, 92, 194, 30, hwnd, (HMENU)ID_BTN_EXIT, NULL, NULL);
        SendMessageW(b1, WM_SETFONT, (WPARAM)g_font, TRUE);
        SendMessageW(b2, WM_SETFONT, (WPARAM)g_font, TRUE);
        SetTimer(hwnd, 1, 1000, NULL); /* 엑셀 창 표시 갱신 */
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
    case WM_APP_TRIGGER:
        if (InterlockedCompareExchange(&g_busy, 1, 0) == 0) {
            HANDLE th = CreateThread(NULL, 0, MacroThread, NULL, 0, NULL);
            if (th) CloseHandle(th);
            else InterlockedExchange(&g_busy, 0);
            InvalidateRect(hwnd, NULL, TRUE);
        }
        return 0;
    case WM_APP_UPDATE:
    case WM_TIMER:
        InvalidateRect(hwnd, NULL, TRUE);
        return 0;
    case WM_MOUSEACTIVATE:
        /* 창을 클릭해도 활성화되지 않게 해서 엑셀 창이 계속 활성 상태로 남도록 함 */
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
    RECT rc = { 0, 0, 424, 134 };
    DWORD style = WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_MINIMIZEBOX;
    DWORD exStyle = WS_EX_TOPMOST | WS_EX_NOACTIVATE | WS_EX_APPWINDOW;
    HANDLE mutex;
    (void)prev; (void)cmd; (void)show;

    /* 중복 실행 방지 */
    mutex = CreateMutexW(NULL, TRUE, L"ExcelCopy_SingleInstance");
    if (GetLastError() == ERROR_ALREADY_EXISTS) {
        MessageBoxW(NULL, L"ExcelCopy가 이미 실행 중입니다.", L"ExcelCopy", MB_OK | MB_ICONINFORMATION);
        return 0;
    }

    if (IsExcelWindow(GetForegroundWindow()))
        g_lastExcel = GetForegroundWindow();

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
    g_wnd = CreateWindowExW(exStyle, wc.lpszClassName, L"ExcelCopy - 엑셀 아래로 선택 후 복사", style,
                            CW_USEDEFAULT, CW_USEDEFAULT,
                            rc.right - rc.left, rc.bottom - rc.top,
                            NULL, NULL, inst, NULL);
    if (!g_wnd) return 1;

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
