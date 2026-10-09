/*
 * SpaceMacro - 스페이스바 매크로
 *
 * 스페이스바를 누를 때마다 직전(현재 활성) 창에 아래 동작을 번갈아 보낸다.
 *   1회차: 위 방향키 10번 -> 스페이스바
 *   2회차: 아래 방향키 10번 -> 스페이스바
 *   (이후 반복)
 *
 * F8  : 일시정지 / 재개 (일시정지 중에는 스페이스바가 평소처럼 입력됨)
 * 종료: 창의 [종료] 버튼 또는 창 닫기
 *
 * 빌드 (MinGW-w64):
 *   x86_64-w64-mingw32-gcc -O2 -s -static -municode -mwindows \
 *       space_macro.c -o SpaceMacro.exe
 */
#define WIN32_LEAN_AND_MEAN
#ifndef UNICODE
#define UNICODE
#endif
#ifndef _UNICODE
#define _UNICODE
#endif
#include <windows.h>

#define REPEAT_COUNT   10  /* 방향키 반복 횟수 */
#define KEY_HOLD_MS    10  /* 키를 누르고 있는 시간 */
#define KEY_GAP_MS     30  /* 키 입력 사이 간격 */

#define WM_APP_TRIGGER (WM_APP + 1)
#define WM_APP_UPDATE  (WM_APP + 2)

#define ID_BTN_PAUSE   101
#define ID_BTN_EXIT    102

static HWND  g_wnd;
static HHOOK g_kbHook;
static HWINEVENTHOOK g_fgHook;
static HFONT g_font, g_fontBold;

static HWND g_lastTarget;               /* 마지막으로 활성화됐던 다른 프로그램 창 */
static volatile LONG g_busy;            /* 키 전송 중 여부 */
static volatile LONG g_paused;          /* 일시정지 여부 */
static volatile LONG g_nextUp = 1;      /* 다음 동작: 1=위, 0=아래 */
static BOOL g_spaceDown;                /* 물리 스페이스바 눌림 상태 (자동반복 방지) */

static BOOL IsOwnWindow(HWND hwnd)
{
    DWORD pid = 0;
    if (!hwnd) return FALSE;
    GetWindowThreadProcessId(hwnd, &pid);
    return pid == GetCurrentProcessId();
}

static void SendKey(WORD vk, BOOL extended)
{
    INPUT in[1];
    ZeroMemory(in, sizeof(in));
    in[0].type = INPUT_KEYBOARD;
    in[0].ki.wVk = vk;
    in[0].ki.wScan = (WORD)MapVirtualKeyW(vk, MAPVK_VK_TO_VSC);
    in[0].ki.dwFlags = extended ? KEYEVENTF_EXTENDEDKEY : 0;
    SendInput(1, in, sizeof(INPUT));
    Sleep(KEY_HOLD_MS);

    in[0].ki.dwFlags |= KEYEVENTF_KEYUP;
    SendInput(1, in, sizeof(INPUT));
    Sleep(KEY_GAP_MS);
}

/* 실제 키 전송은 별도 스레드에서 수행한다.
 * (키보드 후크가 있는 메인 스레드를 막으면 시스템 전체 입력이 느려진다) */
static DWORD WINAPI MacroThread(LPVOID param)
{
    int i;
    BOOL up = g_nextUp != 0;
    WORD arrow = up ? VK_UP : VK_DOWN;
    (void)param;

    /* 혹시 이 프로그램 창이 활성화되어 있으면 직전 창으로 되돌린다 */
    if (IsOwnWindow(GetForegroundWindow()) && g_lastTarget && IsWindow(g_lastTarget)) {
        SetForegroundWindow(g_lastTarget);
        Sleep(50);
    }

    for (i = 0; i < REPEAT_COUNT; i++)
        SendKey(arrow, TRUE);
    SendKey(VK_SPACE, FALSE);

    InterlockedExchange(&g_nextUp, up ? 0 : 1);
    InterlockedExchange(&g_busy, 0);
    PostMessageW(g_wnd, WM_APP_UPDATE, 0, 0);
    return 0;
}

static BOOL ModifierHeld(void)
{
    return (GetAsyncKeyState(VK_SHIFT)   & 0x8000) ||
           (GetAsyncKeyState(VK_CONTROL) & 0x8000) ||
           (GetAsyncKeyState(VK_MENU)    & 0x8000) ||
           (GetAsyncKeyState(VK_LWIN)    & 0x8000) ||
           (GetAsyncKeyState(VK_RWIN)    & 0x8000);
}

static LRESULT CALLBACK KeyboardProc(int nCode, WPARAM wParam, LPARAM lParam)
{
    if (nCode == HC_ACTION) {
        const KBDLLHOOKSTRUCT *kb = (const KBDLLHOOKSTRUCT *)lParam;
        BOOL down = (wParam == WM_KEYDOWN || wParam == WM_SYSKEYDOWN);

        /* 프로그램이 직접 보낸 키는 그대로 통과 */
        if (kb->flags & LLKHF_INJECTED)
            return CallNextHookEx(g_kbHook, nCode, wParam, lParam);

        if (kb->vkCode == VK_F8) {
            if (down)
                PostMessageW(g_wnd, WM_COMMAND, ID_BTN_PAUSE, 0);
            return 1;
        }

        if (kb->vkCode == VK_SPACE && !g_paused) {
            if (down) {
                if (!g_spaceDown && ModifierHeld()) /* Shift+Space(한/영) 등은 통과 */
                    return CallNextHookEx(g_kbHook, nCode, wParam, lParam);
                if (!g_spaceDown) {
                    g_spaceDown = TRUE;
                    PostMessageW(g_wnd, WM_APP_TRIGGER, 0, 0);
                }
                return 1; /* 원래 스페이스 입력은 막음 */
            }
            if (g_spaceDown) {
                g_spaceDown = FALSE;
                return 1;
            }
        }
    }
    return CallNextHookEx(g_kbHook, nCode, wParam, lParam);
}

static void CALLBACK ForegroundProc(HWINEVENTHOOK hook, DWORD event, HWND hwnd,
                                    LONG idObject, LONG idChild,
                                    DWORD thread, DWORD time)
{
    (void)hook; (void)event; (void)idObject; (void)idChild; (void)thread; (void)time;
    if (hwnd && !IsOwnWindow(hwnd)) {
        g_lastTarget = hwnd;
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
    int y = 10;

    SetBkMode(dc, TRANSPARENT);

    SelectObject(dc, g_fontBold);
    if (g_paused) {
        SetTextColor(dc, RGB(200, 60, 40));
        lstrcpyW(line, L"● 일시정지됨  (F8: 재개)");
    } else {
        SetTextColor(dc, RGB(30, 140, 60));
        lstrcpyW(line, L"● 동작 중  (F8: 일시정지)");
    }
    SetRect(&rc, 12, y, 400, y + 22);
    DrawTextW(dc, line, -1, &rc, DT_LEFT | DT_SINGLELINE | DT_NOPREFIX);
    y += 26;

    SelectObject(dc, g_font);
    SetTextColor(dc, RGB(20, 20, 20));
    wsprintfW(line, L"다음 스페이스바:  %s 방향키 %d번 + 스페이스",
              g_nextUp ? L"↑ 위" : L"↓ 아래", REPEAT_COUNT);
    SetRect(&rc, 12, y, 400, y + 20);
    DrawTextW(dc, line, -1, &rc, DT_LEFT | DT_SINGLELINE | DT_NOPREFIX);
    y += 22;

    if (g_lastTarget && IsWindow(g_lastTarget))
        GetWindowTextW(g_lastTarget, title, 128);
    wsprintfW(line, L"대상 창:  %s", title[0] ? title : L"(활성 창)");
    SetRect(&rc, 12, y, 400, y + 20);
    DrawTextW(dc, line, -1, &rc, DT_LEFT | DT_SINGLELINE | DT_NOPREFIX | DT_END_ELLIPSIS);

    EndPaint(hwnd, &ps);
}

static void UpdatePauseButton(void)
{
    SetDlgItemTextW(g_wnd, ID_BTN_PAUSE, g_paused ? L"재개 (F8)" : L"일시정지 (F8)");
}

static LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
    switch (msg) {
    case WM_CREATE: {
        HWND b1 = CreateWindowW(L"BUTTON", L"일시정지 (F8)", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
                                12, 92, 180, 30, hwnd, (HMENU)ID_BTN_PAUSE, NULL, NULL);
        HWND b2 = CreateWindowW(L"BUTTON", L"종료", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
                                204, 92, 180, 30, hwnd, (HMENU)ID_BTN_EXIT, NULL, NULL);
        SendMessageW(b1, WM_SETFONT, (WPARAM)g_font, TRUE);
        SendMessageW(b2, WM_SETFONT, (WPARAM)g_font, TRUE);
        return 0;
    }
    case WM_COMMAND:
        switch (LOWORD(wParam)) {
        case ID_BTN_PAUSE:
            InterlockedExchange(&g_paused, g_paused ? 0 : 1);
            g_spaceDown = FALSE;
            UpdatePauseButton();
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
        }
        return 0;
    case WM_APP_UPDATE:
        InvalidateRect(hwnd, NULL, TRUE);
        return 0;
    case WM_MOUSEACTIVATE:
        /* 창을 클릭해도 활성화되지 않게 해서 직전 창이 계속 활성 상태로 남도록 함 */
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
    RECT rc = { 0, 0, 396, 134 };
    DWORD style = WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_MINIMIZEBOX;
    HANDLE mutex;
    (void)prev; (void)cmd; (void)show;

    /* 중복 실행 방지 */
    mutex = CreateMutexW(NULL, TRUE, L"SpaceMacro_SingleInstance");
    if (GetLastError() == ERROR_ALREADY_EXISTS) {
        MessageBoxW(NULL, L"SpaceMacro가 이미 실행 중입니다.", L"SpaceMacro", MB_OK | MB_ICONINFORMATION);
        return 0;
    }

    /* 실행 직전에 활성화돼 있던 창 */
    g_lastTarget = GetForegroundWindow();

    g_font = CreateFontW(-15, 0, 0, 0, FW_NORMAL, 0, 0, 0, DEFAULT_CHARSET, 0, 0,
                         CLEARTYPE_QUALITY, 0, L"Malgun Gothic");
    g_fontBold = CreateFontW(-16, 0, 0, 0, FW_BOLD, 0, 0, 0, DEFAULT_CHARSET, 0, 0,
                             CLEARTYPE_QUALITY, 0, L"Malgun Gothic");

    ZeroMemory(&wc, sizeof(wc));
    wc.lpfnWndProc = WndProc;
    wc.hInstance = inst;
    wc.hCursor = LoadCursor(NULL, IDC_ARROW);
    wc.hIcon = LoadIcon(NULL, IDI_APPLICATION);
    wc.hbrBackground = (HBRUSH)(COLOR_BTNFACE + 1);
    wc.lpszClassName = L"SpaceMacroWnd";
    RegisterClassW(&wc);

    AdjustWindowRectEx(&rc, style, FALSE, WS_EX_TOPMOST | WS_EX_NOACTIVATE | WS_EX_APPWINDOW);
    g_wnd = CreateWindowExW(WS_EX_TOPMOST | WS_EX_NOACTIVATE | WS_EX_APPWINDOW,
                            wc.lpszClassName, L"SpaceMacro - 스페이스바 매크로", style,
                            CW_USEDEFAULT, CW_USEDEFAULT,
                            rc.right - rc.left, rc.bottom - rc.top,
                            NULL, NULL, inst, NULL);
    if (!g_wnd) return 1;

    g_kbHook = SetWindowsHookExW(WH_KEYBOARD_LL, KeyboardProc, inst, 0);
    if (!g_kbHook) {
        MessageBoxW(NULL, L"키보드 후크 설치에 실패했습니다.", L"SpaceMacro", MB_OK | MB_ICONERROR);
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
