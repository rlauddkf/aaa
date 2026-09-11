# -*- coding: utf-8 -*-
"""
워드 표 양식 자동 입력기 (Windows)

여러 줄로 입력한 값(예: 숫자 1,2,3...)을 이미 열려 있는 Word 문서의
커서 위치부터 자동으로 타이핑하고, 각 값마다 '아래키'로 다음 칸(표의 아래 셀)로 이동합니다.
설정한 개수(기본 30개)마다 새 페이지에 미리 복사해 둔 빈 표 양식을 붙여넣고
이어서 계속 작성합니다.

필요 라이브러리:
    pip install pyautogui pyperclip

사용법 요약:
    1) Word에서 표 양식 문서를 연다.
    2) 빈 표 양식 하나를 통째로 선택 후 Ctrl+C 로 복사해 둔다. (30개마다 붙여넣기에 사용)
    3) 이 프로그램에 숫자를 줄바꿈으로 붙여넣는다.
    4) '시작'을 누르고, 카운트다운 동안 Word로 전환해 첫 숫자 칸을 클릭한다.
    5) 손을 떼고 기다린다. (중단하려면 마우스를 화면 좌상단 구석으로 휙 옮기거나 'Stop')
"""

import threading
import time
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext

try:
    import pyautogui
except ImportError:  # pragma: no cover
    pyautogui = None

try:
    import pyperclip
except ImportError:  # pragma: no cover
    pyperclip = None


class WordAutoInput:
    def __init__(self, root):
        self.root = root
        self.root.title("워드 표 양식 자동 입력기")
        self.root.geometry("560x680")

        self._worker = None
        self._stop = threading.Event()

        self._build_ui()

    # ------------------------------------------------------------------ UI
    def _build_ui(self):
        pad = {"padx": 10, "pady": 4}

        # 안내
        info = (
            "① Word에서 표 양식을 연다\n"
            "② 빈 표 양식 하나를 선택 → Ctrl+C 로 복사해 둔다\n"
            "③ 아래에 숫자를 줄바꿈으로 붙여넣는다\n"
            "④ '시작' 후 카운트다운 동안 Word로 전환해 첫 칸을 클릭한다"
        )
        ttk.Label(self.root, text=info, justify="left",
                  foreground="#333").pack(anchor="w", **pad)

        # 입력 영역
        ttk.Label(self.root, text="입력값 (한 줄에 하나씩):").pack(anchor="w", **pad)
        self.text = scrolledtext.ScrolledText(self.root, height=12, font=("Consolas", 11))
        self.text.pack(fill="both", expand=True, padx=10, pady=4)
        self.text.insert("1.0", "\n".join(str(i) for i in range(1, 51)))

        # 설정 프레임
        opts = ttk.LabelFrame(self.root, text="설정")
        opts.pack(fill="x", padx=10, pady=8)

        self.items_per_page = tk.IntVar(value=30)
        self.next_key = tk.StringVar(value="down")
        self.up_presses = tk.IntVar(value=30)
        self.auto_up = tk.BooleanVar(value=True)
        self.interval = tk.DoubleVar(value=0.05)
        self.countdown = tk.IntVar(value=5)
        self.use_paste = tk.BooleanVar(value=True)

        def row(parent, r, label, widget):
            ttk.Label(parent, text=label).grid(row=r, column=0, sticky="w", padx=8, pady=4)
            widget.grid(row=r, column=1, sticky="w", padx=8, pady=4)

        row(opts, 0, "페이지당 개수", tk.Spinbox(opts, from_=1, to=999,
            textvariable=self.items_per_page, width=8, command=self._sync_up))
        row(opts, 1, "다음 칸 이동 키", ttk.Combobox(opts, textvariable=self.next_key,
            values=["down", "right", "tab", "enter"], width=8, state="readonly"))

        chk = ttk.Checkbutton(opts, text="위로 이동 횟수 = 페이지당 개수 (자동)",
                              variable=self.auto_up, command=self._sync_up)
        chk.grid(row=2, column=0, columnspan=2, sticky="w", padx=8, pady=2)
        self.up_spin = tk.Spinbox(opts, from_=0, to=999, textvariable=self.up_presses, width=8)
        row(opts, 3, "위로 이동 횟수(복귀)", self.up_spin)

        row(opts, 4, "키 입력 간격(초)", tk.Spinbox(opts, from_=0.0, to=2.0, increment=0.01,
            textvariable=self.interval, width=8, format="%.2f"))
        row(opts, 5, "시작 대기(초)", tk.Spinbox(opts, from_=1, to=30,
            textvariable=self.countdown, width=8))

        ttk.Checkbutton(opts, text="30개마다 새 페이지에 표 양식 붙여넣기(Ctrl+V)",
                        variable=self.use_paste).grid(row=6, column=0, columnspan=2,
                                                      sticky="w", padx=8, pady=2)
        self._sync_up()

        # 버튼
        btns = ttk.Frame(self.root)
        btns.pack(fill="x", padx=10, pady=6)
        self.start_btn = ttk.Button(btns, text="시작", command=self.start)
        self.start_btn.pack(side="left", expand=True, fill="x", padx=4)
        self.stop_btn = ttk.Button(btns, text="Stop", command=self.stop, state="disabled")
        self.stop_btn.pack(side="left", expand=True, fill="x", padx=4)

        # 상태
        self.status = tk.StringVar(value="대기 중")
        ttk.Label(self.root, textvariable=self.status, foreground="#0a6",
                  font=("", 10, "bold")).pack(anchor="w", padx=12, pady=(0, 8))

    def _sync_up(self):
        if self.auto_up.get():
            self.up_presses.set(self.items_per_page.get())
            self.up_spin.config(state="disabled")
        else:
            self.up_spin.config(state="normal")

    # -------------------------------------------------------------- control
    def _set_status(self, text):
        self.root.after(0, lambda: self.status.set(text))

    def start(self):
        if pyautogui is None:
            messagebox.showerror("오류", "pyautogui 가 설치되어 있지 않습니다.\n\n"
                                 "명령창에서:  pip install pyautogui pyperclip")
            return

        items = [line.strip() for line in self.text.get("1.0", "end").splitlines()
                 if line.strip() != ""]
        if not items:
            messagebox.showwarning("입력 없음", "입력값을 한 줄에 하나씩 넣어주세요.")
            return

        if self.use_paste.get() and pyperclip is None:
            if not messagebox.askyesno("확인",
                    "pyperclip 이 없어 클립보드 확인을 건너뜁니다.\n"
                    "미리 빈 표 양식을 Ctrl+C 로 복사해 두셨나요? 계속할까요?"):
                return

        self._stop.clear()
        self.start_btn.config(state="disabled")
        self.stop_btn.config(state="normal")
        self._worker = threading.Thread(target=self._run, args=(items,), daemon=True)
        self._worker.start()

    def stop(self):
        self._stop.set()
        self._set_status("중단 요청됨...")

    def _finish(self, msg):
        self._set_status(msg)
        self.root.after(0, lambda: self.start_btn.config(state="normal"))
        self.root.after(0, lambda: self.stop_btn.config(state="disabled"))

    # ------------------------------------------------------------- automation
    def _run(self, items):
        pyautogui.FAILSAFE = True          # 마우스를 좌상단 구석으로 옮기면 즉시 중단
        pyautogui.PAUSE = 0                 # 우리가 직접 sleep 으로 제어

        per_page = max(1, self.items_per_page.get())
        next_key = self.next_key.get()
        up_n = self.up_presses.get()
        interval = max(0.0, self.interval.get())
        use_paste = self.use_paste.get()

        # 카운트다운
        for s in range(self.countdown.get(), 0, -1):
            if self._stop.is_set():
                self._finish("중단됨")
                return
            self._set_status(f"{s}초 후 시작... Word로 전환해 첫 칸을 클릭하세요")
            time.sleep(1)

        total = len(items)
        try:
            for idx, value in enumerate(items):
                if self._stop.is_set():
                    self._finish(f"중단됨 ({idx}/{total})")
                    return

                self._set_status(f"입력 중 {idx + 1}/{total}: {value}")
                pyautogui.write(value, interval=0)

                is_last_overall = (idx == total - 1)
                is_last_of_page = ((idx + 1) % per_page == 0)

                if is_last_overall:
                    break

                if is_last_of_page:
                    self._new_page(use_paste, up_n)
                else:
                    pyautogui.press(next_key)

                time.sleep(interval)

            self._finish(f"완료! 총 {total}개 입력")
        except pyautogui.FailSafeException:
            self._finish("안전장치(마우스 구석)로 중단됨")
        except Exception as e:  # pragma: no cover
            self._finish(f"오류: {e}")

    def _new_page(self, use_paste, up_n):
        """새 페이지에 빈 표 양식을 붙여넣고 첫 칸으로 커서 복귀."""
        self._set_status("새 페이지에 양식 추가 중...")
        pyautogui.hotkey("ctrl", "end")     # 문서 맨 끝으로
        time.sleep(0.2)
        pyautogui.hotkey("ctrl", "enter")   # 페이지 나누기
        time.sleep(0.2)

        if use_paste:
            pyautogui.hotkey("ctrl", "v")   # 복사해 둔 빈 표 붙여넣기
            time.sleep(0.4)                 # 붙여넣기 렌더 대기
            # 표 아래에서 위로 올라가 첫 번째 숫자 칸으로 복귀
            for _ in range(max(0, up_n)):
                if self._stop.is_set():
                    return
                pyautogui.press("up")
                time.sleep(0.02)


def main():
    root = tk.Tk()
    WordAutoInput(root)
    root.mainloop()


if __name__ == "__main__":
    main()
