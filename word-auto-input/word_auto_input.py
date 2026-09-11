# -*- coding: utf-8 -*-
"""
워드 표 양식 자동 입력기 (Windows, Word COM 연결 방식)

이미 실행 중인 Word 문서에 직접 연결(win32com)하여,
현재 커서가 놓인 표의 셀 위치를 인식하고, 그 열을 따라 아래로 값(숫자 등)을 채웁니다.
설정한 개수(기본 30개)마다 최초 양식(표)을 서식 그대로 정확히 복제하여
새 페이지에 삽입하고, 그 페이지에서 이어서 계속 채웁니다.

키 입력을 흉내내지 않고 Word 개체 모델을 직접 제어하므로
포커스 전환/타이밍 문제 없이 정확하게 동작합니다.

필요 라이브러리 (Windows 전용):
    pip install pywin32

사용법:
    1) Word에서 표 양식 문서를 연다.
    2) 숫자를 넣기 시작할 '첫 번째 칸'에 커서를 둔다(클릭).
    3) 이 프로그램에 숫자를 한 줄에 하나씩 붙여넣는다.
    4) [워드 연결 확인]으로 문서가 잡히는지 확인한 뒤 [시작]을 누른다.
"""

import threading
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext

try:
    import win32com.client as win32
    import pythoncom
    HAS_WIN32 = True
except ImportError:  # pragma: no cover
    HAS_WIN32 = False

# Word 상수
WD_WITHIN_TABLE = 12   # Selection.Information(wdWithInTable)
WD_PAGE_BREAK = 7      # InsertBreak(Type:=wdPageBreak)


class WordAutoInput:
    def __init__(self, root):
        self.root = root
        self.root.title("워드 표 양식 자동 입력기 (Word 연결)")
        self.root.geometry("560x640")

        self._worker = None
        self._stop = threading.Event()

        self._build_ui()

    # ------------------------------------------------------------------ UI
    def _build_ui(self):
        pad = {"padx": 10, "pady": 4}

        info = (
            "① Word에서 표 양식 문서를 연다\n"
            "② 숫자를 넣기 시작할 '첫 칸'에 커서를 둔다(클릭)\n"
            "③ 아래에 숫자를 한 줄에 하나씩 붙여넣는다\n"
            "④ [워드 연결 확인] 후 [시작]"
        )
        ttk.Label(self.root, text=info, justify="left",
                  foreground="#333").pack(anchor="w", **pad)

        ttk.Label(self.root, text="입력값 (한 줄에 하나씩):").pack(anchor="w", **pad)
        self.text = scrolledtext.ScrolledText(self.root, height=12, font=("Consolas", 11))
        self.text.pack(fill="both", expand=True, padx=10, pady=4)
        self.text.insert("1.0", "\n".join(str(i) for i in range(1, 51)))

        opts = ttk.LabelFrame(self.root, text="설정")
        opts.pack(fill="x", padx=10, pady=8)

        self.items_per_page = tk.IntVar(value=30)
        self.copy_scope = tk.StringVar(value="table")

        ttk.Label(opts, text="페이지당 개수").grid(row=0, column=0, sticky="w", padx=8, pady=4)
        tk.Spinbox(opts, from_=1, to=999, textvariable=self.items_per_page,
                   width=8).grid(row=0, column=1, sticky="w", padx=8, pady=4)

        ttk.Label(opts, text="복제할 양식 범위").grid(row=1, column=0, sticky="w", padx=8, pady=4)
        ttk.Combobox(opts, textvariable=self.copy_scope, width=18, state="readonly",
                     values=["table (커서가 있는 표만)", "page (커서가 있는 페이지 전체)"]
                     ).grid(row=1, column=1, sticky="w", padx=8, pady=4)
        # Combobox는 표시 문자열이라 실제 값 매핑
        self.copy_scope.set("table (커서가 있는 표만)")

        btns = ttk.Frame(self.root)
        btns.pack(fill="x", padx=10, pady=6)
        ttk.Button(btns, text="워드 연결 확인",
                   command=self.check_connection).pack(side="left", expand=True, fill="x", padx=4)
        self.start_btn = ttk.Button(btns, text="시작", command=self.start)
        self.start_btn.pack(side="left", expand=True, fill="x", padx=4)
        self.stop_btn = ttk.Button(btns, text="Stop", command=self.stop, state="disabled")
        self.stop_btn.pack(side="left", expand=True, fill="x", padx=4)

        self.status = tk.StringVar(value="대기 중")
        ttk.Label(self.root, textvariable=self.status, foreground="#0a6",
                  font=("", 10, "bold")).pack(anchor="w", padx=12, pady=(0, 8))

    # -------------------------------------------------------------- helpers
    def _set_status(self, text):
        self.root.after(0, lambda: self.status.set(text))

    def _scope(self):
        return "page" if self.copy_scope.get().startswith("page") else "table"

    def _get_values(self):
        return [ln.strip() for ln in self.text.get("1.0", "end").splitlines() if ln.strip() != ""]

    @staticmethod
    def _connect_word():
        """실행 중인 Word에 연결. 없으면 예외."""
        try:
            return win32.GetActiveObject("Word.Application")
        except Exception:
            # 활성 개체가 없을 때 Dispatch 로 재시도(이미 떠 있으면 그 인스턴스에 붙음)
            return win32.Dispatch("Word.Application")

    # -------------------------------------------------------------- actions
    def check_connection(self):
        if not HAS_WIN32:
            messagebox.showerror("오류", "pywin32 가 필요합니다.\n\n명령창에서:  pip install pywin32")
            return

        def work():
            pythoncom.CoInitialize()
            try:
                app = self._connect_word()
                if app.Documents.Count == 0:
                    self._set_status("연결됨: 열린 문서가 없습니다. 문서를 여세요.")
                    return
                doc = app.ActiveDocument
                sel = app.Selection
                in_table = bool(sel.Information(WD_WITHIN_TABLE))
                where = ""
                if in_table:
                    c = sel.Cells(1)
                    where = f" / 커서 위치: {c.RowIndex}행 {c.ColumnIndex}열"
                else:
                    where = " / (커서가 표 안에 없습니다 — 표 셀을 클릭하세요)"
                self._set_status(f"연결됨: '{doc.Name}'{where}")
            except Exception as e:
                self._set_status(f"연결 실패: {e}")
            finally:
                pythoncom.CoUninitialize()

        threading.Thread(target=work, daemon=True).start()

    def start(self):
        if not HAS_WIN32:
            messagebox.showerror("오류", "pywin32 가 필요합니다.\n\n명령창에서:  pip install pywin32")
            return
        values = self._get_values()
        if not values:
            messagebox.showwarning("입력 없음", "입력값을 한 줄에 하나씩 넣어주세요.")
            return

        self._stop.clear()
        self.start_btn.config(state="disabled")
        self.stop_btn.config(state="normal")
        self._worker = threading.Thread(target=self._run, args=(values,), daemon=True)
        self._worker.start()

    def stop(self):
        self._stop.set()
        self._set_status("중단 요청됨...")

    def _finish(self, msg):
        self._set_status(msg)
        self.root.after(0, lambda: self.start_btn.config(state="normal"))
        self.root.after(0, lambda: self.stop_btn.config(state="disabled"))

    # ------------------------------------------------------------- core (COM)
    def _run(self, values):
        pythoncom.CoInitialize()
        app = None
        prev_screen = None
        try:
            per_page = max(1, self.items_per_page.get())
            total = len(values)

            self._set_status("Word에 연결 중...")
            app = self._connect_word()
            if app.Documents.Count == 0:
                self._finish("실패: 열린 Word 문서가 없습니다.")
                return

            doc = app.ActiveDocument
            sel = app.Selection

            if not bool(sel.Information(WD_WITHIN_TABLE)):
                self._finish("실패: 커서를 표 안의 첫 칸에 두고 다시 시작하세요.")
                return

            # 커서가 놓인 셀/표와 시작 위치(행,열) 인식
            start_cell = sel.Cells(1)
            start_row = start_cell.RowIndex
            col = start_cell.ColumnIndex
            form_table = sel.Tables(1)
            scope = self._scope()

            # 복제할 '양식' 범위(FormattedText 스냅샷)를 '비어 있는 지금' 확보
            if scope == "page":
                form_range = self._current_page_range(app, doc, form_table)
            else:
                form_range = form_table.Range

            pages_needed = (total + per_page - 1) // per_page

            prev_screen = app.ScreenUpdating
            app.ScreenUpdating = False

            # 1) 필요한 페이지 수만큼 '빈 양식'을 문서 끝에 미리 복제 (서식 그대로)
            page_tables = [form_table]
            for _ in range(pages_needed - 1):
                if self._stop.is_set():
                    break
                new_table = self._append_form_copy(doc, form_range)
                page_tables.append(new_table)

            # 2) 각 페이지 표의 같은 열을 따라 아래로 값 채우기
            idx = 0
            for p, table in enumerate(page_tables):
                if self._stop.is_set():
                    break
                rows = table.Rows.Count
                for k in range(per_page):
                    if self._stop.is_set() or idx >= total:
                        break
                    row = start_row + k
                    if row > rows:
                        break  # 이 표에 더 채울 행이 없음
                    try:
                        cell = table.Cell(row, col)
                        cell.Range.Text = str(values[idx])
                    except Exception:
                        pass  # 병합/불규칙 셀은 건너뜀
                    idx += 1
                    self._set_status(f"입력 중 {idx}/{total} (페이지 {p + 1}/{len(page_tables)})")

            app.ScreenUpdating = True if prev_screen is None else prev_screen
            try:
                app.ScreenRefresh()
            except Exception:
                pass

            if self._stop.is_set():
                self._finish(f"중단됨 ({idx}/{total} 입력)")
            else:
                self._finish(f"완료! 총 {idx}개 입력, {len(page_tables)}페이지")
        except Exception as e:
            if app is not None:
                try:
                    app.ScreenUpdating = True if prev_screen is None else prev_screen
                except Exception:
                    pass
            self._finish(f"오류: {e}")
        finally:
            pythoncom.CoUninitialize()

    def _append_form_copy(self, doc, form_range):
        """문서 끝에 페이지 나누기 + 양식(FormattedText)을 그대로 삽입하고 새 표를 반환."""
        end = doc.Content.End
        rng = doc.Range(end - 1, end - 1)
        rng.InsertBreak(Type=WD_PAGE_BREAK)     # 새 페이지
        end2 = doc.Content.End
        rng2 = doc.Range(end2 - 1, end2 - 1)
        rng2.FormattedText = form_range.FormattedText   # 서식 그대로 복제
        return doc.Tables(doc.Tables.Count)             # 방금 추가된 마지막 표

    @staticmethod
    def _current_page_range(app, doc, form_table):
        """커서가 있는 페이지 전체 범위를 대략적으로 반환(양식이 표+주변요소일 때)."""
        try:
            # 현재 페이지 번호
            page = form_table.Range.Information(3)  # wdActiveEndPageNumber
            page_start = doc.Range(form_table.Range.Start, form_table.Range.Start)
            # 페이지 시작으로 이동
            start = doc.GoTo(What=1, Which=1, Count=page)  # wdGoToPage=1, wdGoToAbsolute=1
            # 다음 페이지 시작(있으면) 전까지
            try:
                nxt = doc.GoTo(What=1, Which=1, Count=page + 1)
                if nxt.Start > start.Start:
                    return doc.Range(start.Start, nxt.Start)
            except Exception:
                pass
            return doc.Range(start.Start, doc.Content.End)
        except Exception:
            return form_table.Range


def main():
    root = tk.Tk()
    WordAutoInput(root)
    root.mainloop()


if __name__ == "__main__":
    main()
