#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SPV 분석기 GUI  (Windows EXE 로 패키징 가능)
==========================================

파이썬 표준 라이브러리(tkinter)만으로 동작하는 그래픽 버전입니다.
`.SPV` 파일을 열어 요약 정보와 채널 파형을 화면에서 바로 보고,
CSV / HTML 리포트로 내보낼 수 있습니다.

명령줄 없이 쓸 수 있고, PyInstaller 로 단일 `SPV분석기.exe` 로 만들 수 있습니다.
(빌드 방법은 build_exe.bat / README.md 참고)

핵심 파싱 로직은 spv_analyzer.py 를 그대로 재사용합니다.
"""

import os
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# 같은 폴더의 핵심 분석 모듈 재사용
try:
    from spv_analyzer import (SpvFile, channel_stats, export_csv,
                              export_html, NUM_CHANNELS, MAGIC)
except ImportError:
    messagebox.showerror("오류", "spv_analyzer.py 를 찾을 수 없습니다.\n"
                         "같은 폴더에 두거나 함께 패키징하세요.")
    raise

CH_COLORS = ["#2563eb", "#16a34a", "#db2777", "#9333ea", "#ea580c", "#0891b2"]


class SpvGui(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("SPV 파일 분석기")
        self.geometry("880x640")
        self.minsize(720, 520)
        self.spv = None

        # ---- 상단 툴바 --------------------------------------------------
        bar = ttk.Frame(self, padding=8)
        bar.pack(fill="x")
        ttk.Button(bar, text="📂  SPV 파일 열기", command=self.open_file).pack(side="left")
        self.btn_csv = ttk.Button(bar, text="CSV 저장", command=self.save_csv, state="disabled")
        self.btn_csv.pack(side="left", padx=(8, 0))
        self.btn_html = ttk.Button(bar, text="HTML 리포트 저장", command=self.save_html, state="disabled")
        self.btn_html.pack(side="left", padx=(8, 0))
        self.lbl_file = ttk.Label(bar, text="열린 파일 없음", foreground="#64748b")
        self.lbl_file.pack(side="left", padx=16)

        # ---- 본문: 좌 요약 / 우 차트 ------------------------------------
        body = ttk.Panedwindow(self, orient="horizontal")
        body.pack(fill="both", expand=True, padx=8, pady=(0, 8))

        left = ttk.Frame(body)
        self.summary = tk.Text(left, width=42, wrap="word", font=("Consolas", 10),
                               background="#0f172a", foreground="#e2e8f0",
                               relief="flat", padx=10, pady=10)
        self.summary.pack(fill="both", expand=True)
        self.summary.insert("1.0", "  좌측 상단 [SPV 파일 열기] 를 눌러\n  .SPV 파일을 선택하세요.")
        self.summary.config(state="disabled")
        body.add(left, weight=1)

        right = ttk.Frame(body)
        cwrap = ttk.Frame(right)
        cwrap.pack(fill="both", expand=True)
        self.canvas = tk.Canvas(cwrap, background="#f8fafc", highlightthickness=0)
        vsb = ttk.Scrollbar(cwrap, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)
        self.canvas.bind("<Configure>", lambda e: self.draw_charts())
        body.add(right, weight=2)

        # 마우스 휠 스크롤
        self.canvas.bind_all("<MouseWheel>",
                             lambda e: self.canvas.yview_scroll(int(-e.delta / 120), "units"))

    # ------------------------------------------------------------------
    def open_file(self):
        path = filedialog.askopenfilename(
            title="SPV 파일 선택",
            filetypes=[("SPV 파일", "*.SPV *.spv"), ("모든 파일", "*.*")])
        if not path:
            return
        try:
            self.spv = SpvFile(path)
            if not self.spv.blocks:
                raise ValueError("유효한 블록이 없습니다.")
        except Exception as e:
            messagebox.showerror("분석 실패", f"{e}")
            return
        self.lbl_file.config(text=os.path.basename(path))
        self.btn_csv.config(state="normal")
        self.btn_html.config(state="normal")
        self.show_summary()
        self.draw_charts()

    # ------------------------------------------------------------------
    def show_summary(self):
        s = self.spv
        b0 = s.blocks[0]
        t = s.timeline()
        lines = []
        lines.append("═" * 34)
        lines.append(" SPV 파일 분석 요약")
        lines.append("═" * 34)
        lines.append(f" 크기      : {len(s.data):,} B ({len(s.blocks)}블록)")
        lines.append(f" 매직      : {b0.magic.hex(' ')}"
                     + ("  ✔" if b0.magic == MAGIC else "  ✖"))
        lines.append(f" 측정시작  : {b0.start_time}")
        lines.append(f" 측정종료  : {SpvFile.fmt_hms(t[-1] + s.interval_sec)}")
        lines.append(f" 지속시간  : {SpvFile.fmt_hms(t[-1] - t[0] + s.interval_sec)}")
        lines.append(f" 샘플간격  : {s.interval_sec} 초")
        lines.append(f" 총 샘플   : {s.total_records} 개")
        lines.append(f" 채널 수   : {NUM_CHANNELS} 개 (int16)")
        lines.append(f" 장비설정  : {b0.cfg1} / {b0.cfg2} / {b0.cfg3}")
        lines.append("─" * 34)
        lines.append(" [채널별 통계]")
        for c in range(NUM_CHANNELS):
            st = channel_stats(s.channel(c))
            kind = "기준/상태" if st["constant"] else "측정신호"
            lines.append(f"  CH{c+1} {kind}")
            lines.append(f"     min {st['min']:>6}  max {st['max']:>6}")
            lines.append(f"     평균 {st['mean']:>8.1f}  σ {st['std']:>7.1f}")
        lines.append("─" * 34)
        lines.append(" .SPV 는 계측장비의 이진 시계열")
        lines.append(" 로그입니다. 512B 블록 안에 20초")
        lines.append(" 간격 6채널 값이 기록됩니다.")
        self.summary.config(state="normal")
        self.summary.delete("1.0", "end")
        self.summary.insert("1.0", "\n".join(lines))
        self.summary.config(state="disabled")

    # ------------------------------------------------------------------
    def draw_charts(self):
        self.canvas.delete("all")
        if not self.spv:
            return
        s = self.spv
        t = s.timeline()
        W = max(self.canvas.winfo_width(), 300)
        ch_h = 120
        gap = 14
        pad_l, pad_r, pad_t, pad_b = 46, 14, 22, 18
        y = 8
        for c in range(NUM_CHANNELS):
            vals = s.channel(c)
            st = channel_stats(vals)
            top = y
            plot_w = W - pad_l - pad_r
            plot_h = ch_h - pad_t - pad_b
            # 카드 배경
            self.canvas.create_rectangle(6, top, W - 6, top + ch_h,
                                         fill="#ffffff", outline="#e2e8f0")
            kind = "기준/상태" if st["constant"] else "측정신호"
            self.canvas.create_text(pad_l, top + 12, anchor="w",
                                    font=("Segoe UI", 10, "bold"),
                                    text=f"CH{c+1}  ({kind})  min {st['min']} · max {st['max']} · 평균 {st['mean']:.0f}")
            lo, hi = st["min"], st["max"]
            rng = (hi - lo) or 1
            base_y = top + pad_t
            # y 눈금
            for k in range(3):
                val = lo + rng * k / 2
                yy = base_y + plot_h - (val - lo) / rng * plot_h
                self.canvas.create_line(pad_l, yy, W - pad_r, yy, fill="#eef2f7")
                self.canvas.create_text(pad_l - 4, yy, anchor="e",
                                        font=("Segoe UI", 7), fill="#94a3b8",
                                        text=f"{val:.0f}")
            # 선
            n = len(vals)
            pts = []
            for i, v in enumerate(vals):
                px = pad_l + (i / (n - 1 if n > 1 else 1)) * plot_w
                py = base_y + plot_h - (v - lo) / rng * plot_h
                pts.extend((px, py))
            if len(pts) >= 4:
                self.canvas.create_line(*pts, fill=CH_COLORS[c], width=1.6)
            # x 라벨
            for frac in (0.0, 1.0):
                i = int((n - 1) * frac)
                px = pad_l + frac * plot_w
                anc = "w" if frac == 0 else "e"
                self.canvas.create_text(px, top + ch_h - 8, anchor=anc,
                                        font=("Segoe UI", 7), fill="#94a3b8",
                                        text=SpvFile.fmt_hms(t[i]))
            y += ch_h + gap
        self.canvas.configure(scrollregion=(0, 0, W, y))

    # ------------------------------------------------------------------
    def save_csv(self):
        if not self.spv:
            return
        stem = os.path.splitext(os.path.basename(self.spv.path))[0]
        path = filedialog.asksaveasfilename(defaultextension=".csv",
                                            initialfile=stem + ".csv",
                                            filetypes=[("CSV", "*.csv")])
        if path:
            export_csv(self.spv, path)
            messagebox.showinfo("완료", f"CSV 저장됨:\n{path}")

    def save_html(self):
        if not self.spv:
            return
        stem = os.path.splitext(os.path.basename(self.spv.path))[0]
        path = filedialog.asksaveasfilename(defaultextension=".html",
                                            initialfile=stem + ".html",
                                            filetypes=[("HTML", "*.html")])
        if path:
            export_html(self.spv, path)
            messagebox.showinfo("완료", f"HTML 리포트 저장됨:\n{path}")


def main():
    app = SpvGui()
    # 명령줄 인자로 파일을 주면 바로 로드
    if len(sys.argv) > 1 and os.path.isfile(sys.argv[1]):
        try:
            app.spv = SpvFile(sys.argv[1])
            app.lbl_file.config(text=os.path.basename(sys.argv[1]))
            app.btn_csv.config(state="normal")
            app.btn_html.config(state="normal")
            app.show_summary()
            app.after(120, app.draw_charts)
        except Exception:
            pass
    app.mainloop()


if __name__ == "__main__":
    main()
