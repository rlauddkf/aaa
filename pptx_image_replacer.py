# -*- coding: utf-8 -*-
"""
PowerPoint 이미지 일괄 교체 프로그램  (슬라이드 직접 클릭 방식)
================================================================

현재 "열려 있는" PowerPoint 를 그대로 제어합니다.

 1) 프로그램에 현재 슬라이드가 그림으로 그려집니다.
 2) 바꾸고 싶은 사진을 슬라이드에서 "직접 클릭" 하면 1, 2, 3 … 순번이 붙습니다.
    (같은 사진을 다시 누르면 취소, 슬라이드를 넘겨가며 여러 장 선택 가능)
 3) 오른쪽에 순서대로 "새 파일 선택" 칸이 생깁니다. 새 이미지를 고르고
 4) [교체 실행] 을 누르면 각 사진이 "같은 위치 · 같은 크기" 로 바뀝니다.

동작 환경: Windows + PowerPoint(데스크톱) 설치 필요
필요 패키지: pywin32 (필수), Pillow (필수 - 화면 표시용)

    pip install pywin32 Pillow

실행:
    python pptx_image_replacer.py
"""

import os
import sys
import tempfile
import traceback

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

try:
    from PIL import Image, ImageTk
    HAS_PIL = True
except Exception:
    HAS_PIL = False

# ---------------------------------------------------------------------------
# COM 상수 (PowerPoint Object Model)
# ---------------------------------------------------------------------------
MSO_FALSE = 0
MSO_TRUE = -1

MSO_PICTURE = 13          # msoPicture
MSO_LINKED_PICTURE = 11   # msoLinkedPicture
MSO_PLACEHOLDER = 14      # msoPlaceholder

MSO_SEND_BACKWARD = 3     # msoSendBackward

IMAGE_TYPES = [
    ("이미지 파일", "*.png *.jpg *.jpeg *.gif *.bmp *.tif *.tiff *.emf *.wmf"),
    ("모든 파일", "*.*"),
]


def _import_win32():
    import win32com.client  # noqa
    return win32com.client


# ===========================================================================
#  PowerPoint 제어
# ===========================================================================
class PowerPointController:
    def __init__(self):
        self.app = None
        self.pres = None

    def connect(self):
        win32com = _import_win32()
        try:
            self.app = win32com.GetActiveObject("PowerPoint.Application")
        except Exception:
            self.app = win32com.Dispatch("PowerPoint.Application")
        self.app.Visible = True
        if self.app.Presentations.Count == 0:
            raise RuntimeError(
                "열려 있는 프레젠테이션이 없습니다.\n"
                "PowerPoint 에서 파일을 먼저 연 뒤 다시 [PowerPoint 연결] 을 눌러주세요."
            )
        self.pres = self.app.ActivePresentation
        return self.pres.Name

    def slide_count(self):
        return int(self.pres.Slides.Count)

    def current_slide_index(self):
        try:
            return int(self.app.ActiveWindow.View.Slide.SlideIndex)
        except Exception:
            return 1

    def slide_size_pt(self):
        ps = self.pres.PageSetup
        return float(ps.SlideWidth), float(ps.SlideHeight)

    @staticmethod
    def _is_picture(shape):
        try:
            if shape.Type in (MSO_PICTURE, MSO_LINKED_PICTURE):
                return True
            if shape.Type == MSO_PLACEHOLDER:
                try:
                    return shape.Fill.Type == 6
                except Exception:
                    return False
        except Exception:
            return False
        return False

    def pictures_on_slide(self, slide_index):
        """해당 슬라이드의 사진 목록 (위치/크기는 points 단위)."""
        result = []
        slide = self.pres.Slides.Item(slide_index)
        for shape in slide.Shapes:
            if not self._is_picture(shape):
                continue
            try:
                result.append({
                    "slide_index": slide_index,
                    "shape_id": int(shape.Id),
                    "name": str(shape.Name),
                    "left": float(shape.Left),
                    "top": float(shape.Top),
                    "width": float(shape.Width),
                    "height": float(shape.Height),
                })
            except Exception:
                continue
        return result

    def export_slide_png(self, slide_index, out_path, px_w, px_h):
        slide = self.pres.Slides.Item(slide_index)
        slide.Export(out_path, "PNG", int(px_w), int(px_h))

    def find_shape(self, slide_index, shape_id):
        slide = self.pres.Slides.Item(slide_index)
        for shape in slide.Shapes:
            try:
                if int(shape.Id) == shape_id:
                    return slide, shape
            except Exception:
                continue
        return slide, None

    def replace_picture(self, target, image_path, keep_aspect=False, lock_ratio=True):
        slide, old = self.find_shape(target["slide_index"], target["shape_id"])
        if old is None:
            raise RuntimeError("대상 사진을 찾을 수 없습니다 (이미 변경됨).")

        left, top = float(old.Left), float(old.Top)
        width, height = float(old.Width), float(old.Height)
        rotation = float(old.Rotation)
        try:
            old_z = int(old.ZOrderPosition)
        except Exception:
            old_z = None

        new_shape = slide.Shapes.AddPicture(
            FileName=image_path, LinkToFile=MSO_FALSE, SaveWithDocument=MSO_TRUE,
            Left=left, Top=top, Width=width, Height=height,
        )
        if keep_aspect:
            self._fit_keep_aspect(new_shape, left, top, width, height)
        # 교체된 사진에 가로세로 비율 잠금을 자동 적용 (이후 크기 조절 시 찌그러짐 방지)
        if lock_ratio:
            try:
                new_shape.LockAspectRatio = MSO_TRUE
            except Exception:
                pass
        try:
            new_shape.Rotation = rotation
        except Exception:
            pass

        try:
            new_id = int(new_shape.Id)
        except Exception:
            new_id = None
        old.Delete()

        if old_z is not None and new_id is not None:
            self._move_to_zorder(slide, new_id, old_z)

    @staticmethod
    def _fit_keep_aspect(shape, bx, by, bw, bh):
        try:
            shape.LockAspectRatio = MSO_TRUE
            nw, nh = float(shape.Width), float(shape.Height)
            if nw <= 0 or nh <= 0:
                return
            s = min(bw / nw, bh / nh)
            shape.Width, shape.Height = nw * s, nh * s
            shape.Left = bx + (bw - nw * s) / 2.0
            shape.Top = by + (bh - nh * s) / 2.0
        except Exception:
            pass

    @staticmethod
    def _move_to_zorder(slide, shape_id, target_z):
        for _ in range(200):
            shape = None
            for s in slide.Shapes:
                try:
                    if int(s.Id) == shape_id:
                        shape = s
                        break
                except Exception:
                    continue
            if shape is None:
                return
            try:
                cur_z = int(shape.ZOrderPosition)
            except Exception:
                return
            if cur_z <= target_z:
                return
            try:
                shape.ZOrder(MSO_SEND_BACKWARD)
            except Exception:
                return


# ===========================================================================
#  GUI
# ===========================================================================
class App(tk.Tk):
    BADGE_COLORS = "#e53935"
    ZOOM_STEP = 1.25
    ZOOM_MIN = 0.1
    ZOOM_MAX = 4.0

    def __init__(self):
        super().__init__()
        self.title("PowerPoint 사진 일괄 교체 — 슬라이드에서 직접 클릭")
        self.geometry("1160x760")
        self.minsize(1000, 640)

        self.ctrl = PowerPointController()
        self._tmp = tempfile.mkdtemp(prefix="pptimg_")

        self.slide_index = 1
        self.slide_photo = None
        self.scale = 1.0                # points -> display px (= zoom)
        self.zoom = None                # None 이면 '화면 맞춤' 자동 계산
        self.overlays = []              # [(x0,y0,x1,y1,target)] 현재 슬라이드
        self.sequence = []             # 선택된 target dict 리스트 (교체 순서)
        self._row_thumbs = []           # PhotoImage 참조 유지

        self._build_ui()

    # ---------------- UI ----------------
    def _build_ui(self):
        # 상단 툴바
        top = ttk.Frame(self, padding=8)
        top.pack(fill="x")
        ttk.Button(top, text="① PowerPoint 연결", command=self.on_connect).pack(side="left")
        ttk.Button(top, text="새로고침", command=self.render_slide).pack(side="left", padx=(8, 0))

        nav = ttk.Frame(top)
        nav.pack(side="left", padx=16)
        ttk.Button(nav, text="◀ 이전", command=lambda: self.change_slide(-1)).pack(side="left")
        self.slide_lbl = ttk.Label(nav, text="슬라이드 -/-", width=14, anchor="center")
        self.slide_lbl.pack(side="left", padx=4)
        ttk.Button(nav, text="다음 ▶", command=lambda: self.change_slide(+1)).pack(side="left")

        # 줌 컨트롤
        zoom = ttk.Frame(top)
        zoom.pack(side="left", padx=8)
        ttk.Button(zoom, text="－", width=3, command=self.zoom_out).pack(side="left")
        self.zoom_lbl = ttk.Label(zoom, text="100%", width=6, anchor="center")
        self.zoom_lbl.pack(side="left")
        ttk.Button(zoom, text="＋", width=3, command=self.zoom_in).pack(side="left")
        ttk.Button(zoom, text="화면 맞춤", command=self.zoom_fit).pack(side="left", padx=(4, 0))

        self.status = tk.StringVar(value="PowerPoint 를 연결해주세요.")
        ttk.Label(top, textvariable=self.status, foreground="#0060c0").pack(side="left", padx=12)

        # 본문
        body = ttk.Frame(self, padding=(8, 0, 8, 8))
        body.pack(fill="both", expand=True)

        # 좌: 슬라이드 캔버스 (스크롤 + 줌 + 드래그 이동)
        left = ttk.LabelFrame(
            body,
            text="슬라이드 — 사진을 순서대로 클릭 (다시 클릭=취소) · 우클릭 드래그=이동 · Ctrl+휠=확대/축소",
            padding=6)
        left.pack(side="left", fill="both", expand=True)

        cv_wrap = ttk.Frame(left)
        cv_wrap.pack(fill="both", expand=True)
        cv_wrap.rowconfigure(0, weight=1)
        cv_wrap.columnconfigure(0, weight=1)

        self.canvas = tk.Canvas(cv_wrap, bg="#3c3c3c", highlightthickness=0)
        hbar = ttk.Scrollbar(cv_wrap, orient="horizontal", command=self.canvas.xview)
        vbar = ttk.Scrollbar(cv_wrap, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(xscrollcommand=hbar.set, yscrollcommand=vbar.set)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        vbar.grid(row=0, column=1, sticky="ns")
        hbar.grid(row=1, column=0, sticky="ew")

        # 좌클릭 = 사진 선택
        self.canvas.bind("<Button-1>", self.on_canvas_click)
        # 우클릭 드래그 = 화면 이동(panning)
        self.canvas.bind("<ButtonPress-3>", lambda e: self.canvas.scan_mark(e.x, e.y))
        self.canvas.bind("<B3-Motion>", lambda e: self.canvas.scan_dragto(e.x, e.y, gain=1))
        # 가운데 버튼 드래그도 이동
        self.canvas.bind("<ButtonPress-2>", lambda e: self.canvas.scan_mark(e.x, e.y))
        self.canvas.bind("<B2-Motion>", lambda e: self.canvas.scan_dragto(e.x, e.y, gain=1))
        # 휠 스크롤 / Ctrl+휠 줌
        self.canvas.bind("<MouseWheel>", self._on_wheel)
        self.canvas.bind("<Shift-MouseWheel>", self._on_shift_wheel)
        self.canvas.bind("<Control-MouseWheel>", self._on_ctrl_wheel)
        # Linux 휠 (참고용)
        self.canvas.bind("<Button-4>", lambda e: self.canvas.yview_scroll(-1, "units"))
        self.canvas.bind("<Button-5>", lambda e: self.canvas.yview_scroll(1, "units"))

        # 우: 순서 + 새 파일
        right = ttk.LabelFrame(body, text="교체 순서 / 새 사진", padding=6)
        right.pack(side="left", fill="both", expand=False, padx=(8, 0))
        right.configure(width=360)

        info = ttk.Label(right, foreground="#555", justify="left",
                         text="왼쪽 슬라이드에서 사진을 클릭한 순서대로\n아래 목록이 채워집니다.")
        info.pack(anchor="w", pady=(0, 6))

        # 스크롤 가능한 행 영역
        wrap = ttk.Frame(right)
        wrap.pack(fill="both", expand=True)
        self.rows_canvas = tk.Canvas(wrap, highlightthickness=0, width=340)
        vs = ttk.Scrollbar(wrap, orient="vertical", command=self.rows_canvas.yview)
        self.rows_frame = ttk.Frame(self.rows_canvas)
        self.rows_frame.bind(
            "<Configure>",
            lambda e: self.rows_canvas.configure(scrollregion=self.rows_canvas.bbox("all")))
        self.rows_canvas.create_window((0, 0), window=self.rows_frame, anchor="nw")
        self.rows_canvas.configure(yscrollcommand=vs.set)
        self.rows_canvas.pack(side="left", fill="both", expand=True)
        vs.pack(side="right", fill="y")

        rbtn = ttk.Frame(right)
        rbtn.pack(fill="x", pady=(6, 0))
        ttk.Button(rbtn, text="전체 초기화", command=self.clear_sequence).pack(side="left")

        # 하단
        bottom = ttk.Frame(self, padding=8)
        bottom.pack(fill="x")
        self.keep_aspect = tk.BooleanVar(value=False)
        ttk.Checkbutton(bottom, text="원본 비율 유지 (박스 안에 맞춤)",
                        variable=self.keep_aspect).pack(side="left")
        self.lock_ratio = tk.BooleanVar(value=True)
        ttk.Checkbutton(bottom, text="가로세로 비율 잠금 (교체된 사진에 자동 적용)",
                        variable=self.lock_ratio).pack(side="left", padx=(12, 0))
        ttk.Button(bottom, text="② 교체 실행", command=self.on_replace).pack(side="right")
        self.count_lbl = tk.StringVar(value="선택 0장")
        ttk.Label(bottom, textvariable=self.count_lbl).pack(side="right", padx=12)

        if not HAS_PIL:
            messagebox.showwarning(
                "Pillow 필요",
                "이 방식은 Pillow 가 필요합니다.\n\n    pip install Pillow\n\n"
                "설치 후 다시 실행해주세요.")

    # ---------------- 동작 ----------------
    def on_connect(self):
        try:
            name = self.ctrl.connect()
            self.slide_index = self.ctrl.current_slide_index()
            self.status.set(f"연결됨: {name}")
            self.render_slide()
        except Exception as e:
            messagebox.showerror("연결 실패", str(e))
            self.status.set("연결 실패")

    def change_slide(self, delta):
        if self.ctrl.pres is None:
            return
        n = self.ctrl.slide_count()
        self.slide_index = max(1, min(n, self.slide_index + delta))
        self.render_slide()

    def _fit_zoom(self, w_pt, h_pt):
        self.canvas.update_idletasks()
        vw = self.canvas.winfo_width()
        vh = self.canvas.winfo_height()
        if vw <= 1 or vh <= 1:      # 아직 그려지기 전
            vw, vh = 820, 520
        return max(self.ZOOM_MIN, min(vw / w_pt, vh / h_pt) * 0.98)

    def zoom_in(self):
        if self.zoom:
            self.zoom = min(self.ZOOM_MAX, self.zoom * self.ZOOM_STEP)
            self.render_slide()

    def zoom_out(self):
        if self.zoom:
            self.zoom = max(self.ZOOM_MIN, self.zoom / self.ZOOM_STEP)
            self.render_slide()

    def zoom_fit(self):
        self.zoom = None            # render 에서 맞춤 재계산
        self.render_slide()

    def _on_wheel(self, event):
        self.canvas.yview_scroll(-1 if event.delta > 0 else 1, "units")

    def _on_shift_wheel(self, event):
        self.canvas.xview_scroll(-1 if event.delta > 0 else 1, "units")

    def _on_ctrl_wheel(self, event):
        if event.delta > 0:
            self.zoom_in()
        else:
            self.zoom_out()

    def render_slide(self):
        if self.ctrl.pres is None or not HAS_PIL:
            return
        n = self.ctrl.slide_count()
        self.slide_index = max(1, min(n, self.slide_index))
        self.slide_lbl.configure(text=f"슬라이드 {self.slide_index}/{n}")

        w_pt, h_pt = self.ctrl.slide_size_pt()
        if self.zoom is None:
            self.zoom = self._fit_zoom(w_pt, h_pt)
        self.scale = self.zoom
        disp_w = max(1, int(w_pt * self.zoom))
        disp_h = max(1, int(h_pt * self.zoom))
        self.zoom_lbl.configure(text=f"{int(self.zoom * 100)}%")

        png = os.path.join(self._tmp, f"slide_{self.slide_index}.png")
        try:
            # 선명하게 1.5배로 내보낸 뒤 표시 크기로 축소 (과도한 해상도 방지)
            ex_w = min(disp_w * 3, max(disp_w, 2400))
            ex_h = int(ex_w * h_pt / w_pt)
            self.ctrl.export_slide_png(self.slide_index, png, ex_w, ex_h)
            img = Image.open(png).resize((disp_w, disp_h), Image.LANCZOS)
            self.slide_photo = ImageTk.PhotoImage(img)
        except Exception as e:
            self.status.set(f"슬라이드 표시 실패: {e}")
            return

        # 스크롤 영역을 슬라이드 크기로 설정
        self.canvas.configure(scrollregion=(0, 0, disp_w, disp_h))
        self.overlays = [
            (
                p["left"] * self.scale, p["top"] * self.scale,
                (p["left"] + p["width"]) * self.scale,
                (p["top"] + p["height"]) * self.scale,
                p,
            )
            for p in self.ctrl.pictures_on_slide(self.slide_index)
        ]
        self._draw()

    def _draw(self):
        self.canvas.delete("all")
        if self.slide_photo is not None:
            self.canvas.create_image(0, 0, anchor="nw", image=self.slide_photo)

        for (x0, y0, x1, y1, target) in self.overlays:
            order = self._order_of(target)
            if order is None:
                self.canvas.create_rectangle(x0, y0, x1, y1, outline="#4da3ff",
                                             width=2, dash=(4, 3))
            else:
                self.canvas.create_rectangle(x0, y0, x1, y1, outline=self.BADGE_COLORS, width=3)
                r = 15
                cx, cy = x0 + r + 3, y0 + r + 3
                self.canvas.create_oval(cx - r, cy - r, cx + r, cy + r,
                                        fill=self.BADGE_COLORS, outline="white", width=2)
                self.canvas.create_text(cx, cy, text=str(order), fill="white",
                                        font=("Segoe UI", 12, "bold"))

    def on_canvas_click(self, event):
        if not self.overlays:
            return
        # 스크롤/줌 상태를 반영한 실제 캔버스 좌표로 변환
        cx = self.canvas.canvasx(event.x)
        cy = self.canvas.canvasy(event.y)
        hit = None
        best_area = None
        for (x0, y0, x1, y1, target) in self.overlays:
            if x0 <= cx <= x1 and y0 <= cy <= y1:
                area = (x1 - x0) * (y1 - y0)
                if best_area is None or area < best_area:
                    best_area = area
                    hit = target
        if hit is None:
            return
        idx = self._index_in_sequence(hit)
        if idx is None:
            hit.setdefault("image", None)
            self.sequence.append(hit)
        else:
            self.sequence.pop(idx)
        self._draw()
        self._rebuild_rows()

    def _key(self, t):
        return (t["slide_index"], t["shape_id"])

    def _index_in_sequence(self, target):
        for i, t in enumerate(self.sequence):
            if self._key(t) == self._key(target):
                return i
        return None

    def _order_of(self, target):
        i = self._index_in_sequence(target)
        return None if i is None else i + 1

    def clear_sequence(self):
        self.sequence = []
        self._draw()
        self._rebuild_rows()

    def _rebuild_rows(self):
        for w in self.rows_frame.winfo_children():
            w.destroy()
        self._row_thumbs = []

        for i, t in enumerate(self.sequence):
            row = ttk.Frame(self.rows_frame, padding=4)
            row.pack(fill="x", pady=2)

            ttk.Label(row, text=str(i + 1), width=2, anchor="center",
                      background=self.BADGE_COLORS, foreground="white").pack(side="left")

            thumb = tk.Label(row, width=8, height=3, relief="groove", bg="#f0f0f0")
            thumb.pack(side="left", padx=6)
            if t.get("image") and HAS_PIL:
                try:
                    im = Image.open(t["image"])
                    im.thumbnail((64, 48))
                    ph = ImageTk.PhotoImage(im)
                    self._row_thumbs.append(ph)
                    thumb.configure(image=ph, width=64, height=48)
                except Exception:
                    pass

            mid = ttk.Frame(row)
            mid.pack(side="left", fill="x", expand=True)
            name = os.path.basename(t["image"]) if t.get("image") else "(새 파일 없음)"
            ttk.Label(mid, text=f"[슬라이드 {t['slide_index']}]", foreground="#888").pack(anchor="w")
            ttk.Label(mid, text=name, wraplength=140).pack(anchor="w")

            ttk.Button(row, text="파일 선택",
                       command=lambda tt=t: self._pick_file(tt)).pack(side="right")

        self.count_lbl.set(f"선택 {len(self.sequence)}장")

    def _pick_file(self, target):
        path = filedialog.askopenfilename(title="새 이미지 선택", filetypes=IMAGE_TYPES)
        if path:
            target["image"] = path
            self._rebuild_rows()

    def on_replace(self):
        if not self.sequence:
            messagebox.showwarning("안내", "슬라이드에서 바꿀 사진을 먼저 클릭해주세요.")
            return
        missing = [i + 1 for i, t in enumerate(self.sequence) if not t.get("image")]
        if missing:
            if not messagebox.askyesno(
                "새 파일 없음",
                f"{missing} 번 항목에 새 파일이 지정되지 않았습니다.\n"
                f"해당 항목은 건너뛰고 나머지만 교체할까요?"):
                return

        keep = self.keep_aspect.get()
        lock = self.lock_ratio.get()
        ok, fail = 0, []
        for i, t in enumerate(self.sequence):
            if not t.get("image"):
                continue
            try:
                self.ctrl.replace_picture(t, t["image"], keep_aspect=keep, lock_ratio=lock)
                ok += 1
            except Exception as e:
                fail.append(f"{i + 1}번: {e}")

        msg = f"완료: {ok}장 교체됨."
        if fail:
            messagebox.showwarning("결과", msg + "\n\n실패:\n" + "\n".join(fail))
        else:
            messagebox.showinfo("결과", msg)
        self.status.set(f"교체 완료: {ok}장")
        self.clear_sequence()
        self.render_slide()


def main():
    if sys.platform != "win32":
        print("이 프로그램은 Windows + PowerPoint 환경에서 실행해야 합니다.")
    try:
        App().mainloop()
    except Exception:
        traceback.print_exc()
        input("오류가 발생했습니다. Enter 를 눌러 종료...")


if __name__ == "__main__":
    main()
