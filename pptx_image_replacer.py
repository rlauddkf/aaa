# -*- coding: utf-8 -*-
"""
PowerPoint 이미지 일괄 교체 프로그램
======================================

현재 "열려 있는" PowerPoint 를 그대로 제어해서,
 1) 교체할 대상 사진들을 찾아 순서를 정하고
 2) 새로 넣을 이미지 파일들을 순서대로 고른 뒤
 3) [교체 실행] 을 누르면 각 사진을 "같은 위치 · 같은 크기" 로 바꿔줍니다.

동작 환경: Windows + PowerPoint (데스크톱 버전) 설치 필요
필요 패키지: pywin32 (필수), Pillow (미리보기용, 없어도 동작)

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

# ---------------------------------------------------------------------------
# 선택 의존성 (Pillow) : 썸네일 미리보기에만 사용. 없으면 미리보기만 비활성화.
# ---------------------------------------------------------------------------
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

PP_SHAPE_FORMAT_PNG = 2   # ppShapeFormatPNG

# ZOrder 명령
MSO_SEND_BACKWARD = 3     # msoSendBackward


IMAGE_TYPES = [
    ("이미지 파일", "*.png *.jpg *.jpeg *.gif *.bmp *.tif *.tiff *.emf *.wmf"),
    ("모든 파일", "*.*"),
]


def _import_win32():
    """win32com 을 필요할 때 import (비 Windows 환경에서 파일 자체는 열리도록)."""
    import win32com.client  # noqa
    return win32com.client


# ===========================================================================
#  PowerPoint 제어 로직
# ===========================================================================
class PowerPointController:
    def __init__(self):
        self.app = None
        self.pres = None

    def connect(self):
        """실행 중인 PowerPoint 에 연결. 없으면 새로 띄운다."""
        win32com = _import_win32()
        try:
            self.app = win32com.GetActiveObject("PowerPoint.Application")
        except Exception:
            # 실행 중인 인스턴스가 없으면 새로 시작
            self.app = win32com.Dispatch("PowerPoint.Application")
        self.app.Visible = True

        if self.app.Presentations.Count == 0:
            raise RuntimeError(
                "열려 있는 프레젠테이션이 없습니다.\n"
                "PowerPoint 에서 파일을 먼저 연 뒤 다시 [PowerPoint 연결] 을 눌러주세요."
            )
        self.pres = self.app.ActivePresentation
        return self.pres.Name

    # -- 사진 수집 -------------------------------------------------------
    @staticmethod
    def _is_picture(shape):
        try:
            if shape.Type in (MSO_PICTURE, MSO_LINKED_PICTURE):
                return True
            # 그림으로 채워진 자리표시자(placeholder)도 사진으로 취급
            if shape.Type == MSO_PLACEHOLDER:
                try:
                    if shape.PictureFormat is not None and shape.Fill.Type == 6:
                        return True
                except Exception:
                    return False
        except Exception:
            return False
        return False

    def collect_pictures(self, scope="current"):
        """
        scope: 'current' = 현재 슬라이드만, 'all' = 모든 슬라이드
        반환: [{slide_index, shape_id, name, left, top, width, height, rotation}]
        """
        result = []
        if scope == "current":
            try:
                slide_indexes = [self.app.ActiveWindow.View.Slide.SlideIndex]
            except Exception:
                slide_indexes = [1]
        else:
            slide_indexes = list(range(1, self.pres.Slides.Count + 1))

        for si in slide_indexes:
            slide = self.pres.Slides.Item(si)
            for shape in slide.Shapes:
                if not self._is_picture(shape):
                    continue
                try:
                    result.append({
                        "slide_index": si,
                        "shape_id": int(shape.Id),
                        "name": str(shape.Name),
                        "left": float(shape.Left),
                        "top": float(shape.Top),
                        "width": float(shape.Width),
                        "height": float(shape.Height),
                        "rotation": float(shape.Rotation),
                    })
                except Exception:
                    continue
        return result

    def find_shape(self, slide_index, shape_id):
        slide = self.pres.Slides.Item(slide_index)
        for shape in slide.Shapes:
            try:
                if int(shape.Id) == shape_id:
                    return slide, shape
            except Exception:
                continue
        return slide, None

    def export_shape_png(self, slide_index, shape_id, out_path):
        """대상 사진을 PNG 로 내보내 미리보기에 사용."""
        _, shape = self.find_shape(slide_index, shape_id)
        if shape is None:
            return False
        shape.Export(out_path, PP_SHAPE_FORMAT_PNG)
        return True

    def select_shape(self, slide_index, shape_id):
        """PowerPoint 화면에서 해당 사진을 실제로 선택(하이라이트)."""
        try:
            self.app.ActiveWindow.View.GotoSlide(slide_index)
        except Exception:
            pass
        _, shape = self.find_shape(slide_index, shape_id)
        if shape is not None:
            try:
                shape.Select()
            except Exception:
                pass

    # -- 실제 교체 -------------------------------------------------------
    def replace_picture(self, target, image_path, keep_aspect=False):
        """
        target 위치/크기 그대로 image_path 로 교체.
        keep_aspect=True 이면 원래 박스 안에 비율 유지하여 가운데 배치.
        """
        slide, old = self.find_shape(target["slide_index"], target["shape_id"])
        if old is None:
            raise RuntimeError("대상 사진을 찾을 수 없습니다 (이미 삭제/변경됨).")

        left = float(old.Left)
        top = float(old.Top)
        width = float(old.Width)
        height = float(old.Height)
        rotation = float(old.Rotation)

        try:
            old_z = int(old.ZOrderPosition)
        except Exception:
            old_z = None

        # 새 그림 추가 (일단 원래 박스 크기로)
        new_shape = slide.Shapes.AddPicture(
            FileName=image_path,
            LinkToFile=MSO_FALSE,
            SaveWithDocument=MSO_TRUE,
            Left=left,
            Top=top,
            Width=width,
            Height=height,
        )

        # 비율 유지 옵션 : 원본 이미지 비율대로 박스 안에 맞춰 가운데 정렬
        if keep_aspect:
            self._fit_keep_aspect(new_shape, left, top, width, height)

        # 회전값 복원
        try:
            new_shape.Rotation = rotation
        except Exception:
            pass

        # 기존 사진 삭제
        try:
            new_id = int(new_shape.Id)
        except Exception:
            new_id = None
        old.Delete()

        # z-순서 복원 (기존 사진이 있던 층으로 이동)
        if old_z is not None and new_id is not None:
            self._move_to_zorder(slide, new_id, old_z)

    @staticmethod
    def _fit_keep_aspect(shape, box_left, box_top, box_w, box_h):
        try:
            shape.LockAspectRatio = MSO_TRUE
            nat_w = float(shape.Width)
            nat_h = float(shape.Height)
            if nat_w <= 0 or nat_h <= 0:
                return
            scale = min(box_w / nat_w, box_h / nat_h)
            new_w = nat_w * scale
            new_h = nat_h * scale
            shape.Width = new_w
            shape.Height = new_h
            shape.Left = box_left + (box_w - new_w) / 2.0
            shape.Top = box_top + (box_h - new_h) / 2.0
        except Exception:
            pass

    @staticmethod
    def _move_to_zorder(slide, shape_id, target_z):
        """새 그림(현재 맨 위)을 target_z 층까지 내린다."""
        # 새로 추가된 그림은 맨 위에 있으므로 SendBackward 로 내려간다.
        for _ in range(200):  # 안전 상한
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
    THUMB = 180

    def __init__(self):
        super().__init__()
        self.title("PowerPoint 사진 일괄 교체")
        self.geometry("920x620")
        self.minsize(820, 560)

        self.ctrl = PowerPointController()
        self.targets = []      # 대상 사진 목록 (dict)
        self.new_images = []   # 새 이미지 파일 경로 목록
        self._preview_photo = None
        self._tmp_dir = tempfile.mkdtemp(prefix="pptimg_")

        self._build_ui()

    # -- UI 구성 ---------------------------------------------------------
    def _build_ui(self):
        top = ttk.Frame(self, padding=8)
        top.pack(fill="x")

        ttk.Button(top, text="① PowerPoint 연결", command=self.on_connect).pack(side="left")

        self.scope_var = tk.StringVar(value="current")
        ttk.Radiobutton(top, text="현재 슬라이드", variable=self.scope_var,
                        value="current").pack(side="left", padx=(12, 0))
        ttk.Radiobutton(top, text="전체 슬라이드", variable=self.scope_var,
                        value="all").pack(side="left")

        ttk.Button(top, text="② 사진 불러오기", command=self.on_scan).pack(side="left", padx=(12, 0))

        self.status_var = tk.StringVar(value="PowerPoint 를 연결해주세요.")
        ttk.Label(top, textvariable=self.status_var, foreground="#0060c0").pack(side="left", padx=12)

        # 본문 : 좌(대상) - 중(미리보기) - 우(새 이미지)
        body = ttk.Frame(self, padding=8)
        body.pack(fill="both", expand=True)

        # 좌 : 대상 사진
        left = ttk.LabelFrame(body, text="대상 사진 (교체될 순서)  — 클릭하면 PPT 에서 선택됨", padding=6)
        left.pack(side="left", fill="both", expand=True)

        self.target_list = tk.Listbox(left, activestyle="dotbox")
        self.target_list.pack(fill="both", expand=True)
        self.target_list.bind("<<ListboxSelect>>", self.on_target_select)

        tbtn = ttk.Frame(left)
        tbtn.pack(fill="x", pady=(6, 0))
        ttk.Button(tbtn, text="▲ 위로", command=lambda: self._move(self.target_list, self.targets, -1)).pack(side="left")
        ttk.Button(tbtn, text="▼ 아래로", command=lambda: self._move(self.target_list, self.targets, +1)).pack(side="left", padx=4)
        ttk.Button(tbtn, text="제거", command=self.on_remove_target).pack(side="left")

        # 중 : 미리보기
        mid = ttk.LabelFrame(body, text="미리보기", padding=6)
        mid.pack(side="left", fill="y", padx=8)
        self.preview_label = ttk.Label(mid, text="(미리보기)", anchor="center",
                                       width=26)
        self.preview_label.pack(fill="both", expand=True)
        if not HAS_PIL:
            ttk.Label(mid, text="Pillow 미설치\n(미리보기 비활성)",
                      foreground="#a00", justify="center").pack()

        # 우 : 새 이미지
        right = ttk.LabelFrame(body, text="새 사진 (같은 순서로 매칭)", padding=6)
        right.pack(side="left", fill="both", expand=True)

        self.new_list = tk.Listbox(right, activestyle="dotbox")
        self.new_list.pack(fill="both", expand=True)
        self.new_list.bind("<<ListboxSelect>>", self.on_new_select)

        nbtn = ttk.Frame(right)
        nbtn.pack(fill="x", pady=(6, 0))
        ttk.Button(nbtn, text="파일 추가", command=self.on_add_images).pack(side="left")
        ttk.Button(nbtn, text="▲ 위로", command=lambda: self._move(self.new_list, self.new_images, -1)).pack(side="left", padx=4)
        ttk.Button(nbtn, text="▼ 아래로", command=lambda: self._move(self.new_list, self.new_images, +1)).pack(side="left")
        ttk.Button(nbtn, text="제거", command=self.on_remove_new).pack(side="left", padx=4)

        # 하단 : 옵션 + 실행
        bottom = ttk.Frame(self, padding=8)
        bottom.pack(fill="x")

        self.keep_aspect_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(bottom, text="원본 비율 유지 (박스 안에 맞춤)",
                        variable=self.keep_aspect_var).pack(side="left")

        ttk.Button(bottom, text="③ 교체 실행", command=self.on_replace).pack(side="right")
        self.match_var = tk.StringVar(value="")
        ttk.Label(bottom, textvariable=self.match_var).pack(side="right", padx=12)

        self._refresh_match_hint()

    # -- 이벤트 핸들러 ---------------------------------------------------
    def on_connect(self):
        try:
            name = self.ctrl.connect()
            self.status_var.set(f"연결됨: {name}")
        except Exception as e:
            messagebox.showerror("연결 실패", str(e))
            self.status_var.set("연결 실패")

    def on_scan(self):
        if self.ctrl.pres is None:
            messagebox.showwarning("안내", "먼저 [PowerPoint 연결] 을 눌러주세요.")
            return
        try:
            self.targets = self.ctrl.collect_pictures(self.scope_var.get())
        except Exception as e:
            messagebox.showerror("오류", f"사진을 불러오지 못했습니다.\n{e}")
            return

        self.target_list.delete(0, tk.END)
        for i, t in enumerate(self.targets, 1):
            self.target_list.insert(
                tk.END,
                f"{i}. [슬라이드 {t['slide_index']}] {t['name']}  "
                f"({int(t['width'])}×{int(t['height'])})"
            )
        self.status_var.set(f"대상 사진 {len(self.targets)}개 찾음")
        self._refresh_match_hint()

    def on_target_select(self, _event=None):
        idx = self._sel(self.target_list)
        if idx is None:
            return
        t = self.targets[idx]
        # PPT 화면에서 해당 사진 선택
        try:
            self.ctrl.select_shape(t["slide_index"], t["shape_id"])
        except Exception:
            pass
        # 미리보기 (PPT 사진을 PNG 로 내보내서 표시)
        if HAS_PIL:
            try:
                png = os.path.join(self._tmp_dir, f"t_{t['slide_index']}_{t['shape_id']}.png")
                if self.ctrl.export_shape_png(t["slide_index"], t["shape_id"], png):
                    self._show_preview(png)
            except Exception:
                self.preview_label.configure(image="", text="(미리보기 불가)")

    def on_new_select(self, _event=None):
        idx = self._sel(self.new_list)
        if idx is None:
            return
        if HAS_PIL:
            self._show_preview(self.new_images[idx])

    def on_add_images(self):
        paths = filedialog.askopenfilenames(title="새 이미지 선택", filetypes=IMAGE_TYPES)
        if not paths:
            return
        for p in paths:
            self.new_images.append(p)
            self.new_list.insert(tk.END, f"{len(self.new_images)}. {os.path.basename(p)}")
        self._renumber(self.new_list, self.new_images, name_fn=lambda p: os.path.basename(p))
        self._refresh_match_hint()

    def on_remove_target(self):
        idx = self._sel(self.target_list)
        if idx is None:
            return
        del self.targets[idx]
        self._renumber_targets()
        self._refresh_match_hint()

    def on_remove_new(self):
        idx = self._sel(self.new_list)
        if idx is None:
            return
        del self.new_images[idx]
        self._renumber(self.new_list, self.new_images, name_fn=lambda p: os.path.basename(p))
        self._refresh_match_hint()

    def on_replace(self):
        n = min(len(self.targets), len(self.new_images))
        if n == 0:
            messagebox.showwarning("안내", "대상 사진과 새 이미지를 각각 1개 이상 준비해주세요.")
            return
        if len(self.targets) != len(self.new_images):
            if not messagebox.askyesno(
                "개수 불일치",
                f"대상 {len(self.targets)}개 / 새 이미지 {len(self.new_images)}개.\n"
                f"앞에서부터 {n}쌍만 교체합니다. 계속할까요?"
            ):
                return

        keep = self.keep_aspect_var.get()
        ok, fail = 0, []
        for i in range(n):
            t = self.targets[i]
            img = self.new_images[i]
            try:
                self.ctrl.replace_picture(t, img, keep_aspect=keep)
                ok += 1
            except Exception as e:
                fail.append(f"{i+1}번: {os.path.basename(img)} → {e}")

        msg = f"완료: {ok}개 교체됨."
        if fail:
            msg += "\n\n실패:\n" + "\n".join(fail)
            messagebox.showwarning("결과", msg)
        else:
            messagebox.showinfo("결과", msg)
        self.status_var.set(f"교체 완료: {ok}개")
        # 교체 후 대상 목록 갱신
        self.on_scan()

    # -- 보조 함수 -------------------------------------------------------
    def _sel(self, listbox):
        s = listbox.curselection()
        return s[0] if s else None

    def _move(self, listbox, data, delta):
        idx = self._sel(listbox)
        if idx is None:
            return
        j = idx + delta
        if j < 0 or j >= len(data):
            return
        data[idx], data[j] = data[j], data[idx]
        if data is self.targets:
            self._renumber_targets()
        else:
            self._renumber(self.new_list, self.new_images,
                           name_fn=lambda p: os.path.basename(p))
        listbox.selection_clear(0, tk.END)
        listbox.selection_set(j)
        listbox.activate(j)

    def _renumber_targets(self):
        self.target_list.delete(0, tk.END)
        for i, t in enumerate(self.targets, 1):
            self.target_list.insert(
                tk.END,
                f"{i}. [슬라이드 {t['slide_index']}] {t['name']}  "
                f"({int(t['width'])}×{int(t['height'])})"
            )

    def _renumber(self, listbox, data, name_fn):
        listbox.delete(0, tk.END)
        for i, item in enumerate(data, 1):
            listbox.insert(tk.END, f"{i}. {name_fn(item)}")

    def _refresh_match_hint(self):
        nt, nn = len(self.targets), len(self.new_images)
        pairs = min(nt, nn)
        self.match_var.set(f"매칭 {pairs}쌍  (대상 {nt} / 새 {nn})")

    def _show_preview(self, path):
        if not HAS_PIL:
            return
        try:
            img = Image.open(path)
            img.thumbnail((self.THUMB, self.THUMB))
            self._preview_photo = ImageTk.PhotoImage(img)
            self.preview_label.configure(image=self._preview_photo, text="")
        except Exception:
            self.preview_label.configure(image="", text="(미리보기 불가)")


def main():
    if sys.platform != "win32":
        print("이 프로그램은 Windows + PowerPoint 환경에서 실행해야 합니다.")
    try:
        app = App()
        app.mainloop()
    except Exception:
        traceback.print_exc()
        input("오류가 발생했습니다. Enter 를 눌러 종료...")


if __name__ == "__main__":
    main()
