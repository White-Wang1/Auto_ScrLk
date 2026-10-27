#!/usr/bin/env python3
"""
本机屏幕变化自动截屏并保存到 PPT 的 GUI 工具（带暂停与内容去重）
==============================================================
打开程序后自动新建一个本地 PPT 文件，可自行框定一个"检测窗口"（位置与大小均可调整），
程序实时检测该窗口内的画面变化，一旦检测到变化就自动截取该区域并作为一页写入 PPT；
并且会对本次会话内已经写入 PPT 的所有截图做"内容对比去重"：若新画面与已有某一页
画面相同（如窗口切回来又切走、返回首页等），则自动跳过，不生成重复页。

功能特性:
  * 打开即自动新建本地 PPT（宽屏 16:9）
  * 全屏框选检测区域（鼠标拖拽画框，可二次用数值微调位置与大小）
  * 实时垂直刷新检测：周期性截取检测区域并与上一帧比对，变化超阈值即截图入 PPT
  * 内容去重：每个新画面与本次已入 PPT 的所有页面逐一比对，相同画面自动删除/跳过
  * 可点击控制：开始检测 / 暂停(继续) / 停止，随时保存并退出
  * 捕获张数与去重跳过数实时显示

用法:
    python scripts/screen_capture_gui.py [--output ./output] [--interval 0.3] \
        [--threshold 1.5] [--dedup 1.0] [--region "x,y,w,h"]

参数:
    --output      PPT 及截图的保存目录，默认 ./output
    --interval    检测采样间隔(秒)，默认 0.3
    --threshold   画面变化判定阈值，默认 1.5（越大越不敏感）
    --dedup       内容去重阈值，默认 1.0（越小越严格，越小越容易判定'相同'）
    --region      "left,top,width,height" 预置检测区域（跳过框选，便于自动化/复用记忆）
"""
import argparse
import logging
import os
import signal
import sys
import time
import queue
import threading
from typing import List, Optional, Tuple

import cv2
import numpy as np
from PIL import Image
from pptx import Presentation
from pptx.util import Inches
import tkinter as tk
from tkinter import messagebox, ttk

logging.basicConfig(level=logging.INFO, format="[%(asctime)s] %(levelname)s %(message)s")
logger = logging.getLogger("screen_capture_gui")

SLIDE_WIDTH_IN = 13.333
SLIDE_HEIGHT_IN = 7.5
COMPARE_WIDTH = 120
COMPARE_HEIGHT = 68


def detect_display() -> dict:
    """返回主显示器区域 (left, top, width, height)。"""
    from mss import MSS
    with MSS() as sct:
        return dict(sct.monitors[1])


# 复用全局 mss 连接，避免每帧重复建立/销毁连接的开销
_MS_SESSION = None


def _mss_session():
    global _MS_SESSION
    if _MS_SESSION is None:
        from mss import MSS
        _MS_SESSION = MSS()
    return _MS_SESSION


def grab_region(left: int, top: int, width: int, height: int) -> np.ndarray:
    """截取屏幕指定区域，返回 BGR numpy 数组（复用全局连接）。"""
    sct = _mss_session()
    shot = sct.grab({"left": left, "top": top, "width": width, "height": height})
    img = np.asarray(shot)
    return img[:, :, :3].copy()


def small_gray(bgr: np.ndarray) -> np.ndarray:
    """转降采样灰度，用于变化比较。"""
    gray = cv2.cvtColor(bgr, cv2.COLOR_BGR2GRAY)
    return cv2.resize(gray, (COMPARE_WIDTH, COMPARE_HEIGHT), interpolation=cv2.INTER_AREA).astype(np.float32)


def change_score(a: np.ndarray, b: np.ndarray) -> float:
    """计算两帧降采样灰度图的平均绝对差。"""
    diff = cv2.absdiff(a, b)
    return float(cv2.mean(diff)[0])


def make_ppt() -> Presentation:
    """新建一个宽屏(16:9)空 PPT。"""
    prs = Presentation()
    prs.slide_width = Inches(SLIDE_WIDTH_IN)
    prs.slide_height = Inches(SLIDE_HEIGHT_IN)
    return prs


def save_ppt(prs: Presentation, path: str) -> None:
    """把 PPT 保存到本地，若已有图片则逐页保留。"""
    prs.save(path)


class CaptureApp:
    def __init__(self, root: tk.Tk, output_dir: str, interval: float, threshold: float,
                 dedup: float, region: Optional[Tuple[int, int, int, int]]):
        self.root = root
        self.output_dir = output_dir
        self.interval = interval
        self.threshold = threshold
        self.dedup_threshold = dedup
        os.makedirs(output_dir, exist_ok=True)
        self.frame_dir = os.path.join(output_dir, "frames")
        os.makedirs(self.frame_dir, exist_ok=True)

        # 新建本地 PPT 文件
        self.ppt_path = os.path.join(output_dir, "screenshots.pptx")
        self.prs = make_ppt()
        save_ppt(self.prs, self.ppt_path)  # 打开即自动新建本地 PPT
        self.capture_count = 0
        self.dup_skipped = 0
        self.screenshots: List[str] = []

        # 检测区域
        self.display = detect_display()
        if region:
            self.region = list(region)
        else:
            self.region = [0, 0, max(200, self.display["width"] // 2), max(150, self.display["height"] // 2)]

        self.running = False
        self.paused = False
        self.prev_gray: Optional[np.ndarray] = None
        # 已写入 PPT 的所有截图指纹（降采样灰度），用于内容去重
        self._saved_sigs: List[np.ndarray] = []

        # 后台异步写 PPT 的队列与线程
        self._ppt_queue: "queue.Queue[str]" = queue.Queue()
        self._ppt_worker = threading.Thread(target=self._ppt_worker_loop, daemon=True)
        self._ppt_worker.start()

        self._build_ui()
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        signal.signal(signal.SIGINT, self._sig_stop)

    # ---------- UI ----------
    def _build_ui(self) -> None:
        self.root.title("屏幕变化截屏 → PPT（自动去重）")
        self.root.geometry("520x600")
        self.root.resizable(False, False)

        pad = {"padx": 10, "pady": 4}
        frm = ttk.Frame(self.root, padding=10)
        frm.pack(fill="both", expand=True)

        ttk.Label(frm, text="检测区域", font=("", 11, "bold")).grid(row=0, column=0, sticky="w", **pad)
        ttk.Button(frm, text="框选区域", command=self.open_region_selector).grid(row=0, column=1, sticky="e", **pad)

        pos_frame = ttk.LabelFrame(frm, text="区域坐标与大小（可直接修改数值调整）", padding=8)
        pos_frame.grid(row=1, column=0, columnspan=2, sticky="ew", **pad)
        ttk.Label(pos_frame, text="X").grid(row=0, column=0)
        self.x_var = tk.IntVar(value=self.region[0])
        self.y_var = tk.IntVar(value=self.region[1])
        self.w_var = tk.IntVar(value=self.region[2])
        self.h_var = tk.IntVar(value=self.region[3])
        ttk.Entry(pos_frame, textvariable=self.x_var, width=7).grid(row=0, column=1)
        ttk.Label(pos_frame, text="Y").grid(row=0, column=2)
        ttk.Entry(pos_frame, textvariable=self.y_var, width=7).grid(row=0, column=3)
        ttk.Label(pos_frame, text="宽").grid(row=0, column=4)
        ttk.Entry(pos_frame, textvariable=self.w_var, width=7).grid(row=0, column=5)
        ttk.Label(pos_frame, text="高").grid(row=0, column=6)
        ttk.Entry(pos_frame, textvariable=self.h_var, width=7).grid(row=0, column=7)
        ttk.Button(pos_frame, text="应用", command=self.apply_region).grid(row=0, column=8, padx=(8, 0))

        # 检测区域缩略预览
        self.preview_label = ttk.Label(frm, text="（点击'框选区域'或输入坐标后，此处显示检测区域预览）",
                                       relief="groove", anchor="center")
        self.preview_label.grid(row=2, column=0, columnspan=2, sticky="nsew", **pad)

        param_frame = ttk.LabelFrame(frm, text="检测参数", padding=8)
        param_frame.grid(row=3, column=0, columnspan=2, sticky="ew", **pad)
        self.interval_var = tk.DoubleVar(value=self.interval)
        self.threshold_var = tk.DoubleVar(value=self.threshold)
        self.dedup_var = tk.DoubleVar(value=self.dedup_threshold)
        ttk.Label(param_frame, text="采样间隔(秒)").grid(row=0, column=0, sticky="w")
        ttk.Scale(param_frame, from_=0.1, to=2.0, variable=self.interval_var, orient="horizontal").grid(row=0, column=1, sticky="ew", padx=8)
        ttk.Label(param_frame, text=f"{self.interval_var.get():.2f}").grid(row=0, column=2)
        ttk.Label(param_frame, text="变化阈值").grid(row=1, column=0, sticky="w", pady=(6, 0))
        ttk.Scale(param_frame, from_=0.5, to=10.0, variable=self.threshold_var, orient="horizontal").grid(row=1, column=1, sticky="ew", padx=8, pady=(6, 0))
        ttk.Label(param_frame, text=f"{self.threshold_var.get():.1f}").grid(row=1, column=2, pady=(6, 0))
        ttk.Label(param_frame, text="去重阈值").grid(row=2, column=0, sticky="w", pady=(6, 0))
        ttk.Scale(param_frame, from_=0.1, to=3.0, variable=self.dedup_var, orient="horizontal").grid(row=2, column=1, sticky="ew", padx=8, pady=(6, 0))
        ttk.Label(param_frame, text=f"{self.dedup_var.get():.1f}").grid(row=2, column=2, pady=(6, 0))
        ttk.Label(param_frame, text="去重阈值越小越严格：新画面与已保存任一页差异低于该值时视为相同，自动删除该页",
                  foreground="#888").grid(row=3, column=0, columnspan=3, sticky="w")

        # 控制
        ctrl = ttk.Frame(frm)
        ctrl.grid(row=4, column=0, columnspan=2, sticky="ew", **pad)
        self.start_btn = ttk.Button(ctrl, text="▶ 开始检测", command=self.start_capture)
        self.start_btn.pack(side="left", expand=True, fill="x", padx=(0, 4))
        self.pause_btn = ttk.Button(ctrl, text="⏸ 暂停", command=self.toggle_pause, state="disabled")
        self.pause_btn.pack(side="left", expand=True, fill="x", padx=(2, 2))
        self.stop_btn = ttk.Button(ctrl, text="■ 停止", command=self.stop_capture, state="disabled")
        self.stop_btn.pack(side="left", expand=True, fill="x", padx=(4, 0))

        # 状态
        self.status_var = tk.StringVar(value=f"检测区域: {self.region} | 已捕获 0 张 | 去重跳过 0 | PPT: {self.ppt_path}")
        ttk.Label(frm, textvariable=self.status_var, anchor="w", wraplength=490).grid(row=5, column=0, columnspan=2, sticky="w", **pad)

        self._refresh_preview()

    def _refresh_preview(self) -> None:
        try:
            img = grab_region(*self.region)
            rgb = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
            pil = Image.fromarray(rgb)
            pil.thumbnail((440, 200))
            from PIL import ImageTk
            self._preview_img = ImageTk.PhotoImage(pil)
            self.preview_label.config(image=self._preview_img, text="")
        except Exception as e:
            logger.warning("预览刷新失败: %s", e)
            self._preview_img = None

    # ---------- 区域框选 ----------
    def open_region_selector(self) -> None:
        d = self.display
        sel = tk.Toplevel(self.root)
        sel.attributes("-fullscreen", True)
        sel.attributes("-alpha", 0.25)
        sel.configure(bg="black")
        canvas = tk.Canvas(sel, cursor="crosshair")
        canvas.pack(fill="both", expand=True)

        rect = {"x0": 0, "y0": 0, "x1": 0, "y1": 0, "id": None, "active": False}

        def on_press(e):
            rect["active"] = True
            rect["x0"], rect["y0"] = e.x, e.y
            rect["x1"], rect["y1"] = e.x, e.y
            if rect["id"]:
                canvas.delete(rect["id"])
            rect["id"] = canvas.create_rectangle(e.x, e.y, e.x, e.y, outline="#00ff88", width=2)

        def on_drag(e):
            if rect["active"]:
                rect["x1"], rect["y1"] = e.x, e.y
                canvas.coords(rect["id"], rect["x0"], rect["y0"], rect["x1"], rect["y1"])

        def on_release(e):
            rect["active"] = False
            x0, y0 = min(rect["x0"], rect["x1"]), min(rect["y0"], rect["y1"])
            x1, y1 = max(rect["x0"], rect["x1"]), max(rect["y0"], rect["y1"])
            w, h = x1 - x0, y1 - y0
            if w > 10 and h > 10:
                self.region = [int(x0), int(y0), int(w), int(h)]
                self.x_var.set(self.region[0])
                self.y_var.set(self.region[1])
                self.w_var.set(self.region[2])
                self.h_var.set(self.region[3])
                self._refresh_preview()
                self._update_status()
            sel.destroy()

        canvas.bind("<ButtonPress-1>", on_press)
        canvas.bind("<B1-Motion>", on_drag)
        canvas.bind("<ButtonRelease-1>", on_release)

    def apply_region(self) -> None:
        try:
            x = max(0, int(self.x_var.get()))
            y = max(0, int(self.y_var.get()))
            w = max(20, int(self.w_var.get()))
            h = max(20, int(self.h_var.get()))
            self.region = [x, y, w, h]
            self._refresh_preview()
            self._update_status()
        except Exception:
            messagebox.showerror("错误", "请输入有效的整数坐标")

    def _update_status(self) -> None:
        state = "已暂停" if self.paused else ("检测中" if self.running else "空闲")
        self.status_var.set(
            f"状态: {state} | 检测区域: {self.region} | 已捕获 {self.capture_count} 张"
            f" | 去重跳过 {self.dup_skipped} | PPT: {self.ppt_path}"
        )

    # ---------- 控制：开始 / 暂停 / 停止 ----------
    def start_capture(self) -> None:
        self.apply_region()
        self.running = True
        self.paused = False
        self.prev_gray = None
        self.start_btn.config(state="disabled")
        self.pause_btn.config(state="normal", text="⏸ 暂停")
        self.stop_btn.config(state="normal")
        self._update_status()
        self._tick()

    def toggle_pause(self) -> None:
        """点击在'暂停'与'继续'之间切换。"""
        if not self.running:
            return
        self.paused = not self.paused
        if self.paused:
            self.pause_btn.config(text="▶ 继续")
            # 暂停时清空上一帧，继续后重新以首帧逻辑捕获
            self.prev_gray = None
        else:
            self.pause_btn.config(text="⏸ 暂停")
            self.prev_gray = None
        self._update_status()

    # ---------- 检测循环 ----------
    def _tick(self) -> None:
        if not self.running:
            return
        if self.paused:
            # 暂停期间不截屏，仅保持定时器轮询，便于随时"继续"
            self.root.after(int(self.interval_var.get() * 1000), self._tick)
            return
        try:
            frame = grab_region(*self.region)
            cur = small_gray(frame)
            if self.prev_gray is not None:
                score = change_score(cur, self.prev_gray)
            else:
                score = 0.0
            # 首帧，或变化超过阈值 → 视为候选，进一步做内容去重
            is_candidate = self.prev_gray is None or score >= self.threshold_var.get()
            if is_candidate:
                duplicate = False
                for sig in self._saved_sigs:
                    if change_score(cur, sig) < self.dedup_var.get():
                        duplicate = True
                        break
                if duplicate:
                    # 与本次已保存的某一页内容相同 → 去重，不写入 PPT
                    self.dup_skipped += 1
                    logger.info("检测到与已有页相同内容，去重跳过(差异 %.1f)", score)
                else:
                    self.capture_count += 1
                    name = f"shot_{self.capture_count:04d}.png"
                    path = os.path.join(self.frame_dir, name)
                    cv2.imwrite(path, frame)
                    self.screenshots.append(path)
                    self._saved_sigs.append(cur.copy())  # 记录指纹用于后续去重
                    self._ppt_queue.put(path)  # 入队，不阻塞检测循环
                    logger.info("捕获画面变化(%.1f): %s", score, path)
            self.prev_gray = cur
        except Exception as e:
            logger.error("检测失败: %s", e)
        self._update_status()
        self.root.after(int(self.interval_var.get() * 1000), self._tick)

    def _ppt_worker_loop(self) -> None:
        """后台线程：串行消费队列，把截图追加进 PPT 并落盘，避免阻塞检测。"""
        while True:
            try:
                img_path = self._ppt_queue.get(timeout=0.5)
            except queue.Empty:
                continue
            try:
                with Image.open(img_path) as im:
                    img_w, img_h = im.size
                blank = self.prs.slide_layouts[6]
                slide = self.prs.slides.add_slide(blank)
                ratio_w = int(self.prs.slide_width) / img_w
                ratio_h = int(self.prs.slide_height) / img_h
                scale = max(ratio_w, ratio_h)
                width = int(img_w * scale)
                height = int(img_h * scale)
                left = int((int(self.prs.slide_width) - width) / 2)
                top = int((int(self.prs.slide_height) - height) / 2)
                slide.shapes.add_picture(img_path, left, top, width=width, height=height)
                # 实时落盘到本地 PPT（最终 PPT 包含所有已捕获截图）
                save_ppt(self.prs, self.ppt_path)
            except Exception as e:
                logger.error("写入PPT失败: %s", e)
            finally:
                self._ppt_queue.task_done()

    def _flush_ppt(self) -> None:
        """等待后台队列处理完毕，确保 PPT 完整后落盘。"""
        try:
            self._ppt_queue.join()
            save_ppt(self.prs, self.ppt_path)
        except Exception as e:
            logger.error("保存PPT失败: %s", e)

    def stop_capture(self) -> None:
        self.running = False
        self.paused = False
        self.start_btn.config(state="normal")
        self.pause_btn.config(state="disabled", text="⏸ 暂停")
        self.stop_btn.config(state="disabled")
        self._flush_ppt()
        self._update_status()
        logger.info("已停止，PPT 已保存: %s（共 %d 张，去重跳过 %d）",
                    self.ppt_path, self.capture_count, self.dup_skipped)

    def on_close(self) -> None:
        self.running = False
        self.paused = False
        if self.capture_count > 0:
            self._flush_ppt()
        self.root.destroy()

    def _sig_stop(self, *args):  # type: ignore
        self.running = False


def main() -> int:
    parser = argparse.ArgumentParser(description="本机屏幕变化自动截屏并保存到PPT的GUI工具（支持暂停与去重）")
    parser.add_argument("--output", type=str, default="./output", help="保存目录，默认 ./output")
    parser.add_argument("--interval", type=float, default=0.3, help="检测采样间隔(秒)，默认0.3")
    parser.add_argument("--threshold", type=float, default=1.5, help="画面变化判定阈值，默认1.5")
    parser.add_argument("--dedup", type=float, default=1.0, help="内容去重阈值，默认1.0（越小越严格）")
    parser.add_argument("--region", type=str, default=None,
                        help='"left,top,width,height" 预置检测区域，如 "0,0,800,450"')
    args = parser.parse_args()

    region: Optional[Tuple[int, int, int, int]] = None
    if args.region:
        try:
            parts = [int(p.strip()) for p in args.region.split(",")]
            if len(parts) != 4:
                raise ValueError
            region = (parts[0], parts[1], parts[2], parts[3])
        except Exception:
            logger.error("--region 格式错误，应为 left,top,width,height")
            return 1

    root = tk.Tk()
    app = CaptureApp(root, args.output, args.interval, args.threshold, args.dedup, region)
    # 便于自动化测试：若指定了 --region 与 --auto-stop 环境变量，自动开始
    auto = os.getenv("CAPTURE_AUTO_START")
    if auto and args.region:
        root.after(500, app.start_capture)
    root.mainloop()
    return 0


if __name__ == "__main__":
    sys.exit(main())