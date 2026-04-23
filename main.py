import os
import sys
import shutil
import subprocess
import tempfile
import threading
import queue
import json
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path

try:
    import yt_dlp
except ImportError:
    print("缺少依赖 yt-dlp。请在终端执行：pip install yt-dlp")
    sys.exit(1)


AUDIO_FORMATS = ["mp3", "wav", "m4a", "aac", "flac", "ogg", "opus"]
QUALITY_OPTIONS = ["320", "256", "192", "128", "96", "64"]


def _bundled_base_dirs():
    """Return candidate dirs to look for bundled binaries (dev + PyInstaller)."""
    dirs = []
    if getattr(sys, "frozen", False):
        # PyInstaller --onefile unpacks data/binaries into sys._MEIPASS
        meipass = getattr(sys, "_MEIPASS", None)
        if meipass:
            dirs.append(Path(meipass))
        # Also check next to the executable (non-onefile layout)
        dirs.append(Path(sys.executable).parent)
    else:
        dirs.append(Path(__file__).parent)
    return dirs


def _find_tool(name_unix, name_win):
    name = name_win if os.name == "nt" else name_unix
    for base in _bundled_base_dirs():
        p = base / name
        if p.exists():
            return str(p)
    return shutil.which(name_unix)


def find_ffmpeg():
    return _find_tool("ffmpeg", "ffmpeg.exe")


def find_ffprobe():
    return _find_tool("ffprobe", "ffprobe.exe")


def parse_time(s):
    """解析 HH:MM:SS / MM:SS / SS / 123.5 → 秒；空返回 None"""
    s = (s or "").strip()
    if not s:
        return None
    try:
        if ":" in s:
            parts = s.split(":")
            parts = [p.strip() for p in parts]
            if len(parts) == 2:
                return int(parts[0]) * 60 + float(parts[1])
            if len(parts) == 3:
                return int(parts[0]) * 3600 + int(parts[1]) * 60 + float(parts[2])
            raise ValueError("时间段数不对")
        return float(s)
    except Exception as e:
        raise ValueError(f"时间格式错误: {s!r} ({e})")


def format_time(sec):
    if sec is None:
        return "—"
    sec = max(0.0, float(sec))
    h = int(sec // 3600)
    m = int((sec % 3600) // 60)
    s = sec - h * 3600 - m * 60
    if h:
        return f"{h:d}:{m:02d}:{s:05.2f}"
    return f"{m:d}:{s:05.2f}"


def probe_duration(ffprobe_path, file):
    """返回秒数；失败返回 None"""
    if not ffprobe_path:
        return None
    try:
        r = subprocess.run(
            [ffprobe_path, "-v", "error", "-show_entries", "format=duration",
             "-of", "json", file],
            capture_output=True, text=True, timeout=15,
        )
        if r.returncode != 0:
            return None
        data = json.loads(r.stdout)
        return float(data["format"]["duration"])
    except Exception:
        return None


# ──────────────────────────────────────────────────────────────
#  下载 Tab
# ──────────────────────────────────────────────────────────────
class DownloadTab(ttk.Frame):
    def __init__(self, parent, ffmpeg_path, logger):
        super().__init__(parent)
        self.ffmpeg_path = ffmpeg_path
        self.logger = logger
        self.is_running = False
        self._build_ui()

    def _build_ui(self):
        frame_url = ttk.LabelFrame(self, text="视频链接（每行一个，支持 YouTube / B站 / 抖音 / Twitter 等）")
        frame_url.pack(fill="both", expand=False, padx=10, pady=(10, 5))
        self.txt_urls = tk.Text(frame_url, height=6)
        self.txt_urls.pack(fill="both", expand=True, padx=5, pady=5)

        frame_opts = ttk.Frame(self)
        frame_opts.pack(fill="x", padx=10, pady=5)

        ttk.Label(frame_opts, text="输出格式：").grid(row=0, column=0, sticky="w")
        self.var_format = tk.StringVar(value="mp3")
        ttk.Combobox(frame_opts, textvariable=self.var_format,
                     values=AUDIO_FORMATS, width=8, state="readonly").grid(row=0, column=1, padx=5)

        ttk.Label(frame_opts, text="音质 (kbps)：").grid(row=0, column=2, sticky="w", padx=(15, 0))
        self.var_quality = tk.StringVar(value="192")
        ttk.Combobox(frame_opts, textvariable=self.var_quality,
                     values=QUALITY_OPTIONS, width=8, state="readonly").grid(row=0, column=3, padx=5)

        self.var_keep_video = tk.BooleanVar(value=False)
        ttk.Checkbutton(frame_opts, text="保留原视频",
                        variable=self.var_keep_video).grid(row=0, column=4, padx=(15, 0))

        ttk.Label(frame_opts, text="输出目录：").grid(row=1, column=0, sticky="w", pady=(10, 0))
        self.var_outdir = tk.StringVar(value=str(Path.home() / "Downloads"))
        ttk.Entry(frame_opts, textvariable=self.var_outdir).grid(
            row=1, column=1, columnspan=3, sticky="we", pady=(10, 0), padx=5)
        ttk.Button(frame_opts, text="浏览...", command=self._browse).grid(row=1, column=4, pady=(10, 0))
        frame_opts.columnconfigure(3, weight=1)

        frame_btn = ttk.Frame(self)
        frame_btn.pack(fill="x", padx=10, pady=5)
        self.btn_start = ttk.Button(frame_btn, text="开始提取", command=self._start)
        self.btn_start.pack(side="left")
        ttk.Button(frame_btn, text="打开输出目录", command=self._open_outdir).pack(side="left", padx=5)

        self.progress = ttk.Progressbar(self, mode="determinate")
        self.progress.pack(fill="x", padx=10, pady=5)

        self.var_status = tk.StringVar(value="就绪")
        ttk.Label(self, textvariable=self.var_status, anchor="w").pack(fill="x", padx=10, pady=(0, 5))

    def _browse(self):
        d = filedialog.askdirectory(initialdir=self.var_outdir.get() or str(Path.home()))
        if d:
            self.var_outdir.set(d)

    def _open_outdir(self):
        d = self.var_outdir.get().strip()
        if not d or not Path(d).exists():
            messagebox.showwarning("提示", "输出目录不存在")
            return
        if os.name == "nt":
            os.startfile(d)
        elif sys.platform == "darwin":
            os.system(f'open "{d}"')
        else:
            os.system(f'xdg-open "{d}"')

    def _start(self):
        if self.is_running:
            return
        raw = self.txt_urls.get("1.0", "end").strip()
        urls = [u.strip() for u in raw.splitlines() if u.strip()]
        if not urls:
            messagebox.showwarning("提示", "请输入至少一个视频链接")
            return
        outdir = self.var_outdir.get().strip()
        if not outdir:
            messagebox.showwarning("提示", "请选择输出目录")
            return
        try:
            Path(outdir).mkdir(parents=True, exist_ok=True)
        except Exception as e:
            messagebox.showerror("错误", f"输出目录创建失败：{e}")
            return
        if not self.ffmpeg_path:
            if not messagebox.askyesno(
                "缺少 ffmpeg",
                "未检测到 ffmpeg，无法转换为所选音频格式。\n是否仍然继续？（将下载原始音频流）"
            ):
                return

        self.is_running = True
        self.btn_start.config(state="disabled")
        self.progress["value"] = 0
        self.progress["maximum"] = len(urls)

        threading.Thread(
            target=self._worker,
            args=(urls, outdir, self.var_format.get(),
                  self.var_quality.get(), self.var_keep_video.get()),
            daemon=True,
        ).start()

    def _worker(self, urls, outdir, fmt, quality, keep_video):
        ok, fail = 0, 0
        for i, url in enumerate(urls, 1):
            self.var_status.set(f"正在处理 {i}/{len(urls)}")
            self.logger(f"[{i}/{len(urls)}] {url}")
            try:
                self._download(url, outdir, fmt, quality, keep_video)
                self.logger(f"  ✓ 完成\n")
                ok += 1
            except Exception as e:
                self.logger(f"  ✗ 失败：{e}\n")
                fail += 1
            self.progress["value"] = i
        self.var_status.set(f"全部完成 — 成功 {ok}，失败 {fail}")
        self.logger(f"═══ 下载完成：成功 {ok}，失败 {fail} ═══")
        self.is_running = False
        self.btn_start.config(state="normal")

    def _download(self, url, outdir, fmt, quality, keep_video):
        outtmpl = str(Path(outdir) / "%(title)s.%(ext)s")
        # 按站点给合适的 Referer / UA，否则 B 站会返回 412，抖音会返回 403
        ua = ("Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
              "AppleWebKit/537.36 (KHTML, like Gecko) "
              "Chrome/124.0.0.0 Safari/537.36")
        lower = url.lower()
        if "bilibili." in lower or "b23.tv" in lower:
            referer = "https://www.bilibili.com/"
        elif "douyin." in lower or "iesdouyin." in lower:
            referer = "https://www.douyin.com/"
        elif "ixigua." in lower:
            referer = "https://www.ixigua.com/"
        elif "youku." in lower:
            referer = "https://www.youku.com/"
        elif "v.qq.com" in lower:
            referer = "https://v.qq.com/"
        elif "weibo." in lower:
            referer = "https://weibo.com/"
        elif "twitter." in lower or "x.com" in lower:
            referer = "https://twitter.com/"
        else:
            referer = None
        http_headers = {"User-Agent": ua}
        if referer:
            http_headers["Referer"] = referer

        opts = {
            "format": "bestaudio/best",
            "outtmpl": outtmpl,
            "quiet": True,
            "no_warnings": True,
            "noprogress": True,
            "progress_hooks": [self._hook],
            "keepvideo": keep_video,
            "http_headers": http_headers,
            "retries": 5,
            "fragment_retries": 5,
        }
        if self.ffmpeg_path:
            opts["ffmpeg_location"] = self.ffmpeg_path
            opts["postprocessors"] = [{
                "key": "FFmpegExtractAudio",
                "preferredcodec": fmt,
                "preferredquality": quality,
            }]
        with yt_dlp.YoutubeDL(opts) as ydl:
            ydl.download([url])

    def _hook(self, d):
        status = d.get("status")
        if status == "downloading":
            pct = d.get("_percent_str", "").strip()
            speed = d.get("_speed_str", "").strip()
            self.var_status.set(f"下载中 {pct}  {speed}")
        elif status == "finished":
            fn = Path(d.get("filename", "")).name
            self.logger(f"  下载完成：{fn}")
            self.var_status.set("正在转换音频...")


# ──────────────────────────────────────────────────────────────
#  片段添加 / 编辑 对话框
# ──────────────────────────────────────────────────────────────
class ClipDialog(tk.Toplevel):
    def __init__(self, parent, ffprobe_path, initial=None):
        super().__init__(parent)
        self.ffprobe_path = ffprobe_path
        self.result = None
        self.src_dur = None  # 源文件总时长
        self.title("编辑片段" if initial else "添加片段")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()

        init = initial or {}

        frm = ttk.Frame(self, padding=12)
        frm.pack(fill="both", expand=True)

        ttk.Label(frm, text="音频/视频文件：").grid(row=0, column=0, sticky="w", pady=2)
        self.var_file = tk.StringVar(value=init.get("file", ""))
        self.var_file.trace_add("write", lambda *_: self._on_file_change())
        ent = ttk.Entry(frm, textvariable=self.var_file, width=50)
        ent.grid(row=0, column=1, columnspan=2, sticky="we", padx=5, pady=2)
        ttk.Button(frm, text="浏览...", command=self._browse).grid(row=0, column=3, pady=2)

        self.var_duration = tk.StringVar(value="源文件时长: —")
        ttk.Label(frm, textvariable=self.var_duration, foreground="gray").grid(
            row=1, column=0, columnspan=4, sticky="w", pady=(2, 8))

        ttk.Label(frm, text="起始时间：").grid(row=2, column=0, sticky="w", pady=2)
        self.var_start = tk.StringVar(value=init.get("start", ""))
        ttk.Entry(frm, textvariable=self.var_start, width=16).grid(row=2, column=1, sticky="w", padx=5)
        ttk.Label(frm, text="(空=从头, 格式: mm:ss 或 hh:mm:ss 或 秒)",
                  foreground="gray").grid(row=2, column=2, columnspan=2, sticky="w")

        ttk.Label(frm, text="结束时间：").grid(row=3, column=0, sticky="w", pady=2)
        self.var_end = tk.StringVar(value=init.get("end", ""))
        ttk.Entry(frm, textvariable=self.var_end, width=16).grid(row=3, column=1, sticky="w", padx=5)
        ttk.Label(frm, text="(空=到尾)", foreground="gray").grid(row=3, column=2, columnspan=2, sticky="w")

        ttk.Label(frm, text="循环次数：").grid(row=4, column=0, sticky="w", pady=2)
        self.var_loops = tk.IntVar(value=init.get("loops", 1))
        sp = ttk.Spinbox(frm, from_=1, to=9999, textvariable=self.var_loops, width=8)
        sp.grid(row=4, column=1, sticky="w", padx=5)
        ttk.Label(frm, text="(该片段重复几次)", foreground="gray").grid(
            row=4, column=2, columnspan=2, sticky="w")

        btns = ttk.Frame(frm)
        btns.grid(row=5, column=0, columnspan=4, pady=(12, 0), sticky="e")
        ttk.Button(btns, text="取消", command=self._cancel).pack(side="right", padx=5)
        ttk.Button(btns, text="确定", command=self._ok).pack(side="right")

        frm.columnconfigure(2, weight=1)
        self.bind("<Return>", lambda e: self._ok())
        self.bind("<Escape>", lambda e: self._cancel())

        if self.var_file.get():
            self._on_file_change()

        self.update_idletasks()
        x = parent.winfo_rootx() + (parent.winfo_width() - self.winfo_width()) // 2
        y = parent.winfo_rooty() + (parent.winfo_height() - self.winfo_height()) // 2
        self.geometry(f"+{max(0, x)}+{max(0, y)}")

    def _browse(self):
        f = filedialog.askopenfilename(
            parent=self,
            filetypes=[("音视频文件", "*.mp3 *.wav *.m4a *.aac *.flac *.ogg *.opus *.mp4 *.mkv *.webm *.mov *.avi"),
                       ("所有文件", "*.*")],
        )
        if f:
            self.var_file.set(f)

    def _on_file_change(self):
        f = self.var_file.get().strip()
        if not f or not Path(f).exists():
            self.src_dur = None
            self.var_duration.set("源文件时长: —")
            return
        dur = probe_duration(self.ffprobe_path, f)
        self.src_dur = dur
        if dur is None:
            self.var_duration.set("源文件时长: 未知（未安装 ffprobe 或文件不可读）")
        else:
            self.var_duration.set(f"源文件时长: {format_time(dur)}")

    def _ok(self):
        f = self.var_file.get().strip()
        if not f:
            messagebox.showwarning("提示", "请选择文件", parent=self)
            return
        if not Path(f).exists():
            messagebox.showwarning("提示", "文件不存在", parent=self)
            return
        try:
            start = parse_time(self.var_start.get())
            end = parse_time(self.var_end.get())
        except ValueError as e:
            messagebox.showwarning("时间格式错误", str(e), parent=self)
            return
        if start is not None and end is not None and end <= start:
            messagebox.showwarning("提示", "结束时间必须大于起始时间", parent=self)
            return
        try:
            loops = int(self.var_loops.get())
            if loops < 1:
                raise ValueError()
        except Exception:
            messagebox.showwarning("提示", "循环次数必须 ≥ 1", parent=self)
            return

        self.result = {
            "file": f,
            "start": self.var_start.get().strip(),
            "end": self.var_end.get().strip(),
            "start_sec": start,
            "end_sec": end,
            "loops": loops,
            "src_dur": self.src_dur,
        }
        self.destroy()

    def _cancel(self):
        self.result = None
        self.destroy()


# ──────────────────────────────────────────────────────────────
#  拼接 Tab
# ──────────────────────────────────────────────────────────────
class EditorTab(ttk.Frame):
    def __init__(self, parent, ffmpeg_path, ffprobe_path, logger):
        super().__init__(parent)
        self.ffmpeg_path = ffmpeg_path
        self.ffprobe_path = ffprobe_path
        self.logger = logger
        self.clips = []
        self.is_running = False
        self._build_ui()

    def _build_ui(self):
        frame_list = ttk.LabelFrame(self, text="片段列表（按顺序从上到下拼接）")
        frame_list.pack(fill="both", expand=True, padx=10, pady=(10, 5))

        cols = ("idx", "file", "start", "end", "seg", "loops", "sub")
        self.tree = ttk.Treeview(frame_list, columns=cols, show="headings", height=8)
        self.tree.heading("idx", text="#")
        self.tree.heading("file", text="文件")
        self.tree.heading("start", text="起始")
        self.tree.heading("end", text="结束")
        self.tree.heading("seg", text="单段时长")
        self.tree.heading("loops", text="循环 ×")
        self.tree.heading("sub", text="小计时长")
        self.tree.column("idx", width=34, anchor="center", stretch=False)
        self.tree.column("file", width=280, anchor="w")
        self.tree.column("start", width=75, anchor="center", stretch=False)
        self.tree.column("end", width=75, anchor="center", stretch=False)
        self.tree.column("seg", width=85, anchor="center", stretch=False)
        self.tree.column("loops", width=60, anchor="center", stretch=False)
        self.tree.column("sub", width=95, anchor="center", stretch=False)

        sb = ttk.Scrollbar(frame_list, command=self.tree.yview)
        self.tree.config(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.tree.pack(fill="both", expand=True, padx=5, pady=5)
        self.tree.bind("<Double-1>", lambda e: self._edit())

        frame_ops = ttk.Frame(self)
        frame_ops.pack(fill="x", padx=10, pady=5)
        ttk.Button(frame_ops, text="＋ 添加片段", command=self._add).pack(side="left")
        ttk.Button(frame_ops, text="编辑", command=self._edit).pack(side="left", padx=5)
        ttk.Button(frame_ops, text="删除", command=self._delete).pack(side="left")
        ttk.Button(frame_ops, text="↑ 上移", command=lambda: self._move(-1)).pack(side="left", padx=(15, 5))
        ttk.Button(frame_ops, text="↓ 下移", command=lambda: self._move(1)).pack(side="left")
        ttk.Button(frame_ops, text="清空", command=self._clear).pack(side="left", padx=(15, 0))

        self.var_total = tk.StringVar(value="总时长: 0:00.00   |   共 0 段 × 0 次")
        ttk.Label(self, textvariable=self.var_total,
                  font=("", 11, "bold")).pack(anchor="w", padx=12, pady=(6, 0))

        frame_out = ttk.LabelFrame(self, text="输出设置")
        frame_out.pack(fill="x", padx=10, pady=5)

        ttk.Label(frame_out, text="格式：").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.var_fmt = tk.StringVar(value="mp3")
        ttk.Combobox(frame_out, textvariable=self.var_fmt, values=AUDIO_FORMATS,
                     width=8, state="readonly").grid(row=0, column=1, padx=5)
        ttk.Label(frame_out, text="音质 (kbps)：").grid(row=0, column=2, sticky="w", padx=(15, 0))
        self.var_q = tk.StringVar(value="192")
        ttk.Combobox(frame_out, textvariable=self.var_q, values=QUALITY_OPTIONS,
                     width=8, state="readonly").grid(row=0, column=3, padx=5)

        ttk.Label(frame_out, text="输出文件：").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.var_outfile = tk.StringVar(value=str(Path.home() / "Downloads" / "merged.mp3"))
        ttk.Entry(frame_out, textvariable=self.var_outfile).grid(
            row=1, column=1, columnspan=3, sticky="we", padx=5)
        ttk.Button(frame_out, text="另存为...", command=self._save_as).grid(row=1, column=4, padx=5)
        frame_out.columnconfigure(3, weight=1)

        frame_btn = ttk.Frame(self)
        frame_btn.pack(fill="x", padx=10, pady=5)
        self.btn_go = ttk.Button(frame_btn, text="开始拼接", command=self._start)
        self.btn_go.pack(side="left")

        self.progress = ttk.Progressbar(self, mode="determinate")
        self.progress.pack(fill="x", padx=10, pady=5)

        self.var_status = tk.StringVar(value="就绪")
        ttk.Label(self, textvariable=self.var_status, anchor="w").pack(fill="x", padx=10, pady=(0, 5))

    # —— 片段数据增删改 ——
    def _add(self):
        dlg = ClipDialog(self, self.ffprobe_path)
        self.wait_window(dlg)
        if dlg.result:
            self.clips.append(dlg.result)
            self._refresh()

    def _edit(self):
        idx = self._selected_index()
        if idx is None:
            return
        dlg = ClipDialog(self, self.ffprobe_path, initial=self.clips[idx])
        self.wait_window(dlg)
        if dlg.result:
            self.clips[idx] = dlg.result
            self._refresh()

    def _delete(self):
        idx = self._selected_index()
        if idx is None:
            return
        del self.clips[idx]
        self._refresh()

    def _move(self, delta):
        idx = self._selected_index()
        if idx is None:
            return
        new = idx + delta
        if 0 <= new < len(self.clips):
            self.clips[idx], self.clips[new] = self.clips[new], self.clips[idx]
            self._refresh()
            items = self.tree.get_children()
            if new < len(items):
                self.tree.selection_set(items[new])

    def _clear(self):
        if not self.clips:
            return
        if messagebox.askyesno("确认", "清空所有片段？"):
            self.clips.clear()
            self._refresh()

    def _selected_index(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("提示", "请先在列表中选中一行")
            return None
        return self.tree.index(sel[0])

    def _clip_segment_seconds(self, clip):
        """单段（不算循环）的秒数；未知返回 None"""
        start = clip.get("start_sec")
        end = clip.get("end_sec")
        src = clip.get("src_dur")
        s = start if start is not None else 0.0
        if end is not None:
            e = end
        elif src is not None:
            e = src
        else:
            return None
        return max(0.0, e - s)

    def _refresh(self):
        self.tree.delete(*self.tree.get_children())
        total = 0.0
        total_unknown = False
        total_loops = 0
        for i, c in enumerate(self.clips, 1):
            seg = self._clip_segment_seconds(c)
            loops = c["loops"]
            if seg is None:
                sub = None
                total_unknown = True
            else:
                sub = seg * loops
                total += sub
            total_loops += loops
            self.tree.insert("", "end", values=(
                i,
                Path(c["file"]).name,
                c["start"] or "—",
                c["end"] or "—",
                format_time(seg) if seg is not None else "?",
                f"{loops}",
                format_time(sub) if sub is not None else "?",
            ))
        suffix = " (含未知)" if total_unknown else ""
        self.var_total.set(
            f"总时长: {format_time(total)}{suffix}   |   共 {len(self.clips)} 段 × 总计 {total_loops} 次循环"
        )

    # —— 输出 ——
    def _save_as(self):
        fmt = self.var_fmt.get()
        f = filedialog.asksaveasfilename(
            defaultextension=f".{fmt}",
            initialfile=f"merged.{fmt}",
            filetypes=[(f"{fmt.upper()} 音频", f"*.{fmt}"), ("所有文件", "*.*")],
        )
        if f:
            self.var_outfile.set(f)

    def _start(self):
        if self.is_running:
            return
        if not self.clips:
            messagebox.showwarning("提示", "请先添加至少一个片段")
            return
        if not self.ffmpeg_path:
            messagebox.showerror("错误", "未找到 ffmpeg。请把 ffmpeg.exe 放到本程序同目录。")
            return
        outfile = self.var_outfile.get().strip()
        if not outfile:
            messagebox.showwarning("提示", "请指定输出文件")
            return
        try:
            Path(outfile).parent.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            messagebox.showerror("错误", f"输出路径创建失败：{e}")
            return

        self.is_running = True
        self.btn_go.config(state="disabled")
        self.progress["value"] = 0
        self.progress["maximum"] = len(self.clips) + 1

        threading.Thread(
            target=self._worker,
            args=(list(self.clips), outfile, self.var_fmt.get(), self.var_q.get()),
            daemon=True,
        ).start()

    def _worker(self, clips, outfile, fmt, quality):
        tmpdir = tempfile.mkdtemp(prefix="vta_merge_")
        try:
            concat_lines = []
            for i, c in enumerate(clips):
                self.var_status.set(f"切片 {i+1}/{len(clips)}...")
                seg_path = Path(tmpdir) / f"seg_{i:04d}.wav"
                args = [self.ffmpeg_path, "-y", "-hide_banner", "-loglevel", "error"]
                if c.get("start_sec") is not None:
                    args += ["-ss", str(c["start_sec"])]
                if c.get("end_sec") is not None:
                    args += ["-to", str(c["end_sec"])]
                args += ["-i", c["file"], "-vn", "-ar", "44100", "-ac", "2", str(seg_path)]
                self.logger(f"[{i+1}/{len(clips)}] 切片: {Path(c['file']).name}  "
                            f"{c['start'] or '0'} → {c['end'] or '末尾'}  × {c['loops']}")
                r = subprocess.run(args, capture_output=True, text=True,
                                   creationflags=_NO_WINDOW)
                if r.returncode != 0:
                    raise RuntimeError(r.stderr.strip().splitlines()[-1]
                                       if r.stderr.strip() else "ffmpeg 切片失败")
                for _ in range(c["loops"]):
                    safe = seg_path.as_posix().replace("'", "'\\''")
                    concat_lines.append(f"file '{safe}'")
                self.progress["value"] = i + 1

            list_file = Path(tmpdir) / "list.txt"
            list_file.write_text("\n".join(concat_lines) + "\n", encoding="utf-8")

            self.var_status.set("合并并导出...")
            self.logger(f"合并 → {outfile}  ({fmt} @ {quality}k)")
            args = [self.ffmpeg_path, "-y", "-hide_banner", "-loglevel", "error",
                    "-f", "concat", "-safe", "0", "-i", str(list_file),
                    "-b:a", f"{quality}k", outfile]
            r = subprocess.run(args, capture_output=True, text=True,
                               creationflags=_NO_WINDOW)
            if r.returncode != 0:
                raise RuntimeError(r.stderr.strip().splitlines()[-1]
                                   if r.stderr.strip() else "ffmpeg 合并失败")

            self.progress["value"] = len(clips) + 1
            self.var_status.set(f"✓ 已导出：{outfile}")
            self.logger(f"✓ 完成：{outfile}")
        except Exception as e:
            self.logger(f"✗ 失败：{e}")
            self.var_status.set("失败")
        finally:
            shutil.rmtree(tmpdir, ignore_errors=True)
            self.is_running = False
            self.btn_go.config(state="normal")


# Windows 下创建子进程不弹黑窗
_NO_WINDOW = 0x08000000 if os.name == "nt" else 0


# ──────────────────────────────────────────────────────────────
#  主窗口
# ──────────────────────────────────────────────────────────────
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("视频音频提取器 v1.2 —— 下载 · 剪辑拼接")
        self.geometry("820x660")
        self.minsize(720, 580)

        self.log_queue = queue.Queue()
        self.ffmpeg_path = find_ffmpeg()
        self.ffprobe_path = find_ffprobe()

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=5, pady=(5, 0))

        self.tab_dl = DownloadTab(nb, self.ffmpeg_path, self._log)
        self.tab_ed = EditorTab(nb, self.ffmpeg_path, self.ffprobe_path, self._log)
        nb.add(self.tab_dl, text="  1. 下载提取  ")
        nb.add(self.tab_ed, text="  2. 剪辑拼接  ")

        frame_log = ttk.LabelFrame(self, text="日志")
        frame_log.pack(fill="both", expand=False, padx=5, pady=5)
        self.txt_log = tk.Text(frame_log, height=8, state="disabled", wrap="word")
        sb = ttk.Scrollbar(frame_log, command=self.txt_log.yview)
        self.txt_log.config(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.txt_log.pack(fill="both", expand=True, padx=5, pady=5)

        self.after(100, self._poll_log)
        self._banner()

    def _banner(self):
        if not self.ffmpeg_path:
            self._log("⚠ 未找到 ffmpeg — 无法转换格式/剪辑拼接。")
            self._log("  下载 https://www.gyan.dev/ffmpeg/builds/ffmpeg-release-essentials.zip")
            self._log("  解压后把 bin/ffmpeg.exe (和 ffprobe.exe) 放到本程序同目录。")
        else:
            self._log(f"✓ ffmpeg: {self.ffmpeg_path}")
            if not self.ffprobe_path:
                self._log("⚠ 找到 ffmpeg 但未找到 ffprobe —— 无法自动读取时长。")
                self._log("  把 ffprobe.exe 也放到同目录即可显示时长。")
            else:
                self._log(f"✓ ffprobe: {self.ffprobe_path}")
        self._log(f"✓ yt-dlp: {yt_dlp.version.__version__}\n")

    def _log(self, msg):
        self.log_queue.put(msg)

    def _poll_log(self):
        try:
            while True:
                msg = self.log_queue.get_nowait()
                self.txt_log.config(state="normal")
                self.txt_log.insert("end", msg + "\n")
                self.txt_log.see("end")
                self.txt_log.config(state="disabled")
        except queue.Empty:
            pass
        self.after(100, self._poll_log)


if __name__ == "__main__":
    App().mainloop()
