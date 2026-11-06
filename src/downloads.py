#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
================================================================================
🇻🇳 GCN-ME: TRÌNH TẢI THÔNG MINH - CHỈ TẢI MỤC CẦN TẢI (downloads.py)
Help the Vietnamese state (in matters related to land) to operate more quickly
and accurately.
Made by TonygamingTL
================================================================================
Cơ chế thông minh (Smart Selective Downloader):
- Tự động kiểm tra máy tính đã có sẵn thành phần nào.
- KHÔNG tải full tràn lan gây tốn dung lượng và thời gian.
- CHỈ TẢI đúng những mục đang thiếu (cần tải) hoặc mục người dùng tích chọn:
  1. Python 3.10+ (Chỉ tải nếu máy chưa có hoặc phiên bản < 3.10)
  2. Tesseract OCR (Chỉ tải nếu máy chưa cài đặt tesseract.exe)
  3. Gói ngôn ngữ Tiếng Việt 'vie.traineddata' (Chỉ tải nếu chưa có file vie)
  4. Tệp chuẩn requirements.txt (Chỉ tạo/tải các gói thư viện còn thiếu)
"""

import os
import sys
import shutil
import subprocess
import threading
import urllib.request
import ssl
from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox

# ==============================================================================
# CẤU HÌNH LINK TẢI CHÍNH THỨC & DANH SÁCH KIỂM TRA
# ==============================================================================
PYTHON_INSTALLER_URL = "https://www.python.org/ftp/python/3.11.9/python-3.11.9-amd64.exe"
TESSERACT_INSTALLER_URL = (
    "https://github.com/tesseract-ocr/tesseract/releases/download/5.5.0/"
    "tesseract-ocr-w64-setup-5.5.0.20241111.exe"
)
VIE_LANG_URL = "https://github.com/tesseract-ocr/tessdata_fast/raw/main/vie.traineddata"
ENG_LANG_URL = "https://github.com/tesseract-ocr/tessdata_fast/raw/main/eng.traineddata"

REQUIREMENTS_TEXT = """pymupdf>=1.23.0
pdfplumber>=0.10.0
pandas>=2.1.0
openpyxl>=3.1.0
pytesseract>=0.3.10
Pillow>=10.0.0
gdown>=5.1.0
pyinstaller>=6.0.0
"""

REQUIRED_LIBS = [
    ("fitz", "pymupdf>=1.23.0", "pymupdf"),
    ("pdfplumber", "pdfplumber>=0.10.0", "pdfplumber"),
    ("pandas", "pandas>=2.1.0", "pandas"),
    ("openpyxl", "openpyxl>=3.1.0", "openpyxl"),
    ("pytesseract", "pytesseract>=0.3.10", "pytesseract"),
    ("PIL", "Pillow>=10.0.0", "Pillow"),
    ("gdown", "gdown>=5.1.0", "gdown"),
    ("PyInstaller", "pyinstaller>=6.0.0", "pyinstaller"),
]


def get_project_root() -> str:
    cur = os.path.dirname(os.path.abspath(__file__))
    if os.path.basename(cur).lower() == "src":
        return os.path.dirname(cur)
    return cur


def get_tessdata_dir() -> str:
    d = os.path.join(get_project_root(), "tessdata")
    os.makedirs(d, exist_ok=True)
    return d


def is_python_310_or_higher() -> tuple[bool, str]:
    v = sys.version_info
    ver_str = f"{v.major}.{v.minor}.{v.micro}"
    return (v.major > 3 or (v.major == 3 and v.minor >= 10)), ver_str


def find_tesseract_binary() -> str:
    paths = [
        shutil.which("tesseract") or "",
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
        os.path.expandvars(r"%LOCALAPPDATA%\Programs\Tesseract-OCR\tesseract.exe"),
        os.path.expandvars(r"%LOCALAPPDATA%\Tesseract-OCR\tesseract.exe"),
        "/usr/bin/tesseract",
    ]
    for p in paths:
        if p and os.path.isfile(p):
            return p
    return ""


def find_vie_traineddata() -> tuple[bool, str]:
    dirs = [get_tessdata_dir()]
    tess_bin = find_tesseract_binary()
    if tess_bin:
        dirs.insert(0, os.path.join(os.path.dirname(tess_bin), "tessdata"))
    for d in dirs:
        f = os.path.join(d, "vie.traineddata")
        if os.path.isfile(f) and os.path.getsize(f) > 100_000:
            return True, f
    return False, ""


def get_missing_pip_packages() -> tuple[list[str], list[str]]:
    """Trả về: (danh_sách_gói_cần_tải, danh_sách_gói_đã_có)"""
    missing = []
    installed = []
    for mod_name, spec, short_name in REQUIRED_LIBS:
        try:
            __import__(mod_name)
            installed.append(short_name)
        except ImportError:
            missing.append(spec)
    return missing, installed


def download_with_progress(url: str, dest_path: str, progress_cb=None):
    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0 GCN-ME-Downloader"})
    with urllib.request.urlopen(req, context=ctx, timeout=60) as resp:
        total = int(resp.info().get("Content-Length", 0).strip() or 0)
        downloaded = 0
        with open(dest_path, "wb") as out:
            while True:
                chunk = resp.read(65536)
                if not chunk:
                    break
                out.write(chunk)
                downloaded += len(chunk)
                if progress_cb and total > 0:
                    progress_cb(min(100, int(downloaded * 100 / total)), downloaded, total)


# ==============================================================================
# GIAO DIỆN CHỌN LỌC NHỮNG MỤC CẦN TẢI (KHÔNG TẢI DƯ THỪA)
# ==============================================================================
class SmartDownloaderGUI:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("🇻🇳 GCN-ME | Downloads — Chỉ Tải Những Mục Cần Tải (Made by TonygamingTL)")
        self.root.geometry("880x640")
        self.root.minsize(820, 580)
        self.root.configure(bg="#F1F5F9")

        # Biến checkbox cho phép tự động bật những mục CHƯA CÓ và tắt những mục ĐÃ CÓ
        self.var_dl_python = tk.BooleanVar(value=False)
        self.var_dl_reqs = tk.BooleanVar(value=False)
        self.var_dl_tesseract = tk.BooleanVar(value=False)
        self.var_dl_vie = tk.BooleanVar(value=False)

        self.is_downloading = False
        self._setup_styles()
        self._build_ui()
        self.scan_needed_downloads()

    def _setup_styles(self):
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Main.TFrame", background="#F1F5F9")
        style.configure("Card.TLabelframe", background="#FFFFFF", foreground="#1E3A8A", font=("Segoe UI", 10, "bold"))
        style.configure("Card.TLabelframe.Label", background="#FFFFFF", foreground="#1E3A8A", font=("Segoe UI", 10, "bold"))
        style.configure("TLabel", background="#FFFFFF", foreground="#1E293B", font=("Segoe UI", 10))
        style.configure("TCheckbutton", background="#FFFFFF", foreground="#1E293B", font=("Segoe UI", 10, "bold"))
        style.configure("Primary.TButton", font=("Segoe UI", 11, "bold"), background="#2563EB", foreground="#FFFFFF", padding=(14, 8))
        style.map("Primary.TButton", background=[("active", "#1D4ED8"), ("disabled", "#94A3B8")])
        style.configure("Secondary.TButton", font=("Segoe UI", 9, "bold"), background="#E2E8F0", foreground="#1E293B", padding=(10, 6))

    def _build_ui(self):
        header = tk.Frame(self.root, bg="#1E3A8A", height=70)
        header.pack(fill="x")
        header.pack_propagate(False)

        tk.Label(
            header,
            text="⬇️ GCN-ME SMART DOWNLOADS: CHỈ TẢI NHỮNG THỨ CẦN TẢI (CÒN THIẾU)",
            bg="#1E3A8A", fg="#FFFFFF", font=("Segoe UI", 13, "bold")
        ).pack(anchor="w", padx=18, pady=(12, 2))

        tk.Label(
            header,
            text="Tự động bỏ qua các phần đã cài sẵn trên máy — Không tải full lại từ đầu gây tốn thời gian!",
            bg="#1E3A8A", fg="#BFDBFE", font=("Segoe UI", 9, "italic")
        ).pack(anchor="w", padx=18)

        main = ttk.Frame(self.root, style="Main.TFrame", padding=14)
        main.pack(fill="both", expand=True)

        card = ttk.LabelFrame(
            main,
            text=" Danh Sách Kiểm Tra Tự Động (Mục nào CHƯA CÓ mới tự động tích chọn để tải) ",
            style="Card.TLabelframe", padding=14
        )
        card.pack(fill="x", pady=(0, 10))
        card.columnconfigure(1, weight=1)

        # Mục 1: Python 3.10+
        ttk.Checkbutton(card, text="1. Bộ cài Python 3.10+:", variable=self.var_dl_python).grid(row=0, column=0, sticky="w", pady=6)
        self.lbl_py = ttk.Label(card, text="Đang kiểm tra...")
        self.lbl_py.grid(row=0, column=1, sticky="w", padx=10, pady=6)

        # Mục 2: Các thư viện còn thiếu trong requirements.txt
        ttk.Checkbutton(card, text="2. Thư viện thiếu (requirements.txt):", variable=self.var_dl_reqs).grid(row=1, column=0, sticky="w", pady=6)
        self.lbl_reqs = ttk.Label(card, text="Đang kiểm tra...")
        self.lbl_reqs.grid(row=1, column=1, sticky="w", padx=10, pady=6)

        # Mục 3: Tesseract OCR
        ttk.Checkbutton(card, text="3. Bộ cài Tesseract OCR:", variable=self.var_dl_tesseract).grid(row=2, column=0, sticky="w", pady=6)
        self.lbl_tess = ttk.Label(card, text="Đang kiểm tra...")
        self.lbl_tess.grid(row=2, column=1, sticky="w", padx=10, pady=6)

        # Mục 4: Gói Tiếng Việt vie.traineddata
        ttk.Checkbutton(card, text="4. Gói Tiếng Việt (vie.traineddata):", variable=self.var_dl_vie).grid(row=3, column=0, sticky="w", pady=6)
        self.lbl_vie = ttk.Label(card, text="Đang kiểm tra...")
        self.lbl_vie.grid(row=3, column=1, sticky="w", padx=10, pady=6)

        # Thanh nút bấm
        btn_bar = ttk.Frame(main, style="Main.TFrame")
        btn_bar.pack(fill="x", pady=(0, 10))

        self.btn_smart_dl = ttk.Button(
            btn_bar,
            text="🎯 CHỈ TẢI CÁC MỤC CẦN TẢI (ĐÃ TÍCH CHỌN Ở TRÊN)",
            style="Primary.TButton",
            command=self.start_selective_download
        )
        self.btn_smart_dl.pack(side="left", fill="x", expand=True, padx=(0, 8))

        ttk.Button(
            btn_bar,
            text="🔄 Quét kiểm tra lại",
            style="Secondary.TButton",
            command=self.scan_needed_downloads
        ).pack(side="right")

        # Khung nhật ký & Tiến trình
        log_card = ttk.LabelFrame(main, text=" Tiến Độ Tải Về (Chỉ tải mục cần thiết) ", style="Card.TLabelframe", padding=10)
        log_card.pack(fill="both", expand=True)

        self.lbl_prog = ttk.Label(log_card, text="Sẵn sàng.", font=("Segoe UI", 9, "bold"))
        self.lbl_prog.pack(anchor="w", pady=(0, 4))

        self.progress = ttk.Progressbar(log_card, orient="horizontal", mode="determinate")
        self.progress.pack(fill="x", pady=(0, 8))

        self.txt_log = tk.Text(log_card, bg="#0F172A", fg="#E2E8F0", font=("Consolas", 10), padx=10, pady=8, state="disabled")
        self.txt_log.pack(fill="both", expand=True)

    def log(self, msg: str):
        ts = datetime.now().strftime("%H:%M:%S")
        self.txt_log.config(state="normal")
        self.txt_log.insert("end", f"[{ts}] {msg}\n")
        self.txt_log.see("end")
        self.txt_log.config(state="disabled")

    def scan_needed_downloads(self):
        """Quét hệ thống: Mục nào đã có thì BỎ TÍCH (không tải lại), mục nào thiếu mới TÍCH CHỌN."""
        # 1. Python
        py_ok, py_ver = is_python_310_or_higher()
        if py_ok:
            self.var_dl_python.set(False)
            self.lbl_py.config(text=f"✅ Đã có Python {py_ver} (Bỏ qua - Không cần tải)", foreground="#059669")
        else:
            self.var_dl_python.set(True)
            self.lbl_py.config(text=f"⚠️ Máy đang dùng Python {py_ver} -> CẦN TẢI bản 3.10+", foreground="#DC2626")

        # 2. Thư viện trong requirements.txt
        missing_pkgs, installed_pkgs = get_missing_pip_packages()
        if not missing_pkgs:
            self.var_dl_reqs.set(False)
            self.lbl_reqs.config(text=f"✅ Đã có đủ {len(installed_pkgs)}/{len(REQUIRED_LIBS)} thư viện (Bỏ qua - Không cần tải)", foreground="#059669")
        else:
            self.var_dl_reqs.set(True)
            short_missing = [m.split(">=")[0] for m in missing_pkgs]
            self.lbl_reqs.config(
                text=f"⚠️ Đã có {len(installed_pkgs)} gói, CẦN TẢI thêm {len(missing_pkgs)} gói: {', '.join(short_missing)}",
                foreground="#DC2626"
            )

        # 3. Tesseract OCR
        tess_bin = find_tesseract_binary()
        if tess_bin:
            self.var_dl_tesseract.set(False)
            self.lbl_tess.config(text=f"✅ Đã có tại {tess_bin} (Bỏ qua - Không cần tải)", foreground="#059669")
        else:
            self.var_dl_tesseract.set(True)
            self.lbl_tess.config(text="⚠️ Chưa có trên máy -> CẦN TẢI Tesseract OCR", foreground="#DC2626")

        # 4. Gói ngôn ngữ vie.traineddata
        vie_ok, vie_path = find_vie_traineddata()
        if vie_ok:
            self.var_dl_vie.set(False)
            self.lbl_vie.config(text=f"✅ Đã có tại {vie_path} (Bỏ qua - Không cần tải)", foreground="#059669")
        else:
            self.var_dl_vie.set(True)
            self.lbl_vie.config(text="⚠️ Chưa có gói Tiếng Việt -> CẦN TẢI vie.traineddata (~4 MB)", foreground="#DC2626")

        self.log("🔍 Đã quét xong! Những mục ĐÃ CÓ được tự động bỏ qua để không phải tải full lại.")

    def start_selective_download(self):
        if self.is_downloading:
            return
        if not any([self.var_dl_python.get(), self.var_dl_reqs.get(), self.var_dl_tesseract.get(), self.var_dl_vie.get()]):
            messagebox.showinfo("Không cần tải thêm", "Tất cả các mục cần thiết đều đã có sẵn trên máy của bạn!")
            return

        self.is_downloading = True
        self.btn_smart_dl.config(state="disabled")
        threading.Thread(target=self._worker_selective_download, daemon=True).start()

    def _worker_selective_download(self):
        try:
            root_dir = get_project_root()

            # Luôn đảm bảo file requirements.txt tồn tại nhẹ nhàng (vài byte)
            for d in {root_dir, os.path.dirname(os.path.abspath(__file__))}:
                with open(os.path.join(d, "requirements.txt"), "w", encoding="utf-8") as f:
                    f.write(REQUIREMENTS_TEXT)

            # 1. Chỉ tải Python 3.10+ nếu được chọn
            if self.var_dl_python.get() and sys.platform.startswith("win"):
                dest_py = os.path.join(root_dir, "python-3.11.9-amd64.exe")
                self.root.after(0, lambda: self.log("⬇️ Đang tải bộ cài Python 3.11.9 (vì máy chưa có bản 3.10+)..."))
                download_with_progress(
                    PYTHON_INSTALLER_URL, dest_py,
                    lambda pct, c, t: self.root.after(0, lambda p=pct: self._set_prog(p, f"Đang tải Python 3.11.9... ({p}%)"))
                )
                self.root.after(0, lambda: self.log("🚀 Đang cài đặt Python 3.11.9 kèm PATH..."))
                subprocess.call([dest_py, "/passive", "PrependPath=1", "Include_pip=1", "Include_tcltk=1"])

            # 2. Chỉ cài đúng các thư viện còn thiếu trong requirements.txt (không tải lại gói đã có)
            if self.var_dl_reqs.get():
                missing_pkgs, _ = get_missing_pip_packages()
                pkgs_to_install = missing_pkgs if missing_pkgs else [spec for _, spec, _ in REQUIRED_LIBS]
                self.root.after(0, lambda: self.log(f"📦 Chỉ tải và cài đặt các gói còn thiếu: {', '.join(pkgs_to_install)}"))
                self.root.after(0, lambda: self._set_prog(35, "Đang cài đặt các thư viện còn thiếu qua pip..."))
                subprocess.check_call([sys.executable, "-m", "pip", "install"] + pkgs_to_install)
                self.root.after(0, lambda: self.log("✔ Đã cài xong các thư viện còn thiếu!"))

            # 3. Chỉ tải Tesseract OCR nếu máy chưa có
            if self.var_dl_tesseract.get() and sys.platform.startswith("win"):
                if shutil.which("winget"):
                    self.root.after(0, lambda: self.log("⏳ Đang cài Tesseract OCR qua winget..."))
                    subprocess.call([
                        "winget", "install", "--id", "UB-Mannheim.TesseractOCR",
                        "-e", "--silent", "--accept-package-agreements", "--accept-source-agreements"
                    ])
                if not find_tesseract_binary():
                    dest_tess = os.path.join(root_dir, "tesseract-ocr-setup.exe")
                    self.root.after(0, lambda: self.log("⬇️ Đang tải bộ cài Tesseract OCR..."))
                    download_with_progress(
                        TESSERACT_INSTALLER_URL, dest_tess,
                        lambda pct, c, t: self.root.after(0, lambda p=pct: self._set_prog(p, f"Đang tải Tesseract OCR... ({p}%)"))
                    )
                    self.root.after(0, lambda: self.log("🚀 Đang mở trình cài đặt Tesseract OCR..."))
                    subprocess.call([dest_tess])

            # 4. Chỉ tải gói ngôn ngữ Tiếng Việt vie.traineddata (~4MB) nếu chưa có
            if self.var_dl_vie.get():
                tessdata_dir = get_tessdata_dir()
                vie_dest = os.path.join(tessdata_dir, "vie.traineddata")
                self.root.after(0, lambda: self.log("🇻🇳 Đang tải gói ngôn ngữ Tiếng Việt (vie.traineddata - chỉ ~4MB)..."))
                download_with_progress(
                    VIE_LANG_URL, vie_dest,
                    lambda pct, c, t: self.root.after(0, lambda p=pct: self._set_prog(p, f"Đang tải vie.traineddata... ({p}%)"))
                )
                os.environ["TESSDATA_PREFIX"] = tessdata_dir
                self.root.after(0, lambda: self.log(f"✔ Đã lưu gói Tiếng Việt tại: {vie_dest}"))

                # Sao chép vào thư mục Tesseract hệ thống nếu có
                tess_bin = find_tesseract_binary()
                if tess_bin:
                    sys_td = os.path.join(os.path.dirname(tess_bin), "tessdata")
                    if os.path.isdir(sys_td):
                        try:
                            shutil.copy2(vie_dest, os.path.join(sys_td, "vie.traineddata"))
                        except Exception:
                            pass

            self.root.after(0, lambda: self._set_prog(100, "✅ Đã tải xong đúng các mục cần tải!"))
            self.root.after(0, lambda: messagebox.showinfo("Hoàn tất", "Đã tải và cài đặt xong các mục còn thiếu!"))
        except Exception as e:
            self.root.after(0, lambda err=e: self.log(f"❌ Lỗi: {err}"))
            self.root.after(0, lambda err=e: messagebox.showerror("Lỗi tải về", str(err)))
        finally:
            self.is_downloading = False
            self.root.after(0, lambda: self.btn_smart_dl.config(state="normal"))
            self.root.after(0, self.scan_needed_downloads)

    def _set_prog(self, pct: int, text: str):
        self.progress["value"] = pct
        self.lbl_prog.config(text=text)


if __name__ == "__main__":
    root = tk.Tk()
    app = SmartDownloaderGUI(root)
    root.mainloop()
