
updates.py
100%
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
================================================================================
🇻🇳 GCN-ME: TRÌNH CẬP NHẬT & ĐỒNG BỘ MÔI TRƯỜNG NHANH (updates.py)
Help the Vietnamese state (in matters related to land) to operate more quickly
and accurately.
Made by TonygamingTL
================================================================================
Chức năng (Lightweight & Smart Updater):
1. Tự động sửa lỗi lệnh 'python', 'pip', 'pyinstaller' is not recognized trên CMD.
2. Tự động đồng bộ file requirements.txt ở cả thư mục gốc và thư mục src/.
3. Chạy cập nhật thông minh (chỉ cài những gói còn thiếu trong requirements.txt,
   không tải lại từ đầu các gói đã đạt chuẩn).
"""

import os
import sys
import shutil
import subprocess
import threading
from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox

REQUIREMENTS_CONTENT = """pymupdf>=1.23.0
pdfplumber>=0.10.0
pandas>=2.1.0
openpyxl>=3.1.0
pytesseract>=0.3.10
Pillow>=10.0.0
gdown>=5.1.0
pyinstaller>=6.0.0
"""

CHECK_MODULES = [
    ("fitz", "pymupdf>=1.23.0"),
    ("pdfplumber", "pdfplumber>=0.10.0"),
    ("pandas", "pandas>=2.1.0"),
    ("openpyxl", "openpyxl>=3.1.0"),
    ("pytesseract", "pytesseract>=0.3.10"),
    ("PIL", "Pillow>=10.0.0"),
    ("gdown", "gdown>=5.1.0"),
    ("PyInstaller", "pyinstaller>=6.0.0"),
]


def get_project_root() -> str:
    cur = os.path.dirname(os.path.abspath(__file__))
    if os.path.basename(cur).lower() == "src":
        return os.path.dirname(cur)
    return cur


def sync_requirements_files() -> list[str]:
    """Tạo file requirements.txt ở cả thư mục gốc và src/ để lệnh pip install -r requirements.txt luôn chạy được."""
    root_dir = get_project_root()
    cur_dir = os.path.dirname(os.path.abspath(__file__))
    paths = []
    for d in {root_dir, cur_dir}:
        p = os.path.join(d, "requirements.txt")
        with open(p, "w", encoding="utf-8") as f:
            f.write(REQUIREMENTS_CONTENT)
        paths.append(p)
    return paths


def fix_windows_path() -> list[str]:
    """Thêm Python, Scripts và Tesseract vào User PATH trên Windows."""
    added = []
    if not sys.platform.startswith("win"):
        return added
    try:
        import winreg
        py_dir = os.path.dirname(sys.executable)
        scripts_dir = os.path.join(py_dir, "Scripts")
        tess_dir = r"C:\Program Files\Tesseract-OCR"
        candidates = [d for d in (py_dir, scripts_dir, tess_dir) if os.path.isdir(d)]

        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Environment", 0, winreg.KEY_ALL_ACCESS) as key:
            try:
                cur_path, _ = winreg.QueryValueEx(key, "Path")
            except FileNotFoundError:
                cur_path = ""
            parts = [p.strip() for p in cur_path.split(";") if p.strip()]
            lower_parts = [p.lower().rstrip("\\") for p in parts]
            for c in candidates:
                if c.lower().rstrip("\\") not in lower_parts:
                    parts.insert(0, c)
                    added.append(c)
            if added:
                winreg.SetValueEx(key, "Path", 0, winreg.REG_EXPAND_SZ, ";".join(parts))
        os.environ["PATH"] = ";".join(candidates) + ";" + os.environ.get("PATH", "")
    except Exception:
        pass
    return added


class SmartUpdaterGUI:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("🇻🇳 GCN-ME | Updates — Cập Nhật Nhanh & Sửa Lỗi Lệnh CMD (Made by TonygamingTL)")
        self.root.geometry("820x540")
        self.root.configure(bg="#F1F5F9")

        self.only_missing_var = tk.BooleanVar(value=True)
        self._build_ui()
        self._initial_check()

    def _build_ui(self):
        header = tk.Frame(self.root, bg="#1E3A8A", height=68)
        header.pack(fill="x")
        header.pack_propagate(False)

        tk.Label(
            header,
            text="🔄 GCN-ME SMART UPDATES: CẬP NHẬT NHANH & SỬA LỖI PIP / PATH",
            bg="#1E3A8A", fg="#FFFFFF", font=("Segoe UI", 13, "bold")
        ).pack(anchor="w", padx=18, pady=(12, 2))

        tk.Label(
            header,
            text="Khắc phục lỗi 'python / pip / pyinstaller is not recognized' & Chỉ cài gói còn thiếu trong requirements.txt",
            bg="#1E3A8A", fg="#BFDBFE", font=("Segoe UI", 9, "italic")
        ).pack(anchor="w", padx=18)

        body = tk.Frame(self.root, bg="#F1F5F9", padx=14, pady=12)
        body.pack(fill="both", expand=True)

        opt_frame = tk.LabelFrame(
            body, text=" Tùy chọn chế độ cập nhật ", bg="#FFFFFF", fg="#1E3A8A",
            font=("Segoe UI", 10, "bold"), padx=12, pady=10
        )
        opt_frame.pack(fill="x", pady=(0, 10))

        tk.Checkbutton(
            opt_frame,
            text="⚡ Chế độ thông minh: Chỉ cài đặt những thư viện còn thiếu (Không tải full lại các gói đã có)",
            variable=self.only_missing_var, bg="#FFFFFF", font=("Segoe UI", 10, "bold")
        ).pack(anchor="w")

        btn_row = tk.Frame(body, bg="#F1F5F9")
        btn_row.pack(fill="x", pady=(0, 10))

        self.btn_update = tk.Button(
            btn_row,
            text="🚀 SỬA LỖI PATH & CẬP NHẬT REQUIREMENTS.TXT NGAY",
            bg="#2563EB", fg="#FFFFFF", font=("Segoe UI", 10, "bold"),
            relief="flat", padx=14, pady=8, cursor="hand2",
            command=self.run_smart_update
        )
        self.btn_update.pack(side="left", fill="x", expand=True, padx=(0, 8))

        tk.Button(
            btn_row,
            text="⬇️ Mở trình tải OCR (downloads.py)",
            bg="#059669", fg="#FFFFFF", font=("Segoe UI", 10, "bold"),
            relief="flat", padx=12, pady=8, cursor="hand2",
            command=self.open_downloads_tool
        ).pack(side="left", padx=(0, 8))

        tk.Button(
            btn_row,
            text="🇻🇳 Mở GCN_Extractor",
            bg="#0F172A", fg="#FFFFFF", font=("Segoe UI", 10, "bold"),
            relief="flat", padx=12, pady=8, cursor="hand2",
            command=self.open_extractor
        ).pack(side="right")

        self.txt_log = tk.Text(body, bg="#0F172A", fg="#E2E8F0", font=("Consolas", 10), padx=10, pady=8, state="disabled")
        self.txt_log.pack(fill="both", expand=True)

    def log(self, msg: str):
        ts = datetime.now().strftime("%H:%M:%S")
        self.txt_log.config(state="normal")
        self.txt_log.insert("end", f"[{ts}] {msg}\n")
        self.txt_log.see("end")
        self.txt_log.config(state="disabled")

    def _initial_check(self):
        files = sync_requirements_files()
        self.log(f"✔ Đã đồng bộ file requirements.txt tại: {', '.join(files)}")
        missing = [spec for mod, spec in CHECK_MODULES if not self._can_import(mod)]
        if not missing:
            self.log("✅ Tất cả 8/8 thư viện trong requirements.txt đều đã được cài đặt đầy đủ!")
        else:
            self.log(f"⚠️ Phát hiện {len(missing)} thư viện cần cài đặt thêm: {', '.join(missing)}")

    @staticmethod
    def _can_import(mod_name: str) -> bool:
        try:
            __import__(mod_name)
            return True
        except ImportError:
            return False

    def run_smart_update(self):
        self.btn_update.config(state="disabled")
        threading.Thread(target=self._worker_update, daemon=True).start()

    def _worker_update(self):
        try:
            # 1. Sửa PATH
            added = fix_windows_path()
            if added:
                self.root.after(0, lambda a=added: self.log(f"✔ Đã thêm vào Windows PATH: {', '.join(a)}"))
            else:
                self.root.after(0, lambda: self.log("✔ Đường dẫn Python & Scripts đã có sẵn trong Windows PATH."))

            # 2. Đồng bộ requirements.txt
            req_files = sync_requirements_files()

            # 3. Kiểm tra chỉ cài gói thiếu hoặc cài từ requirements.txt
            if self.only_missing_var.get():
                missing = [spec for mod, spec in CHECK_MODULES if not self._can_import(mod)]
                if not missing:
                    self.root.after(0, lambda: self.log("✅ Không có thư viện nào bị thiếu! Bỏ qua bước tải lại để tiết kiệm thời gian."))
                    self.root.after(0, lambda: messagebox.showinfo("Đã đầy đủ", "Mọi thư viện trong requirements.txt đều đã có sẵn, không cần tải lại!"))
                    return
                cmd = [sys.executable, "-m", "pip", "install"] + missing
                self.root.after(0, lambda m=missing: self.log(f"📦 Đang cài đặt đúng {len(m)} gói còn thiếu: {', '.join(m)}"))
            else:
                cmd = [sys.executable, "-m", "pip", "install", "-r", req_files[0]]
                self.root.after(0, lambda: self.log(f"📦 Đang chạy: py -m pip install -r \"{req_files[0]}\""))

            proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, encoding="utf-8", errors="replace")
            for line in proc.stdout:
                if line.strip():
                    self.root.after(0, lambda l=line.strip(): self.log(f"   {l}"))
            proc.wait()

            if proc.returncode == 0:
                self.root.after(0, lambda: self.log("🎉 Cập nhật hoàn tất!"))
                self.root.after(0, lambda: messagebox.showinfo("Thành công", "Đã cập nhật xong các thư viện cần thiết!"))
            else:
                self.root.after(0, lambda c=proc.returncode: self.log(f"❌ Lệnh pip trả về mã lỗi {c}"))
        except Exception as e:
            self.root.after(0, lambda err=e: self.log(f"❌ Lỗi: {err}"))
        finally:
            self.root.after(0, lambda: self.btn_update.config(state="normal"))

    def open_downloads_tool(self):
        try:
            import downloads
            downloads.open_downloader_window(self.root)
        except Exception as e:
            messagebox.showerror("Lỗi mở downloads", f"Không thể mở trình tải về:\n{e}")

    def open_extractor(self):
        try:
            import GCN_Extractor
            win = tk.Toplevel(self.root)
            GCN_Extractor.GCNExtractorGUI(win)
        except Exception as e:
            messagebox.showerror("Lỗi mở GCN_Extractor", f"Không thể mở GCN_Extractor:\n{e}")


def open_updater_window(parent=None):
    """Mở cửa sổ SmartUpdaterGUI trực tiếp trong cùng tiến trình (.exe)."""
    win = tk.Toplevel(parent) if parent is not None else tk.Tk()
    app = SmartUpdaterGUI(win)
    if parent is None:
        win.mainloop()
    return app


if __name__ == "__main__":
    root = tk.Tk()
    app = SmartUpdaterGUI(root)
    root.mainloop()
Đang hiển thị updates.py.
