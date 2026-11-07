#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
================================================================================
🇻🇳 GCN-ME (GCN_Extractor.py)
Help the Vietnamese state (in matters related to land) to operate more quickly
and accurately.
Made by TonygamingTL
================================================================================
Công cụ tự động trích xuất và chuẩn hóa dữ liệu Giấy chứng nhận quyền sử dụng đất
(Sổ đỏ / Sổ hồng) từ tệp PDF (Văn bản số & Scan OCR) sang bảng tính Excel (.xlsx).
Giao diện trực quan từng bước (Wizard UI) dành cho người mới bắt đầu.
"""

import os
import sys
import re
import glob
import shutil
import subprocess
import threading
from datetime import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# Liên kết trực tiếp 2 module downloads.py và updates.py trong src/ để PyInstaller tự động gom chung vào 1 file .exe
try:
    import downloads
except ImportError:
    downloads = None

try:
    import updates
except ImportError:
    updates = None

# ==============================================================================
# 1. KIỂM TRA CÁC THƯ VIỆN PHỤ THUỘC (HỖ TRỢ NGƯỜI MỚI CÀI ĐẶT TỰ ĐỘNG)
# ==============================================================================
OPTIONAL_LIBS = {
    "fitz": "pymupdf",
    "pdfplumber": "pdfplumber",
    "openpyxl": "openpyxl",
    "pytesseract": "pytesseract",
    "PIL": "Pillow",
    "gdown": "gdown",
}

INSTALLED_STATUS = {}
for mod_name, pip_name in OPTIONAL_LIBS.items():
    try:
        __import__(mod_name)
        INSTALLED_STATUS[mod_name] = True
    except ImportError:
        INSTALLED_STATUS[mod_name] = False


# ==============================================================================
# 2. BỘ PHÂN TÍCH & BÓC TÁCH DỮ LIỆU ĐỊA CHÍNH VIỆT NAM (REGEX & PARSER)
# ==============================================================================
LAND_CODE_MAP = [
    ("đất ở tại đô thị", "ODT"),
    ("đất ở tại nông thôn", "ONT"),
    ("đất chuyên trồng lúa nước", "LUC"),
    ("đất trồng lúa nước còn lại", "LUK"),
    ("đất trồng lúa nước", "LUA"),
    ("đất trồng lúa", "LUA"),
    ("đất trồng cây lâu năm", "CLN"),
    ("đất trồng cây hàng năm khác", "BHK"),
    ("đất trồng cây hàng năm", "CHN"),
    ("đất nuôi trồng thủy sản", "NTS"),
    ("đất rừng sản xuất", "RSX"),
    ("đất rừng phòng hộ", "RPH"),
    ("đất thương mại, dịch vụ", "TMD"),
    ("đất cơ sở sản xuất phi nông nghiệp", "SKC"),
]


class GCNParser:
    """Bộ bóc tách thông tin từ văn bản Giấy chứng nhận QSDĐ (Sổ đỏ / Sổ hồng)."""

    @staticmethod
    def clean_text(text: str) -> str:
        if not text:
            return ""
        text = re.sub(r"[ \t]+", " ", text)
        text = re.sub(r"\n{3,}", "\n\n", text)
        return text.strip()

    @staticmethod
    def detect_land_code(purpose_text: str) -> str:
        if not purpose_text:
            return ""
        lower_p = purpose_text.lower()
        codes = []
        for phrase, code in LAND_CODE_MAP:
            if phrase in lower_p and code not in codes:
                codes.append(code)
        explicit = re.findall(r"\b(ODT|ONT|LUC|LUK|LUA|CLN|BHK|NTS|RSX|RPH|TMD|SKC)\b", purpose_text.upper())
        for c in explicit:
            if c not in codes:
                codes.append(c)
        return ", ".join(codes) if codes else "Khác"

    @classmethod
    def extract_text_from_pdf(cls, pdf_path: str, use_ocr: bool = True, tesseract_cmd: str = "") -> tuple[str, str]:
        """
        Đọc văn bản từ file PDF.
        Trả về: (nội_dung_văn_bản, phương_thức_đọc)
        """
        text_content = ""
        method_used = "PDF Điện tử"

        # Ưu tiên 1: Dùng PyMuPDF (fitz) tốc độ cao
        if INSTALLED_STATUS.get("fitz"):
            import fitz
            doc = fitz.open(pdf_path)
            pages_text = []
            for page in doc:
                pages_text.append(page.get_text("text"))
            text_content = "\n".join(pages_text).strip()

            # Nếu file PDF là bản scan (ít hơn 40 ký tự chữ), tự động chuyển sang OCR bằng PyMuPDF Pixmap
            if len(text_content) < 40 and use_ocr and INSTALLED_STATUS.get("pytesseract") and INSTALLED_STATUS.get("PIL"):
                import pytesseract
                from PIL import Image
                import io

                if tesseract_cmd and os.path.exists(tesseract_cmd):
                    pytesseract.pytesseract.tesseract_cmd = tesseract_cmd

                ocr_pages = []
                for page in doc:
                    pix = page.get_pixmap(dpi=250)
                    img = Image.open(io.BytesIO(pix.tobytes("png")))
                    try:
                        page_ocr = pytesseract.image_to_string(img, lang="vie")
                    except Exception:
                        page_ocr = pytesseract.image_to_string(img)
                    ocr_pages.append(page_ocr)
                text_content = "\n".join(ocr_pages).strip()
                method_used = "OCR (Bản Scan)"
            doc.close()

        # Ưu tiên 2: Dùng pdfplumber nếu chưa cài PyMuPDF
        elif INSTALLED_STATUS.get("pdfplumber"):
            import pdfplumber
            with pdfplumber.open(pdf_path) as pdf:
                pages_text = [p.extract_text() or "" for p in pdf.pages]
                text_content = "\n".join(pages_text).strip()

        return cls.clean_text(text_content), method_used

    @classmethod
    def parse_gcn_fields(cls, raw_text: str, file_name: str, read_method: str, split_owners: bool = True) -> list[dict]:
        """
        Bóc tách các trường chuẩn địa chính từ văn bản GCN.
        Hỗ trợ tách dòng nếu có nhiều đồng chủ sử dụng (Ông/Bà).
        """
        if not raw_text:
            return [{
                "ten_file": file_name,
                "so_phat_hanh": "",
                "so_vao_so": "",
                "chu_su_dung": "[Không đọc được chữ - Hãy bật OCR]",
                "nam_sinh": "",
                "so_giay_to": "",
                "dia_chi_thuong_tru": "",
                "so_thua": "",
                "so_to_ban_do": "",
                "dia_chi_thua_dat": "",
                "dien_tich": "",
                "hinh_thuc_su_dung": "",
                "muc_dich_su_dung": "",
                "ma_loai_dat": "",
                "thoi_han_su_dung": "",
                "nguon_goc_su_dung": "",
                "bien_dong_muc_iv": "",
                "phuong_thuc_doc": read_method,
                "trang_thai": "Cần kiểm tra",
            }]

        # 1. Số phát hành (Số seri trên trang 1, VD: DI 123456, ĐG 654321)
        so_phat_hanh = ""
        m_seri = re.search(r"\b([A-ZĐ]{1,2}\s?\d{6,8})\b", raw_text)
        if m_seri:
            so_phat_hanh = re.sub(r"\s+", " ", m_seri.group(1)).strip()

        # 2. Số vào sổ cấp GCN (VD: CH 01234, CS 05678, CT00123)
        so_vao_so = ""
        m_vaoso = re.search(r"(?:Số vào sổ cấp GCN|vào sổ cấp giấy chứng nhận)\s*[:\.]?\s*([^\n]+)", raw_text, re.IGNORECASE)
        if m_vaoso:
            so_vao_so = m_vaoso.group(1).strip()

        # 3. Bóc tách thông tin Chủ sử dụng đất (Chỉ tìm trong phạm vi Mục I trước Mục II)
        sec1_text = raw_text
        m_sec_bound = re.search(r"II\.\s*Thửa đất", raw_text, re.IGNORECASE)
        if m_sec_bound:
            sec1_text = raw_text[:m_sec_bound.start()]

        owners = []
        owner_pattern = re.compile(
            r"(?:^|\n)\s*(Ông|Bà|Hộ ông|Hộ bà|Công ty|Tổ chức)\s*[:\.]?\s*([^\n,;]+)",
            re.IGNORECASE
        )
        for match in owner_pattern.finditer(sec1_text):
            danh_xung = match.group(1).strip().title()
            ho_ten = match.group(2).strip().upper()
            snippet = sec1_text[match.end(): match.end() + 260]

            nam_sinh = ""
            m_ns = re.search(r"(?:sinh năm|năm sinh)\s*[:\.]?\s*(\d{4})", snippet, re.IGNORECASE)
            if m_ns:
                nam_sinh = m_ns.group(1)

            so_cccd = ""
            m_cccd = re.search(r"(?:CMND|CCCD|Căn cước công dân|Định danh cá nhân|Giấy tờ tùy thân|số)\s*[:\.]?\s*(\d{9,12})", snippet, re.IGNORECASE)
            if m_cccd:
                so_cccd = m_cccd.group(1)

            dia_chi_tt = ""
            m_dc = re.search(r"(?:Địa chỉ thường trú|Địa chỉ)\s*[:\.]?\s*([^\n]+)", snippet, re.IGNORECASE)
            if m_dc:
                dia_chi_tt = m_dc.group(1).strip(" .,;")

            owners.append({
                "chu_su_dung": f"{danh_xung} {ho_ten}",
                "nam_sinh": nam_sinh,
                "so_giay_to": so_cccd,
                "dia_chi_thuong_tru": dia_chi_tt,
            })

        # Nếu không khớp mẫu Ông/Bà chuẩn ở đầu dòng, thử tìm dòng ngay sau tiêu đề Mục I
        if not owners:
            m_sec1 = re.search(r"I\.\s*Người sử dụng đất[^\n]*\n+([^\n]+)", sec1_text, re.IGNORECASE)
            fallback_name = m_sec1.group(1).strip() if m_sec1 else "Chưa rõ tên chủ sử dụng"
            m_any_cccd = re.search(r"\b(\d{12}|\d{9})\b", sec1_text)
            owners.append({
                "chu_su_dung": fallback_name,
                "nam_sinh": "",
                "so_giay_to": m_any_cccd.group(1) if m_any_cccd else "",
                "dia_chi_thuong_tru": "",
            })

        # 4. Thông tin thửa đất (Mục II)
        so_thua = ""
        m_thua = re.search(r"Thửa đất số\s*[:\.]?\s*([0-9A-Za-z\-/]+)", raw_text, re.IGNORECASE)
        if m_thua:
            so_thua = m_thua.group(1).strip()

        so_to = ""
        m_to = re.search(r"Tờ bản đồ số\s*[:\.]?\s*([0-9A-Za-z\-/]+)", raw_text, re.IGNORECASE)
        if m_to:
            so_to = m_to.group(1).strip()

        dia_chi_thua = ""
        m_dcthua = re.search(r"(?:b\)|2\.|-)\s*Địa chỉ(?:\s*thửa đất)?\s*[:\.]?\s*([^\n]+)", raw_text, re.IGNORECASE)
        if not m_dcthua:
            m_dcthua = re.search(r"Tờ bản đồ số[^\n]*\n+\s*(?:b\)\s*)?Địa chỉ\s*[:\.]?\s*([^\n]+)", raw_text, re.IGNORECASE)
        if m_dcthua:
            dia_chi_thua = m_dcthua.group(1).strip(" .,;")

        dien_tich = ""
        m_dt = re.search(r"Diện tích\s*[:\.]?\s*([\d\.,]+)\s*m", raw_text, re.IGNORECASE)
        if m_dt:
            dien_tich = m_dt.group(1).strip()

        hinh_thuc = ""
        m_ht = re.search(r"Hình thức sử dụng\s*[:\.]?\s*([^\n]+)", raw_text, re.IGNORECASE)
        if m_ht:
            hinh_thuc = m_ht.group(1).strip(" .,;")

        muc_dich = ""
        m_md = re.search(r"Mục đích sử dụng\s*[:\.]?\s*([^\n]+)", raw_text, re.IGNORECASE)
        if m_md:
            muc_dich = m_md.group(1).strip(" .,;")

        ma_loai_dat = cls.detect_land_code(muc_dich)

        thoi_han = ""
        m_th = re.search(r"Thời hạn sử dụng\s*[:\.]?\s*([^\n]+)", raw_text, re.IGNORECASE)
        if m_th:
            thoi_han = m_th.group(1).strip(" .,;")

        nguon_goc = ""
        m_ng = re.search(r"Nguồn gốc sử dụng\s*[:\.]?\s*([^\n]+)", raw_text, re.IGNORECASE)
        if m_ng:
            nguon_goc = m_ng.group(1).strip(" .,;")

        # 5. Nội dung biến động đất đai (Mục IV)
        bien_dong = ""
        m_bd = re.search(
            r"(?:IV\.\s*Những thay đổi sau khi cấp giấy chứng nhận|Nội dung thay đổi và cơ sở pháp lý)([\s\S]{10,600})",
            raw_text,
            re.IGNORECASE
        )
        if m_bd:
            cleaned_bd = re.sub(r"\s+", " ", m_bd.group(1)).strip()
            cleaned_bd = re.sub(r"^(?:Nội dung thay đổi và cơ sở pháp lý|Xác nhận của cơ quan có thẩm quyền)[\s\.:]*", "", cleaned_bd, flags=re.IGNORECASE).strip()
            if len(cleaned_bd) > 8:
                bien_dong = cleaned_bd[:350]

        trang_thai = "Hoàn tất" if (so_thua or so_to or dien_tich) else "Thiếu thông tin thửa"

        # Tạo danh sách bản ghi kết quả (Tách dòng theo chuẩn VNLIS hoặc Gộp chung)
        records = []
        if split_owners and len(owners) > 1:
            for ow in owners:
                records.append({
                    "ten_file": file_name,
                    "so_phat_hanh": so_phat_hanh,
                    "so_vao_so": so_vao_so,
                    "chu_su_dung": ow["chu_su_dung"],
                    "nam_sinh": ow["nam_sinh"],
                    "so_giay_to": ow["so_giay_to"],
                    "dia_chi_thuong_tru": ow["dia_chi_thuong_tru"],
                    "so_thua": so_thua,
                    "so_to_ban_do": so_to,
                    "dia_chi_thua_dat": dia_chi_thua,
                    "dien_tich": dien_tich,
                    "hinh_thuc_su_dung": hinh_thuc,
                    "muc_dich_su_dung": muc_dich,
                    "ma_loai_dat": ma_loai_dat,
                    "thoi_han_su_dung": thoi_han,
                    "nguon_goc_su_dung": nguon_goc,
                    "bien_dong_muc_iv": bien_dong,
                    "phuong_thuc_doc": read_method,
                    "trang_thai": trang_thai,
                })
        else:
            combined_names = " & ".join([o["chu_su_dung"] for o in owners])
            combined_years = ", ".join([o["nam_sinh"] for o in owners if o["nam_sinh"]])
            combined_ids = ", ".join([o["so_giay_to"] for o in owners if o["so_giay_to"]])
            first_addr = next((o["dia_chi_thuong_tru"] for o in owners if o["dia_chi_thuong_tru"]), "")
            records.append({
                "ten_file": file_name,
                "so_phat_hanh": so_phat_hanh,
                "so_vao_so": so_vao_so,
                "chu_su_dung": combined_names,
                "nam_sinh": combined_years,
                "so_giay_to": combined_ids,
                "dia_chi_thuong_tru": first_addr,
                "so_thua": so_thua,
                "so_to_ban_do": so_to,
                "dia_chi_thua_dat": dia_chi_thua,
                "dien_tich": dien_tich,
                "hinh_thuc_su_dung": hinh_thuc,
                "muc_dich_su_dung": muc_dich,
                "ma_loai_dat": ma_loai_dat,
                "thoi_han_su_dung": thoi_han,
                "nguon_goc_su_dung": nguon_goc,
                "bien_dong_muc_iv": bien_dong,
                "phuong_thuc_doc": read_method,
                "trang_thai": trang_thai,
            })

        return records


# ==============================================================================
# 3. BỘ XUẤT BÁO CÁO EXCEL CHUYÊN NGHIỆP (OPENPYXL)
# ==============================================================================
class ExcelExporter:
    """Xuất danh sách bản ghi GCN sang file Excel (.xlsx) được định dạng chuẩn."""

    COLUMNS = [
        ("STT", "stt", 6),
        ("Tên File PDF", "ten_file", 22),
        ("Số Phát Hành (Seri)", "so_phat_hanh", 16),
        ("Số Vào Sổ GCN", "so_vao_so", 16),
        ("Họ Tên Chủ Sử Dụng", "chu_su_dung", 28),
        ("Năm Sinh", "nam_sinh", 11),
        ("Số CCCD / CMND", "so_giay_to", 16),
        ("Địa Chỉ Thường Trú", "dia_chi_thuong_tru", 34),
        ("Thửa Đất Số", "so_thua", 13),
        ("Tờ Bản Đồ Số", "so_to_ban_do", 13),
        ("Địa Chỉ Thửa Đất", "dia_chi_thua_dat", 36),
        ("Diện Tích (m2)", "dien_tich", 15),
        ("Hình Thức Sử Dụng", "hinh_thuc_su_dung", 20),
        ("Mục Đích Sử Dụng", "muc_dich_su_dung", 28),
        ("Mã Loại Đất", "ma_loai_dat", 14),
        ("Thời Hạn Sử Dụng", "thoi_han_su_dung", 18),
        ("Nguồn Gốc Sử Dụng", "nguon_goc_su_dung", 28),
        ("Biến Động (Mục IV)", "bien_dong_muc_iv", 38),
        ("Nguồn Đọc", "phuong_thuc_doc", 15),
        ("Trạng Thái", "trang_thai", 16),
    ]

    @classmethod
    def export(cls, records: list[dict], output_path: str):
        import openpyxl
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        from openpyxl.utils import get_column_letter

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Danh_Sach_GCN"

        header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
        banner_fill = PatternFill(start_color="EFF6FF", end_color="EFF6FF", fill_type="solid")
        zebra_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
        warn_fill = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")

        title_font = Font(name="Segoe UI", size=13, bold=True, color="1E3A8A")
        subtitle_font = Font(name="Segoe UI", size=9, italic=True, color="475569")
        header_font = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
        cell_font = Font(name="Segoe UI", size=10, color="0F172A")
        bold_cell_font = Font(name="Segoe UI", size=10, bold=True, color="0F172A")

        thin_side = Side(style="thin", color="CBD5E1")
        border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

        total_cols = len(cls.COLUMNS)
        last_col_letter = get_column_letter(total_cols)

        ws.merge_cells(f"A1:{last_col_letter}1")
        ws["A1"] = "BẢNG TỔNG HỢP DỮ LIỆU GIẤY CHỨNG NHẬN QUYỀN SỬ DỤNG ĐẤT (GCN-ME)"
        ws["A1"].font = title_font
        ws["A1"].fill = banner_fill
        ws["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
        ws.row_dimensions[1].height = 28

        ws.merge_cells(f"A2:{last_col_letter}2")
        ws["A2"] = f"Xuất tự động bởi GCN-ME (Made by TonygamingTL) | Thời gian xuất: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')} | Tổng số dòng: {len(records)}"
        ws["A2"].font = subtitle_font
        ws["A2"].fill = banner_fill
        ws["A2"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
        ws.row_dimensions[2].height = 20

        header_row = 4
        ws.row_dimensions[header_row].height = 26
        for col_idx, (col_title, _, col_width) in enumerate(cls.COLUMNS, start=1):
            cell = ws.cell(row=header_row, column=col_idx, value=col_title)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = border
            ws.column_dimensions[get_column_letter(col_idx)].width = col_width

        for idx, rec in enumerate(records, start=1):
            row_num = header_row + idx
            ws.row_dimensions[row_num].height = 22
            is_warning = rec.get("trang_thai") != "Hoàn tất"
            row_fill = warn_fill if is_warning else (zebra_fill if idx % 2 == 0 else PatternFill(fill_type=None))

            for col_idx, (_, key, _) in enumerate(cls.COLUMNS, start=1):
                val = idx if key == "stt" else rec.get(key, "")
                cell = ws.cell(row=row_num, column=col_idx, value=val)
                cell.font = bold_cell_font if key in ("chu_su_dung", "so_thua", "so_to_ban_do") else cell_font
                if row_fill.fill_type:
                    cell.fill = row_fill
                cell.border = border

                if key in ("stt", "nam_sinh", "so_thua", "so_to_ban_do", "ma_loai_dat", "phuong_thuc_doc", "trang_thai"):
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                elif key == "dien_tich":
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

        ws.freeze_panes = "E5"
        ws.auto_filter.ref = f"A{header_row}:{last_col_letter}{header_row + max(len(records), 1)}"

        wb.save(output_path)


# ==============================================================================
# 4. GIAO DIỆN NGƯỜI DÙNG THÂN THIỆN CHO NGƯỜI MỚI (WIZARD UI)
# ==============================================================================
class GCNExtractorGUI:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("🇻🇳 GCN-ME | Phần Mềm Trích Xuất Hồ Sơ Đất Đai Tự Động (Made by TonygamingTL)")
        self.root.geometry("1080x740")
        self.root.minsize(960, 660)
        self.root.configure(bg="#F1F5F9")

        self.input_mode = tk.StringVar(value="folder")
        self.pdf_source_path = tk.StringVar()
        self.gdrive_url = tk.StringVar()
        self.output_excel_path = tk.StringVar()
        self.tesseract_path = tk.StringVar(value=self._auto_find_tesseract())
        self.use_ocr = tk.BooleanVar(value=True)
        self.split_owners = tk.BooleanVar(value=True)

        self.extracted_records = []
        self.is_running = False

        self._configure_styles()
        self._build_layout()
        self._check_dependencies_banner()

    def _auto_find_tesseract(self) -> str:
        """Tự động dò tìm đường dẫn cài đặt Tesseract OCR trên máy Windows/Linux."""
        common_paths = [
            shutil.which("tesseract") or "",
            r"C:\Program Files\Tesseract-OCR\tesseract.exe",
            r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
            os.path.expandvars(r"%LOCALAPPDATA%\Tesseract-OCR\tesseract.exe"),
            "/usr/bin/tesseract",
        ]
        for p in common_paths:
            if p and os.path.exists(p):
                return p
        return ""

    def _configure_styles(self):
        style = ttk.Style()
        style.theme_use("clam")

        style.configure("Main.TFrame", background="#F1F5F9")
        style.configure("Card.TLabelframe", background="#FFFFFF", foreground="#1E3A8A", font=("Segoe UI", 10, "bold"))
        style.configure("Card.TLabelframe.Label", background="#FFFFFF", foreground="#1E3A8A", font=("Segoe UI", 10, "bold"))
        style.configure("CardInner.TFrame", background="#FFFFFF")
        style.configure("TLabel", background="#FFFFFF", foreground="#1E293B", font=("Segoe UI", 10))
        style.configure("Hint.TLabel", background="#FFFFFF", foreground="#64748B", font=("Segoe UI", 9, "italic"))
        style.configure("TRadiobutton", background="#FFFFFF", foreground="#1E293B", font=("Segoe UI", 10))
        style.configure("TCheckbutton", background="#FFFFFF", foreground="#1E293B", font=("Segoe UI", 10))

        style.configure("Primary.TButton", font=("Segoe UI", 11, "bold"), background="#2563EB", foreground="#FFFFFF", padding=(14, 8))
        style.map("Primary.TButton", background=[("active", "#1D4ED8"), ("disabled", "#94A3B8")])

        style.configure("Success.TButton", font=("Segoe UI", 10, "bold"), background="#059669", foreground="#FFFFFF", padding=(12, 8))
        style.map("Success.TButton", background=[("active", "#047857"), ("disabled", "#CBD5E1")])

        style.configure("Secondary.TButton", font=("Segoe UI", 9, "bold"), background="#E2E8F0", foreground="#1E293B", padding=(10, 5))
        style.map("Secondary.TButton", background=[("active", "#CBD5E1")])

        style.configure("Treeview", font=("Segoe UI", 9), rowheight=26, background="#FFFFFF", fieldbackground="#FFFFFF")
        style.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"), background="#1E3A8A", foreground="#FFFFFF", padding=5)
        style.map("Treeview.Heading", background=[("active", "#1D4ED8")])

    def _build_layout(self):
        # ------------------------------------------------------------------
        # PHẦN ĐẦU TRANG (HEADER BANNER)
        # ------------------------------------------------------------------
        header = tk.Frame(self.root, bg="#1E3A8A", height=74)
        header.pack(fill="x")
        header.pack_propagate(False)

        header_left = tk.Frame(header, bg="#1E3A8A")
        header_left.pack(side="left", fill="y", padx=18, pady=10)

        tk.Label(
            header_left,
            text="🇻🇳 GCN-ME: TRÍCH XUẤT HỒ SƠ ĐẤT ĐAI TỪ PDF SANG EXCEL",
            bg="#1E3A8A", fg="#FFFFFF", font=("Segoe UI", 14, "bold")
        ).pack(anchor="w")

        tk.Label(
            header_left,
            text="Help the Vietnamese state (in matters related to land) to operate more quickly and accurately — Made by TonygamingTL",
            bg="#1E3A8A", fg="#BFDBFE", font=("Segoe UI", 9, "italic")
        ).pack(anchor="w", pady=(2, 0))

        header_right = tk.Frame(header, bg="#1E3A8A")
        header_right.pack(side="right", padx=18)

        btn_guide = tk.Button(
            header_right, text="❓ Hướng dẫn cho người mới", bg="#3B82F6", fg="#FFFFFF",
            activebackground="#2563EB", activeforeground="#FFFFFF", relief="flat",
            font=("Segoe UI", 9, "bold"), padx=10, pady=5, cursor="hand2",
            command=self._show_beginner_guide
        )
        btn_guide.pack(side="left", padx=(0, 8))

        btn_downloads = tk.Button(
            header_right, text="⬇️ Tải mục thiếu (Downloads)", bg="#059669", fg="#FFFFFF",
            activebackground="#047857", activeforeground="#FFFFFF", relief="flat",
            font=("Segoe UI", 9, "bold"), padx=10, pady=5, cursor="hand2",
            command=self._open_downloads_window
        )
        btn_downloads.pack(side="left", padx=(0, 8))

        btn_install_deps = tk.Button(
            header_right, text="🔄 Sửa PATH & Updates", bg="#0F172A", fg="#F8FAFC",
            activebackground="#334155", activeforeground="#FFFFFF", relief="flat",
            font=("Segoe UI", 9, "bold"), padx=10, pady=5, cursor="hand2",
            command=self._open_dependency_helper
        )
        btn_install_deps.pack(side="left")

        # ------------------------------------------------------------------
        # KHUNG NỘI DUNG CHÍNH
        # ------------------------------------------------------------------
        container = ttk.Frame(self.root, style="Main.TFrame", padding=12)
        container.pack(fill="both", expand=True)

        top_grid = ttk.Frame(container, style="Main.TFrame")
        top_grid.pack(fill="x", pady=(0, 10))
        top_grid.columnconfigure(0, weight=1)
        top_grid.columnconfigure(1, weight=1)

        # === BƯỚC 1: CHỌN NGUỒN FILE PDF ===
        step1 = ttk.LabelFrame(
            top_grid,
            text=" BƯỚC 1: Chọn nguồn hồ sơ PDF (Sổ đỏ / Sổ hồng) ",
            style="Card.TLabelframe", padding=12
        )
        step1.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        step1.columnconfigure(1, weight=1)

        ttk.Radiobutton(
            step1, text="📁 Thư mục chứa nhiều file PDF:", variable=self.input_mode,
            value="folder", command=self._on_mode_change
        ).grid(row=0, column=0, sticky="w", pady=4)

        self.entry_source = ttk.Entry(step1, textvariable=self.pdf_source_path)
        self.entry_source.grid(row=0, column=1, sticky="ew", padx=6, pady=4)

        self.btn_browse_source = ttk.Button(step1, text="Chọn...", style="Secondary.TButton", command=self._browse_pdf_source)
        self.btn_browse_source.grid(row=0, column=2, pady=4)

        ttk.Radiobutton(
            step1, text="📄 Chỉ chọn 1 file PDF lẻ:", variable=self.input_mode,
            value="single_file", command=self._on_mode_change
        ).grid(row=1, column=0, sticky="w", pady=4)

        ttk.Label(step1, text="(Bấm nút 'Chọn...' ở trên để lấy 1 file PDF hoặc cả thư mục)", style="Hint.TLabel").grid(
            row=1, column=1, columnspan=2, sticky="w", padx=6
        )

        ttk.Radiobutton(
            step1, text="☁️ Link thư mục Google Drive:", variable=self.input_mode,
            value="gdrive", command=self._on_mode_change
        ).grid(row=2, column=0, sticky="w", pady=4)

        self.entry_gdrive = ttk.Entry(step1, textvariable=self.gdrive_url, state="disabled")
        self.entry_gdrive.grid(row=2, column=1, columnspan=2, sticky="ew", padx=6, pady=4)

        # === BƯỚC 2: CHỌN NƠI LƯU EXCEL & CHẾ ĐỘ ĐỌC ===
        step2 = ttk.LabelFrame(
            top_grid,
            text=" BƯỚC 2: Chọn nơi lưu bảng Excel & Tùy chọn ",
            style="Card.TLabelframe", padding=12
        )
        step2.grid(row=0, column=1, sticky="nsew", padx=(6, 0))
        step2.columnconfigure(1, weight=1)

        ttk.Label(step2, text="💾 File Excel xuất ra:").grid(row=0, column=0, sticky="w", pady=4)
        ttk.Entry(step2, textvariable=self.output_excel_path).grid(row=0, column=1, sticky="ew", padx=6, pady=4)
        ttk.Button(step2, text="Đổi nơi lưu...", style="Secondary.TButton", command=self._browse_output_excel).grid(row=0, column=2, pady=4)

        ttk.Checkbutton(
            step2,
            text="🔍 Tự động bật OCR Tiếng Việt nếu gặp file PDF bản scan / chụp ảnh",
            variable=self.use_ocr
        ).grid(row=1, column=0, columnspan=3, sticky="w", pady=(6, 2))

        ttk.Checkbutton(
            step2,
            text="👥 Tách mỗi đồng chủ sử dụng đất (Ông / Bà) thành 1 dòng riêng (Chuẩn VNLIS)",
            variable=self.split_owners
        ).grid(row=2, column=0, columnspan=3, sticky="w", pady=2)

        # ------------------------------------------------------------------
        # BƯỚC 3: THANH ĐIỀU KHIỂN CHÍNH & TIẾN ĐỘ
        # ------------------------------------------------------------------
        action_bar = ttk.Frame(container, style="Main.TFrame")
        action_bar.pack(fill="x", pady=(0, 10))

        self.btn_start = ttk.Button(
            action_bar,
            text="🚀 BƯỚC 3: BẮT ĐẦU QUÉT & TRÍCH XUẤT NGAY",
            style="Primary.TButton",
            command=self._start_extraction_thread
        )
        self.btn_start.pack(side="left", padx=(0, 10))

        self.btn_open_excel = ttk.Button(
            action_bar,
            text="📊 Mở File Excel Kết Quả",
            style="Success.TButton",
            state="disabled",
            command=self._open_output_excel
        )
        self.btn_open_excel.pack(side="left", padx=(0, 8))

        self.btn_clear = ttk.Button(
            action_bar,
            text="🧹 Xóa danh sách",
            style="Secondary.TButton",
            command=self._clear_results
        )
        self.btn_clear.pack(side="left")

        progress_box = ttk.Frame(action_bar, style="Main.TFrame")
        progress_box.pack(side="right", fill="x", expand=True, padx=(15, 0))

        self.lbl_status = ttk.Label(
            progress_box,
            text="Trạng thái: Sẵn sàng — Hãy chọn thư mục PDF ở Bước 1 để bắt đầu.",
            background="#F1F5F9", foreground="#1E3A8A", font=("Segoe UI", 9, "bold")
        )
        self.lbl_status.pack(anchor="w")

        self.progress_bar = ttk.Progressbar(progress_box, orient="horizontal", mode="determinate")
        self.progress_bar.pack(fill="x", pady=(4, 0))

        # ------------------------------------------------------------------
        # BẢNG XEM TRƯỚC KẾT QUẢ TRỰC TIẾP & NHẬT KÝ (TABS DỄ NHÌN)
        # ------------------------------------------------------------------
        notebook = ttk.Notebook(container)
        notebook.pack(fill="both", expand=True)

        tab_preview = ttk.Frame(notebook, style="CardInner.TFrame", padding=6)
        notebook.add(tab_preview, text=" 📋 Bảng Xem Trước Dữ Liệu Đã Trích Xuất (0 dòng) ")
        self.notebook = notebook
        self.tab_preview = tab_preview

        preview_cols = (
            "stt", "ten_file", "so_phat_hanh", "chu_su_dung", "so_giay_to",
            "so_thua", "so_to_ban_do", "dien_tich", "ma_loai_dat", "thoi_han_su_dung", "trang_thai"
        )
        self.tree = ttk.Treeview(tab_preview, columns=preview_cols, show="headings", selectmode="browse")

        headings_config = [
            ("stt", "STT", 45, "center"),
            ("ten_file", "Tên File PDF", 150, "w"),
            ("so_phat_hanh", "Số Seri GCN", 105, "center"),
            ("chu_su_dung", "Họ Tên Chủ Sử Dụng", 190, "w"),
            ("so_giay_to", "Số CCCD/CMND", 115, "center"),
            ("so_thua", "Số Thửa", 75, "center"),
            ("so_to_ban_do", "Tờ BĐ", 70, "center"),
            ("dien_tich", "Diện Tích (m2)", 100, "e"),
            ("ma_loai_dat", "Mã Đất", 80, "center"),
            ("thoi_han_su_dung", "Thời Hạn", 110, "center"),
            ("trang_thai", "Trạng Thái", 110, "center"),
        ]
        for col_id, title, width, anchor in headings_config:
            self.tree.heading(col_id, text=title)
            self.tree.column(col_id, width=width, anchor=anchor)

        y_scroll = ttk.Scrollbar(tab_preview, orient="vertical", command=self.tree.yview)
        x_scroll = ttk.Scrollbar(tab_preview, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll.grid(row=1, column=0, sticky="ew")
        tab_preview.rowconfigure(0, weight=1)
        tab_preview.columnconfigure(0, weight=1)

        self.tree.bind("<Double-1>", self._on_row_double_click)

        tab_log = ttk.Frame(notebook, style="CardInner.TFrame", padding=6)
        notebook.add(tab_log, text=" 🖥️ Nhật Ký Quá Trình Xử Lý & Cài Đặt OCR ")

        ocr_bar = ttk.Frame(tab_log, style="CardInner.TFrame")
        ocr_bar.pack(fill="x", pady=(0, 6))
        ttk.Label(ocr_bar, text="Đường dẫn Tesseract OCR (nếu quét ảnh):").pack(side="left")
        ttk.Entry(ocr_bar, textvariable=self.tesseract_path, width=55).pack(side="left", padx=6)
        ttk.Button(ocr_bar, text="Chọn tesseract.exe...", style="Secondary.TButton", command=self._browse_tesseract).pack(side="left")

        self.txt_log = tk.Text(
            tab_log, bg="#0F172A", fg="#E2E8F0", font=("Consolas", 10),
            padx=10, pady=8, state="disabled", wrap="word"
        )
        self.txt_log.pack(fill="both", expand=True)

        footer = tk.Label(
            self.root,
            text="💡 Mẹo: Sau khi trích xuất xong, bạn có thể nhấp đúp chuột (Double-click) vào bất kỳ dòng nào trên bảng để xem toàn bộ chi tiết thửa đất & biến động Mục IV.",
            bg="#E2E8F0", fg="#334155", font=("Segoe UI", 9), anchor="w", padx=12, pady=5
        )
        footer.pack(fill="x", side="bottom")

    # ----------------------------------------------------------------------
    # CÁC HÀM XỬ LÝ SỰ KIỆN GIAO DIỆN
    # ----------------------------------------------------------------------
    def _log(self, msg: str):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.txt_log.config(state="normal")
        self.txt_log.insert("end", f"[{timestamp}] {msg}\n")
        self.txt_log.see("end")
        self.txt_log.config(state="disabled")

    def _check_dependencies_banner(self):
        missing = [pip_name for mod, pip_name in OPTIONAL_LIBS.items() if not INSTALLED_STATUS.get(mod)]
        if not missing:
            self._log("✅ Hệ thống đã cài đặt đầy đủ tất cả thư viện xử lý PDF, OCR và Excel.")
        else:
            self._log(f"⚠️ Một số thư viện bổ trợ chưa được cài đặt: {', '.join(missing)}")
            self._log("👉 Bạn có thể bấm nút '⚙️ Kiểm tra & Cài thư viện' ở góc trên bên phải để tự động cài đặt chỉ với 1 cú nhấp chuột!")

    def _on_mode_change(self):
        mode = self.input_mode.get()
        if mode in ("folder", "single_file"):
            self.entry_source.config(state="normal")
            self.btn_browse_source.config(state="normal")
            self.entry_gdrive.config(state="disabled")
        else:
            self.entry_source.config(state="disabled")
            self.btn_browse_source.config(state="disabled")
            self.entry_gdrive.config(state="normal")

    def _browse_pdf_source(self):
        mode = self.input_mode.get()
        if mode == "single_file":
            path = filedialog.askopenfilename(
                title="Chọn 1 tệp PDF Giấy chứng nhận đất đai",
                filetypes=[("Tệp PDF", "*.pdf"), ("Tất cả tệp", "*.*")]
            )
            if path:
                self.pdf_source_path.set(path)
                base_dir = os.path.dirname(path)
                self.output_excel_path.set(os.path.join(base_dir, "Ket_Qua_GCN_ME.xlsx"))
                self._log(f"📄 Đã chọn file PDF: {path}")
        else:
            folder = filedialog.askdirectory(title="Chọn thư mục chứa các tệp PDF Giấy chứng nhận")
            if folder:
                self.pdf_source_path.set(folder)
                self.output_excel_path.set(os.path.join(folder, "Tong_Hop_GCN_ME.xlsx"))
                pdf_count = len(glob.glob(os.path.join(folder, "**", "*.pdf"), recursive=True))
                self._log(f"📁 Đã chọn thư mục: {folder} (Tìm thấy {pdf_count} tệp PDF)")
                self.lbl_status.config(text=f"Trạng thái: Đã tìm thấy {pdf_count} tệp PDF. Bấm 'BƯỚC 3: BẮT ĐẦU' để chạy!")

    def _browse_output_excel(self):
        path = filedialog.asksaveasfilename(
            title="Chọn nơi lưu tệp bảng tính Excel kết quả",
            defaultextension=".xlsx",
            initialfile="Tong_Hop_GCN_ME.xlsx",
            filetypes=[("Excel Workbook (*.xlsx)", "*.xlsx")]
        )
        if path:
            self.output_excel_path.set(path)
            self._log(f"💾 Đã đặt nơi lưu kết quả Excel: {path}")

    def _browse_tesseract(self):
        path = filedialog.askopenfilename(
            title="Chọn file tesseract.exe",
            filetypes=[("Tesseract Executable", "tesseract.exe"), ("All files", "*.*")]
        )
        if path:
            self.tesseract_path.set(path)
            self._log(f"🔧 Đã cập nhật đường dẫn Tesseract OCR: {path}")

    def _clear_results(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.extracted_records.clear()
        self.notebook.tab(0, text=" 📋 Bảng Xem Trước Dữ Liệu Đã Trích Xuất (0 dòng) ")
        self.progress_bar["value"] = 0
        self.lbl_status.config(text="Trạng thái: Đã làm mới danh sách.")

    # ----------------------------------------------------------------------
    # LUỒNG XỬ LÝ TRÍCH XUẤT CHÍNH
    # ----------------------------------------------------------------------
    def _start_extraction_thread(self):
        if self.is_running:
            return

        mode = self.input_mode.get()
        if mode in ("folder", "single_file") and not self.pdf_source_path.get().strip():
            messagebox.showwarning(
                "Chưa chọn nguồn PDF (Bước 1)",
                "Bạn chưa chọn thư mục hoặc tệp PDF cần trích xuất!\n\n👉 Hãy bấm nút 'Chọn...' ở BƯỚC 1 trước nhé."
            )
            return

        if mode == "gdrive" and not self.gdrive_url.get().strip():
            messagebox.showwarning(
                "Chưa nhập link Google Drive (Bước 1)",
                "Vui lòng dán đường dẫn thư mục hoặc tệp Google Drive vào ô ở BƯỚC 1!"
            )
            return

        if not INSTALLED_STATUS.get("openpyxl"):
            messagebox.showerror(
                "Thiếu thư viện xuất Excel",
                "Máy bạn chưa cài thư viện 'openpyxl' hoặc 'pymupdf'.\n\nHãy bấm nút '⚙️ Kiểm tra & Cài thư viện' ở góc trên bên phải để cài tự động!"
            )
            return

        if not self.output_excel_path.get().strip():
            default_out = os.path.join(os.getcwd(), "Tong_Hop_GCN_ME.xlsx")
            self.output_excel_path.set(default_out)

        self.is_running = True
        self.btn_start.config(state="disabled")
        self.btn_open_excel.config(state="disabled")
        self._clear_results()

        threading.Thread(target=self._worker_extract, daemon=True).start()

    def _worker_extract(self):
        try:
            mode = self.input_mode.get()
            pdf_files = []

            if mode == "single_file":
                f_path = self.pdf_source_path.get().strip()
                if os.path.isfile(f_path):
                    pdf_files = [f_path]
            elif mode == "folder":
                folder = self.pdf_source_path.get().strip()
                pdf_files = sorted(glob.glob(os.path.join(folder, "**", "*.pdf"), recursive=True))
            elif mode == "gdrive":
                if not INSTALLED_STATUS.get("gdown"):
                    raise RuntimeError("Chưa cài đặt thư viện 'gdown'. Hãy bấm nút '⚙️ Kiểm tra & Cài thư viện' để cài.")
                import gdown
                url = self.gdrive_url.get().strip()
                temp_dir = os.path.join(os.getcwd(), "gcn_gdrive_downloads")
                os.makedirs(temp_dir, exist_ok=True)
                self.root.after(0, lambda: self._log(f"☁️ Đang tải tệp PDF từ Google Drive về {temp_dir}..."))
                if "/folders/" in url:
                    gdown.download_folder(url=url, output=temp_dir, quiet=True, use_cookies=False)
                else:
                    gdown.download(url=url, output=temp_dir + os.sep, quiet=True, fuzzy=True)
                pdf_files = sorted(glob.glob(os.path.join(temp_dir, "**", "*.pdf"), recursive=True))

            if not pdf_files:
                self.root.after(0, lambda: messagebox.showwarning(
                    "Không tìm thấy tệp PDF",
                    "Không tìm thấy tệp .pdf nào trong đường dẫn bạn đã chọn. Vui lòng kiểm tra lại thư mục!"
                ))
                return

            total = len(pdf_files)
            self.root.after(0, lambda: self._log(f"🚀 Bắt đầu bóc tách {total} tệp PDF..."))

            for idx, pdf_path in enumerate(pdf_files, start=1):
                fname = os.path.basename(pdf_path)
                pct = int((idx / total) * 100)
                self.root.after(0, lambda i=idx, t=total, fn=fname, p=pct: self._update_progress(i, t, fn, p))

                try:
                    raw_text, method_used = GCNParser.extract_text_from_pdf(
                        pdf_path,
                        use_ocr=self.use_ocr.get(),
                        tesseract_cmd=self.tesseract_path.get().strip()
                    )
                    recs = GCNParser.parse_gcn_fields(
                        raw_text,
                        file_name=fname,
                        read_method=method_used,
                        split_owners=self.split_owners.get()
                    )
                except Exception as file_err:
                    self.root.after(0, lambda fn=fname, err=file_err: self._log(f"❌ Lỗi khi đọc {fn}: {err}"))
                    recs = [{
                        "ten_file": fname,
                        "so_phat_hanh": "",
                        "so_vao_so": "",
                        "chu_su_dung": f"Lỗi đọc file: {file_err}",
                        "nam_sinh": "",
                        "so_giay_to": "",
                        "dia_chi_thuong_tru": "",
                        "so_thua": "",
                        "so_to_ban_do": "",
                        "dia_chi_thua_dat": "",
                        "dien_tich": "",
                        "hinh_thuc_su_dung": "",
                        "muc_dich_su_dung": "",
                        "ma_loai_dat": "",
                        "thoi_han_su_dung": "",
                        "nguon_goc_su_dung": "",
                        "bien_dong_muc_iv": "",
                        "phuong_thuc_doc": "Lỗi",
                        "trang_thai": "Lỗi",
                    }]

                for r in recs:
                    self.extracted_records.append(r)
                    row_index = len(self.extracted_records)
                    self.root.after(0, lambda ri=row_index, rec=r: self._append_preview_row(ri, rec))

            out_path = self.output_excel_path.get().strip()
            ExcelExporter.export(self.extracted_records, out_path)

            self.root.after(0, lambda p=out_path, cnt=len(self.extracted_records): self._on_finish_success(p, cnt))

        except Exception as e:
            self.root.after(0, lambda err=e: messagebox.showerror("Lỗi xử lý", f"Đã xảy ra lỗi trong quá trình chạy:\n{err}"))
            self.root.after(0, lambda err=e: self._log(f"❌ Lỗi dừng chương trình: {err}"))
        finally:
            self.is_running = False
            self.root.after(0, lambda: self.btn_start.config(state="normal"))

    def _update_progress(self, current: int, total: int, filename: str, percent: int):
        self.progress_bar["value"] = percent
        self.lbl_status.config(text=f"Đang xử lý ({current}/{total}): {filename} ({percent}%)")
        self._log(f"⏳ [{current}/{total}] Đang trích xuất hồ sơ: {filename}")

    def _append_preview_row(self, stt: int, rec: dict):
        values = (
            stt,
            rec.get("ten_file", ""),
            rec.get("so_phat_hanh", ""),
            rec.get("chu_su_dung", ""),
            rec.get("so_giay_to", ""),
            rec.get("so_thua", ""),
            rec.get("so_to_ban_do", ""),
            rec.get("dien_tich", ""),
            rec.get("ma_loai_dat", ""),
            rec.get("thoi_han_su_dung", ""),
            rec.get("trang_thai", ""),
        )
        self.tree.insert("", "end", values=values)
        self.notebook.tab(0, text=f" 📋 Bảng Xem Trước Dữ Liệu Đã Trích Xuất ({stt} dòng) ")

    def _on_finish_success(self, out_path: str, total_rows: int):
        self.progress_bar["value"] = 100
        self.lbl_status.config(text=f"✅ Hoàn tất! Đã xuất {total_rows} dòng dữ liệu ra file Excel.")
        self.btn_open_excel.config(state="normal")
        self._log(f"🎉 Hoàn tất xuất báo cáo Excel tại: {out_path}")
        messagebox.showinfo(
            "Trích xuất thành công!",
            f"Đã bóc tách thành công {total_rows} bản ghi hồ sơ đất đai!\n\n"
            f"📁 File Excel đã lưu tại:\n{out_path}\n\n"
            f"Bạn có thể bấm nút '📊 Mở File Excel Kết Quả' màu xanh lá để mở xem ngay."
        )

    def _open_output_excel(self):
        path = self.output_excel_path.get().strip()
        if not os.path.exists(path):
            messagebox.showwarning("Không tìm thấy tệp", "Tệp Excel chưa tồn tại hoặc đã bị đổi tên.")
            return
        if sys.platform.startswith("win"):
            os.startfile(path)
        elif sys.platform.startswith("darwin"):
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path])

    def _on_row_double_click(self, event):
        selected = self.tree.selection()
        if not selected:
            return
        item_values = self.tree.item(selected[0], "values")
        if not item_values:
            return
        idx = int(item_values[0]) - 1
        if idx < 0 or idx >= len(self.extracted_records):
            return

        rec = self.extracted_records[idx]
        detail_win = tk.Toplevel(self.root)
        detail_win.title(f"Chi tiết hồ sơ #{idx + 1} — {rec.get('ten_file', '')}")
        detail_win.geometry("620x520")
        detail_win.configure(bg="#FFFFFF")

        tk.Label(
            detail_win,
            text=f"📄 CHI TIẾT HỒ SƠ: {rec.get('ten_file', '')}",
            bg="#1E3A8A", fg="#FFFFFF", font=("Segoe UI", 11, "bold"), pady=10, padx=12, anchor="w"
        ).pack(fill="x")

        txt = tk.Text(detail_win, font=("Segoe UI", 10), padx=14, pady=12, wrap="word", bg="#F8FAFC")
        txt.pack(fill="both", expand=True, padx=10, pady=10)

        for title, key, _ in ExcelExporter.COLUMNS:
            if key == "stt":
                continue
            val = rec.get(key, "") or "(Trống)"
            txt.insert("end", f"• {title}: ", "bold")
            txt.insert("end", f"{val}\n\n")

        txt.tag_configure("bold", font=("Segoe UI", 10, "bold"), foreground="#1E3A8A")
        txt.config(state="disabled")

    def _show_beginner_guide(self):
        guide_text = (
            "HƯỚNG DẪN SỬ DỤNG NHANH TRONG 3 BƯỚC:\n\n"
            "1️⃣ BƯỚC 1 (Chọn nguồn PDF):\n"
            "   • Bấm nút 'Chọn...' ở khung Bước 1 và trỏ tới thư mục chứa các file PDF Sổ đỏ / Sổ hồng trên máy tính.\n"
            "   • Hoặc chọn 'Link thư mục Google Drive' rồi dán đường dẫn Drive công khai vào.\n\n"
            "2️⃣ BƯỚC 2 (Kiểm tra nơi lưu Excel):\n"
            "   • Phần mềm tự động đặt sẵn tên file Excel kết quả nằm ngay trong thư mục PDF của bạn (hoặc bạn bấm 'Đổi nơi lưu...' tùy ý).\n"
            "   • Nếu hồ sơ PDF là ảnh chụp/scan mờ, hãy tích chọn ô 'Tự động bật OCR Tiếng Việt'.\n\n"
            "3️⃣ BƯỚC 3 (Bắt đầu chạy & Xem kết quả):\n"
            "   • Bấm nút màu xanh '🚀 BƯỚC 3: BẮT ĐẦU QUÉT & TRÍCH XUẤT NGAY'.\n"
            "   • Kết quả sẽ hiện trực tiếp từng dòng lên bảng bên dưới.\n"
            "   • Khi chạy xong, bấm '📊 Mở File Excel Kết Quả' để mở bảng tính hoàn chỉnh!"
        )
        messagebox.showinfo("Hướng dẫn cho người mới sử dụng — GCN-ME", guide_text)

    def _open_downloads_window(self):
        """Mở trực tiếp cửa sổ SmartDownloaderGUI từ module downloads.py (hoạt động mượt trong 1 file .exe)."""
        if downloads is not None and hasattr(downloads, "open_downloader_window"):
            downloads.open_downloader_window(self.root)
            return
        self._open_dependency_helper_fallback()

    def _open_dependency_helper(self):
        """Mở trực tiếp cửa sổ SmartUpdaterGUI từ module updates.py (hoạt động mượt trong 1 file .exe)."""
        if updates is not None and hasattr(updates, "open_updater_window"):
            updates.open_updater_window(self.root)
            return
        self._open_dependency_helper_fallback()

    def _open_dependency_helper_fallback(self):
        win = tk.Toplevel(self.root)
        win.title("⚙️ Kiểm tra & Tự động cài đặt thư viện — GCN-ME")
        win.geometry("560x380")
        win.configure(bg="#FFFFFF")

        tk.Label(
            win, text="KIỂM TRA THƯ VIỆN HỆ THỐNG PYTHON",
            bg="#1E3A8A", fg="#FFFFFF", font=("Segoe UI", 11, "bold"), pady=10
        ).pack(fill="x")

        body = tk.Frame(win, bg="#FFFFFF", padx=16, pady=12)
        body.pack(fill="both", expand=True)

        for mod_name, pip_name in OPTIONAL_LIBS.items():
            ok = INSTALLED_STATUS.get(mod_name, False)
            status_str = "✅ Đã cài đặt sẵn sàng" if ok else "❌ Chưa cài đặt"
            fg_color = "#059669" if ok else "#DC2626"
            row = tk.Frame(body, bg="#FFFFFF")
            row.pack(fill="x", pady=4)
            tk.Label(row, text=f"• Thư viện {pip_name}:", bg="#FFFFFF", font=("Segoe UI", 10, "bold"), width=22, anchor="w").pack(side="left")
            tk.Label(row, text=status_str, bg="#FFFFFF", fg=fg_color, font=("Segoe UI", 10, "bold")).pack(side="left")

        def auto_install():
            btn_auto.config(state="disabled", text="⏳ Đang tự động cài đặt... Vui lòng đợi 15-30 giây")
            win.update()
            pkgs = list(OPTIONAL_LIBS.values())
            try:
                subprocess.check_call([sys.executable, "-m", "pip", "install", "--upgrade"] + pkgs)
                for m in OPTIONAL_LIBS:
                    try:
                        __import__(m)
                        INSTALLED_STATUS[m] = True
                    except ImportError:
                        pass
                messagebox.showinfo("Hoàn tất", "Đã cài đặt xong toàn bộ thư viện! Bạn có thể bắt đầu sử dụng ngay.")
                win.destroy()
            except Exception as e:
                messagebox.showerror("Lỗi cài đặt", f"Không thể tự động cài qua pip:\n{e}\n\nBạn hãy mở CMD và gõ:\npip install pymupdf pdfplumber openpyxl pytesseract Pillow gdown")
                btn_auto.config(state="normal", text="⚡ Tự động cài đặt tất cả thư viện còn thiếu")

        btn_auto = ttk.Button(
            body, text="⚡ Tự động cài đặt tất cả thư viện còn thiếu",
            style="Primary.TButton", command=auto_install
        )
        btn_auto.pack(fill="x", pady=(18, 6))


if __name__ == "__main__":
    root = tk.Tk()
    app = GCNExtractorGUI(root)
    root.mainloop()
