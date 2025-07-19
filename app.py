from flask import Flask, render_template, request, send_file, flash, redirect, url_for, session
import os
import pandas as pd
import numpy as np
import re
from datetime import datetime
import glob
from werkzeug.utils import secure_filename
import tempfile

app = Flask(__name__)
app.secret_key = 'your-secret-key-here'  # Cần thiết cho flash messages

# Cấu hình thư mục

# Kiểm tra xem có đang chạy trên Vercel không
def is_vercel():
    return os.environ.get('VERCEL') == '1'

if is_vercel():
    # Trên Vercel, sử dụng thư mục tạm thời
    INPUT_FOLDER = tempfile.gettempdir()
    OUTPUT_FOLDER = tempfile.gettempdir()
else:
    # Local development
    INPUT_FOLDER = 'fileBeforHandle'
    OUTPUT_FOLDER = 'fileAfterHandle'

# Cấu hình upload
ALLOWED_EXTENSIONS = {'xls', 'xlsx'}
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size

# Tạo thư mục nếu chưa tồn tại (chỉ khi không phải trên Vercel)
def create_folders():
    try:
        if not os.path.exists(INPUT_FOLDER):
            os.makedirs(INPUT_FOLDER)
        if not os.path.exists(OUTPUT_FOLDER):
            os.makedirs(OUTPUT_FOLDER)
    except OSError:
        # Trên Vercel, hệ thống file là read-only, bỏ qua lỗi
        pass

# Chỉ tạo thư mục khi chạy locally
if __name__ == '__main__':
    create_folders()

def allowed_file(filename):
    """Kiểm tra xem file có được phép upload không"""
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def map_loai_hinh(to_khai):
    """
    Map cột Tờ khai thành Loại hình dựa trên công thức IF phức tạp
    """
    if pd.isna(to_khai) or not isinstance(to_khai, str):
        return to_khai

    to_khai = str(to_khai).strip()
    # Tách lấy mã loại hình ở giữa (giả sử luôn là phần thứ 2 sau dấu '/')
    parts = to_khai.split('/')
    if len(parts) >= 2:
        loai_hinh_code = parts[1].strip()
    else:
        loai_hinh_code = to_khai


    # Logic theo công thức IF
    if loai_hinh_code in ["A11", "A12", "A21", "A31", "A41", "A42", "A43"]:
        return "NKD"
    elif loai_hinh_code == "A44":
        return "NMT"
    elif loai_hinh_code in ["E11", "E13", "E15"]:
        return "NCX"
    elif loai_hinh_code in ["E21", "E23", "E41"]:
        return "NGC"
    elif loai_hinh_code in ["E31", "E33"]:
        return "NSX"
    elif loai_hinh_code in ["G11", "G12", "G13", "G14", "G51"]:
        return "TNTX"
    elif loai_hinh_code in ["C11", "C21"]:
        return "NKNQ"
    elif loai_hinh_code == "H11":
        return "NKH"
    else:
        return loai_hinh_code  # hoặc return to_khai nếu muốn giữ nguyên toàn bộ chuỗi

    
def map_loai_hinh_X(to_khai):
    """
    Map cột Tờ khai thành Loại hình dựa trên công thức IF phức tạp cho file X
    =IF(OR(A2="B11",A2="B12",A2="B13"),"XKD",IF(OR(A2="E42"),"XCX",IF(OR(A2="E52",A2="E54",A2="E56",A2="E82"),"XGC",IF(OR(A2="G21",A2="G22",A2="G23",A2="G24",A2="G61"),"TXTN",IF(A2="C22","XKNQ",IF(A2="H21","XKH","XSX")))))
    """
    if pd.isna(to_khai) or not isinstance(to_khai, str):
        return to_khai

    to_khai = str(to_khai).strip()
    # Tách lấy mã loại hình ở giữa (giả sử luôn là phần thứ 2 sau dấu '/')
    parts = to_khai.split('/')
    if len(parts) >= 2:
        loai_hinh_code = parts[1].strip()
    else:
        loai_hinh_code = to_khai

    # Logic theo công thức IF cho file X
    if loai_hinh_code in ["B11", "B12", "B13"]:
        return "XKD"
    elif loai_hinh_code == "E42":
        return "XCX"
    elif loai_hinh_code in ["E52", "E54", "E56", "E82"]:
        return "XGC"
    elif loai_hinh_code in ["G21", "G22", "G23", "G24", "G61"]:
        return "TXTN"
    elif loai_hinh_code == "C22":
        return "XKNQ"
    elif loai_hinh_code == "H21":
        return "XKH"
    else:
        return "XSX"  # Default value theo công thức Excel

def mask_sensitive_info(text, filename=None):
    """
    Ẩn thông tin nhạy cảm trong text bằng cách thay thế từ khóa và các ký tự sau nó
    - Nếu có dấu chấm/phẩy: ẩn đến dấu chấm/phẩy
    - Nếu không có: ẩn 7 ký tự (không tính khoảng trắng)
    - Trường hợp đặc biệt: từ "tín hiệu" và "hiệu ứng" sẽ không đánh dấu từ "hiệu"
    - Trường hợp đặc biệt: nếu có "100%" thì chỉ giữ lại phần từ đầu đến "100%"
    - Mask tất cả mã sản phẩm (chuỗi viết hoa/số/gạch dài từ 4 ký tự trở lên)
    """

     # Kiểm tra nếu có filename và bắt đầu bằng 'X'
    if filename and os.path.basename(filename).lower().startswith('x'):
        return mask_sensitive_info_X(text)
    if pd.isna(text) or not isinstance(text, str):
        return text   
    
    # Xử lý đặc biệt cho ký tự #& TRƯỚC KHI BẮT ĐẦU MASKING
    if isinstance(text, str):
        # Tìm vị trí #& đầu tiên và cuối cùng
        first_pos = text.find('#&')
        last_pos = text.rfind('#&')
        
        if first_pos != -1:  # Có ít nhất 1 #&
            if first_pos == last_pos:  # Chỉ có 1 #&
                # Kiểm tra xem #& có ở cuối câu không (sau #& có 2-3 ký tự)
                remaining_after_last = text[last_pos + 2:]  # +2 để bỏ qua '#&'
                if len(remaining_after_last.strip()) <= 3:  # Sau #& có 2-3 ký tự
                    # Xóa từ #& đến cuối
                    text = text[:last_pos].strip()
                else:
                    # #& ở giữa hoặc đầu, xóa từ đầu đến #&
                    text = text[first_pos + 2:].strip()
            else:  # Có nhiều #&
                # Xóa từ đầu đến #& đầu tiên
                text = text[first_pos + 2:]
                
                # Tìm #& cuối cùng trong phần còn lại
                last_pos_in_remaining = text.rfind('#&')
                if last_pos_in_remaining != -1:
                    # Kiểm tra xem #& cuối có ở cuối câu không (sau #& có 2-3 ký tự)
                    remaining_after_last = text[last_pos_in_remaining + 2:]
                    if len(remaining_after_last.strip()) <= 3:  # Sau #& có 2-3 ký tự
                        # Xóa từ #& cuối đến cuối
                        text = text[:last_pos_in_remaining].strip()
    
    # Mask tất cả mã sản phẩm (chuỗi viết hoa/số/gạch dài từ 4 ký tự trở lên)
    def mask_code(match):
        return '$' * len(match.group())
    text = re.sub(r'\b[A-Z0-9\-_/]{4,}\b', mask_code, text)
    
    # Danh sách các từ khóa cần ẩn
    sensitive_keywords = [
        'NSX', 'HIỆU', 'NHÃN HIỆU', 'THƯƠNG HIỆU', 'BRAND', 'CSSX', 'NHÀ MÁY', 'CS', 'CSXX',
        'BUYER', 'CSX', 'MFG', 'PPRODUCE', 'MANAFACTURE', 'MNF', 'NXX', 'HSX', 'HÃNG', 'CTY', 'CÔNG TY',
        'CONG TY', 'NCC', 'NPP', 'PO', 'SỐ PO', 'PO NO', 'HĐ', 'SHĐ', 'HỢP ĐỒNG', 'SỐ HỢP ĐỒNG', 'HỢP ĐỒNG SỐ', 'HĐS', 'HOP DONG SO', 'SO HOP DONG', 'CONTRACT', 'CONTRACTNO',
        'PART NUMBER', 'PART', 'SỐ SERIAL', 'SERIAL', 'SERIAL NO', 'TK', 'SERI', 'SỐ SERI', 'SERIES', 'TKHQ', 'TO KHAI', 'TỜ KHAI', 'TỜ KHAI HẢI QUAN',
        'MÃ', 'MÃ QLNB', 'QLNB', 'MÃ HÀNG', 'MA HANG', 'MH', 'SAP', 'ERP', 'C/O', 'PN', 'P.N', 'P/N', 'SN', 'S.N', 'S/N', 'CODE', 'BARCODE',
        'MODEL', 'MDEL', 'SHIP', 'SKU', 'LO', 'LOT', 'LÔ', 'BATCH', 'BATH', 'SLOT', 'ODER', 'INVOICE', 'INVOICENO', 'INVOICE NO',
        'HÓA ĐƠN', 'SỐ HÓA ĐƠN', 'HÓA ĐƠN SỐ', 'ĐG', 'ĐƠN GIÁ', 'PHÍ GIA CÔNG', 'PGC', 'PHÍ THUÊ', 'PHÍ THUÊ GIA CÔNG', 'PTGC', 'PHÍ GC', 'ĐƠN GIÁ GIA CÔNG', 'DON GIA GIA CONG', 'ĐGGC', 'PHÍ CHO THUÊ', 'KÝ HIỆU', 'KÍ HIỆU',
        'TCB', 'LTD', 'HSD', 'HẠN SỬ DỤNG', 'NGÀY SẢN XUẤT', 'HẠN DÙNG', 'SX TẠI', 'NGÀY SX', 'NXS', 'SỐ CÔNG BỐ', 'SẢN XUẤT TẠI'
    ]
    
    # Sắp xếp từ khóa theo độ dài giảm dần để ưu tiên từ khóa dài hơn
    sensitive_keywords.sort(key=len, reverse=True)
    
    masked_text = text
    offset = 0  # Để theo dõi vị trí sau khi đã thay thế
    
    # Trước tiên, bảo vệ từ "tín hiệu" và "hiệu ứng" bằng cách thay thế tạm thời
    protected_text = masked_text
    tin_hieu_placeholder = "___TIN_HIEU_PLACEHOLDER___"
    hieu_ung_placeholder = "___HIEU_UNG_PLACEHOLDER___"
    protected_text = re.sub(r'\btín hiệu\b', tin_hieu_placeholder, protected_text, flags=re.IGNORECASE)
    protected_text = re.sub(r'\bhiệu ứng\b', hieu_ung_placeholder, protected_text, flags=re.IGNORECASE)
    
    # Tìm và thay thế tất cả các từ khóa
    for keyword in sensitive_keywords:
        # Tạo pattern để tìm từ khóa như một từ hoàn chỉnh
        pattern = re.compile(r'\b' + re.escape(keyword) + r'\b', re.IGNORECASE)
        
        # Tìm tất cả các vị trí của từ khóa này
        matches = list(pattern.finditer(protected_text))
        
        # Xử lý từ cuối lên đầu để không ảnh hưởng đến vị trí của các từ khóa khác
        for match in reversed(matches):
            start_pos = match.start()
            end_pos = match.end()
            
            # Tìm ký tự tiếp theo sau từ khóa
            remaining_text = protected_text[end_pos:]
            
            # Kiểm tra xem có dấu chấm hoặc phẩy trong 50 ký tự tiếp theo không
            chars_to_check = min(50, len(remaining_text))
            check_text = remaining_text[:chars_to_check]
            
            # Tìm vị trí dấu chấm, phẩy, chấm phẩy đầu tiên
            dot_pos = check_text.find('.')
            comma_pos = check_text.find(',')
            semicolon_pos = check_text.find(';')

            # Tìm vị trí dừng gần nhất
            stop_positions = []
            if dot_pos != -1:
                stop_positions.append(dot_pos)
            if comma_pos != -1:
                stop_positions.append(comma_pos)
            if semicolon_pos != -1:
                stop_positions.append(semicolon_pos)

            if stop_positions:
                stop_pos = min(stop_positions)
            else:
                # Không có dấu câu nào, sử dụng logic 7 ký tự
                chars_to_mask = 0
                masked_chars = 0
                
                # Đếm 7 ký tự không phải khoảng trắng
                for char in remaining_text:
                    if chars_to_mask >= 7:
                        break
                    if char != ' ':
                        chars_to_mask += 1
                    masked_chars += 1
                
                # Tạo chuỗi thay thế - logic cũ: ẩn từ khóa + 7 ký tự sau
                replacement = '$' * len(keyword)  # Thay thế từ khóa
                if chars_to_mask > 0:
                    replacement += '$' * chars_to_mask  # Thêm dấu sao cho 7 ký tự sau
                
                # Thay thế trong text
                protected_text = protected_text[:start_pos] + replacement + remaining_text[masked_chars:]
                continue
            
            # Có dấu câu, ẩn đến dấu đó
            chars_to_mask = stop_pos + 1  # +1 để bao gồm cả dấu câu
            
            # Tạo chuỗi thay thế - logic cũ: ẩn từ khóa + đến dấu câu
            replacement = '$' * len(keyword)  # Thay thế từ khóa
            replacement += '$' * chars_to_mask  # Thêm dấu sao cho đến dấu câu
            
            # Thay thế trong text
            protected_text = protected_text[:start_pos] + replacement + remaining_text[chars_to_mask:]
    
    # Khôi phục từ "tín hiệu" và "hiệu ứng" về trạng thái ban đầu
    masked_text = protected_text.replace(tin_hieu_placeholder, "tín hiệu")
    masked_text = masked_text.replace(hieu_ung_placeholder, "hiệu ứng")
    
    # Xử lý đặc biệt cho cụm từ "qua sử dụng" - chỉ ẩn khi có format mã sản phẩm
    if 'qua sử dụng' in masked_text.lower():
        # Tìm vị trí cụm từ "qua sử dụng" (không phân biệt hoa thường)
        pattern = re.compile(r'\bqua sử dụng\b', re.IGNORECASE)
        matches = list(pattern.finditer(masked_text))
        
        # Xử lý từ cuối lên đầu để không ảnh hưởng đến vị trí
        for match in reversed(matches):
            start_pos = match.start()
            end_pos = match.end()
            
            # Tìm ký tự tiếp theo sau cụm từ "qua sử dụng"
            remaining_text = masked_text[end_pos:]
            
            # Regex nhận diện mã sản phẩm: ít nhất 2 nhóm liên tiếp, mỗi nhóm là chữ/số/gạch ngang, cách nhau bởi dấu cách
            product_pattern = re.compile(r'(?:\s+[A-Z0-9\-]+){2,}', re.IGNORECASE)
            if product_pattern.match(remaining_text):
                chars_to_mask = len(remaining_text)
                replacement = "qua sử dụng"
                if chars_to_mask > 0:
                    replacement += '$' * chars_to_mask
                masked_text = masked_text[:start_pos] + replacement + remaining_text[chars_to_mask:]
    
    # Sau khi masking, thay * thành $
    masked_text = masked_text.replace('*', '$')   
    
    # Xử lý đặc biệt: nếu có 'mới 100%' hoặc 'Mới 100%' thì chỉ giữ lại phần từ đầu đến '100%'
    if 'mới 100%' in masked_text:
        pos = masked_text.find('mới 100%')
        return masked_text[:pos + len('mới 100%')]
    elif 'Mới 100%' in masked_text:
        pos = masked_text.find('Mới 100%')
        return masked_text[:pos + len('Mới 100%')]
    
    return masked_text

def mask_sensitive_info_X(text):
    """
    Hàm xử lý masking cho file X (file xuất)
    - Ẩn thông tin nhạy cảm theo từ khóa
    - Nếu có dấu chấm/phẩy: ẩn đến dấu chấm/phẩy
    - Nếu không có: ẩn 7 ký tự (không tính khoảng trắng)
    - Trường hợp đặc biệt: từ "tín hiệu" và "hiệu ứng" sẽ không đánh dấu từ "hiệu"
    """
    if pd.isna(text) or not isinstance(text, str):
        return text
    
    # Xử lý đặc biệt cho ký tự #& TRƯỚC KHI BẮT ĐẦU MASKING
    if isinstance(text, str):
        # Tìm vị trí #& đầu tiên và cuối cùng
        first_pos = text.find('#&')
        last_pos = text.rfind('#&')
        
        if first_pos != -1:  # Có ít nhất 1 #&
            if first_pos == last_pos:  # Chỉ có 1 #&
                # Kiểm tra xem #& có ở cuối câu không (sau #& có 2-3 ký tự)
                remaining_after_last = text[last_pos + 2:]  # +2 để bỏ qua '#&'
                if len(remaining_after_last.strip()) <= 3:  # Sau #& có 2-3 ký tự
                    # Xóa từ #& đến cuối
                    text = text[:last_pos].strip()
                else:
                    # #& ở giữa hoặc đầu, xóa từ đầu đến #&
                    text = text[first_pos + 2:].strip()
            else:  # Có nhiều #&
                # Xóa từ đầu đến #& đầu tiên
                text = text[first_pos + 2:]
                
                # Tìm #& cuối cùng trong phần còn lại
                last_pos_in_remaining = text.rfind('#&')
                if last_pos_in_remaining != -1:
                    # Kiểm tra xem #& cuối có ở cuối câu không (sau #& có 2-3 ký tự)
                    remaining_after_last = text[last_pos_in_remaining + 2:]
                    if len(remaining_after_last.strip()) <= 3:  # Sau #& có 2-3 ký tự
                        # Xóa từ #& cuối đến cuối
                        text = text[:last_pos_in_remaining].strip()

     # Mask tất cả mã sản phẩm (chuỗi viết hoa/số/gạch dài từ 4 ký tự trở lên)
    def mask_code(match):
        return '$' * len(match.group())
    text = re.sub(r'\b[A-Z0-9\-_/]{4,}\b', mask_code, text)
    
    # Danh sách các từ khóa cần ẩn cho file X
    sensitive_keywords = [
        'NSX', 'HIỆU', 'NHÃN HIỆU', 'THƯƠNG HIỆU', 'BRAND', 'CSSX', 'NHÀ MÁY', 'CS', 'BUYER', 'CSX', 'MFG', 'PPRODUCE', 'MANAFACTURE', 'MNF', 'NXX', 'HSX', 'HÃNG', 'CTY', 'CÔNG TY', 'CONG TY', 'NCC', 'NPP', 'PO', 'SỐ PO', 'PO NO', 'SHĐ', 'SỐ HỢP ĐỒNG', 'HỢP ĐỒNG SỐ', 'HĐS', 'HOP DONG SO', 'SO HOP DONG', 'CONTRACT', 'CONTRACTNO', 'PART NUMBER', 'PART', 'SỐ SERIAL', 'SERIAL', 'SERIAL NO', 'TK', 'TKHQ', 'TO KHAI', 'TỜ KHAI', 'TỜ KHAI HẢI QUAN', 'MÃ', 'MÃ QLNB', 'QLNB', 'MÃ HÀNG', 'MA HANG', 'MH', 'SAP', 'ERP', 'C/O', 'PN', 'P.N', 'P/N', 'SN', 'S.N', 'S/N', 'CODE', 'BARCODE', 'MODEL', 'SHIP', 'SKU', 'LO', 'LOT', 'LÔ', 'BATCH', 'BATH', 'SLOT', 'ODER', 'INVOICE', 'INVOICENO', 'INVOICE NO', 'HÓA ĐƠN', 'SỐ HÓA ĐƠN', 'HÓA ĐƠN SỐ', 'ĐG', 'ĐƠN GIÁ', 'PHÍ GIA CÔNG', 'PGC', 'PHÍ THUÊ', 'PHÍ THUÊ GIA CÔNG', 'PTGC', 'PHÍ GC', 'ĐƠN GIÁ GIA CÔNG', 'DON GIA GIA CONG', 'ĐGGC', 'PHÍ CHO THUÊ', 'KÝ HIỆU', 'KÍ HIỆU',
        'TCB', 'LTD', 'HSD', 'HẠN SỬ DỤNG', 'NGÀY SẢN XUẤT', 'HẠN DÙNG', 'SX TẠI', 'NGÀY SX', 'NXS', 'SỐ CÔNG BỐ', 'SẢN XUẤT TẠI'
    ]
    
    # Sắp xếp từ khóa theo độ dài giảm dần để ưu tiên từ khóa dài hơn
    sensitive_keywords.sort(key=len, reverse=True)
    
    masked_text = text
    offset = 0  # Để theo dõi vị trí sau khi đã thay thế
    
    # Trước tiên, bảo vệ từ "tín hiệu" và "hiệu ứng" bằng cách thay thế tạm thời
    protected_text = masked_text
    tin_hieu_placeholder = "___TIN_HIEU_PLACEHOLDER___"
    hieu_ung_placeholder = "___HIEU_UNG_PLACEHOLDER___"
    protected_text = re.sub(r'\btín hiệu\b', tin_hieu_placeholder, protected_text, flags=re.IGNORECASE)
    protected_text = re.sub(r'\bhiệu ứng\b', hieu_ung_placeholder, protected_text, flags=re.IGNORECASE)
    
    # Tìm và thay thế tất cả các từ khóa
    for keyword in sensitive_keywords:
        # Tạo pattern để tìm từ khóa như một từ hoàn chỉnh
        pattern = re.compile(r'\b' + re.escape(keyword) + r'\b', re.IGNORECASE)
        
        # Tìm tất cả các vị trí của từ khóa này
        matches = list(pattern.finditer(protected_text))
        
        # Xử lý từ cuối lên đầu để không ảnh hưởng đến vị trí của các từ khóa khác
        for match in reversed(matches):
            start_pos = match.start()
            end_pos = match.end()
            
            # Tìm ký tự tiếp theo sau từ khóa
            remaining_text = protected_text[end_pos:]
            
            # Kiểm tra xem có dấu chấm hoặc phẩy trong 50 ký tự tiếp theo không
            chars_to_check = min(50, len(remaining_text))
            check_text = remaining_text[:chars_to_check]
            
            # Tìm vị trí dấu chấm, phẩy, chấm phẩy đầu tiên
            dot_pos = check_text.find('.')
            comma_pos = check_text.find(',')
            semicolon_pos = check_text.find(';')

            # Tìm vị trí dừng gần nhất
            stop_positions = []
            if dot_pos != -1:
                stop_positions.append(dot_pos)
            if comma_pos != -1:
                stop_positions.append(comma_pos)
            if semicolon_pos != -1:
                stop_positions.append(semicolon_pos)

            if stop_positions:
                stop_pos = min(stop_positions)
                # Có dấu câu, ẩn đến dấu đó
                chars_to_mask = stop_pos + 1  # +1 để bao gồm cả dấu câu
                
                # Tạo chuỗi thay thế - ẩn từ khóa + đến dấu câu
                replacement = '$' * len(keyword)  # Thay thế từ khóa
                replacement += '$' * chars_to_mask  # Thêm dấu sao cho đến dấu câu
                
                # Thay thế trong text
                protected_text = protected_text[:start_pos] + replacement + remaining_text[chars_to_mask:]
            else:
                # Không có dấu câu nào, sử dụng logic 7 ký tự không phải khoảng trắng
                chars_to_mask = 0
                masked_chars = 0
                
                # Đếm 7 ký tự không phải khoảng trắng
                for char in remaining_text:
                    if chars_to_mask >= 7:
                        break
                    if char != ' ':
                        chars_to_mask += 1
                    masked_chars += 1
                
                # Tạo chuỗi thay thế - ẩn từ khóa + 7 ký tự sau
                replacement = '$' * len(keyword)  # Thay thế từ khóa
                if chars_to_mask > 0:
                    replacement += '$' * chars_to_mask  # Thêm dấu sao cho 7 ký tự sau
                
                # Thay thế trong text
                protected_text = protected_text[:start_pos] + replacement + remaining_text[masked_chars:]
    
    # Khôi phục từ "tín hiệu" và "hiệu ứng" về trạng thái ban đầu
    masked_text = protected_text.replace(tin_hieu_placeholder, "tín hiệu")
    masked_text = masked_text.replace(hieu_ung_placeholder, "hiệu ứng")
    
    # Sau khi masking, thay * thành $
    masked_text = masked_text.replace('$', '$')
    
    # Xử lý đặc biệt: nếu có 'mới 100%' hoặc 'Mới 100%' thì chỉ giữ lại phần từ đầu đến '100%'
    if 'mới 100%' in masked_text:
        pos = masked_text.find('mới 100%')
        return masked_text[:pos + len('mới 100%')]
    elif 'Mới 100%' in masked_text:
        pos = masked_text.find('Mới 100%')
        return masked_text[:pos + len('Mới 100%')]
    
    return masked_text

def extract_xuat_xu_from_x_file(text):
    """
    Trích xuất xuất xứ từ tên hàng trong file X
    - Lấy các ký tự sau #& cuối cùng trong tên hàng
    - Ví dụ: "SP05437#&Nhãn nhựa tự dính đã in kích thước 33.8*16.2mm: 72-02331601#&VN" -> "VN"
    """
    if pd.isna(text) or not isinstance(text, str):
        return ""
    
    # Tìm vị trí #& cuối cùng
    last_pos = text.rfind('#&')
    if last_pos != -1:
        # Lấy các ký tự sau #& cuối cùng
        xuat_xu = text[last_pos + 2:]  # +2 để bỏ qua '#&'
        return xuat_xu.strip()
    
    return ""

def process_data(file_path):
    """Xử lý file Excel và trả về DataFrame đã xử lý"""
    try:
        # Đọc file gốc
        
        # Xác định engine dựa trên định dạng file
        file_extension = os.path.splitext(file_path)[1].lower()
        
        # Thử đọc file với các engine khác nhau
        df_original = None
        error_messages = []
        
        if file_extension == '.xls':
            # Thử với xlrd trước
            try:
                df_original = pd.read_excel(file_path, engine='xlrd')
            except Exception as e1:
                error_messages.append(f"xlrd engine: {str(e1)}")
                # Thử với openpyxl nếu xlrd thất bại
                try:
                    df_original = pd.read_excel(file_path, engine='openpyxl')
                except Exception as e2:
                    error_messages.append(f"openpyxl engine: {str(e2)}")
        else:
            # Thử với openpyxl trước cho .xlsx
            try:
                df_original = pd.read_excel(file_path, engine='openpyxl')
            except Exception as e1:
                error_messages.append(f"openpyxl engine: {str(e1)}")
                # Thử với xlrd nếu openpyxl thất bại
                try:
                    df_original = pd.read_excel(file_path, engine='xlrd')
                except Exception as e2:
                    error_messages.append(f"xlrd engine: {str(e2)}")
        
        # Nếu không đọc được với bất kỳ engine nào
        if df_original is None:
            # Thử đọc như file HTML (có thể file HTML được lưu với đuôi .xls/.xlsx)
            try:
                print("Thử đọc file như HTML...")
                # Thử với encoding utf-8 trước
                try:
                    df_original = pd.read_html(file_path, encoding='utf-8')[0]  # Lấy bảng đầu tiên
                except:
                    # Nếu utf-8 không được, thử với encoding khác
                    try:
                        df_original = pd.read_html(file_path, encoding='latin-1')[0]
                    except:
                        # Cuối cùng thử không chỉ định encoding
                        df_original = pd.read_html(file_path)[0]
            except Exception as e3:
                error_messages.append(f"HTML reader: {str(e3)}")
                error_msg = f"Không thể đọc file {os.path.basename(file_path)}. "
                error_msg += "File có thể bị hỏng hoặc không phải định dạng Excel/HTML hợp lệ. "
                error_msg += f"Chi tiết lỗi: {'; '.join(error_messages)}"
                raise Exception(error_msg)
            
        # Nếu cột là số thứ tự (0,1,2...), lấy dòng đầu làm header
        if all(str(col).isdigit() for col in df_original.columns):
            new_header = df_original.iloc[0]
            df_original = df_original[1:]
            df_original.columns = new_header
            df_original = df_original.reset_index(drop=True)
        
        # Tạo DataFrame mới với cấu trúc chuẩn
        df_processed = pd.DataFrame()
        
        # Mapping các cột từ file gốc sang file mới (key: tên cột mới, value: tên cột gốc)
        # Kiểm tra nếu file bắt đầu bằng 'X' thì sử dụng mapping khác
        if os.path.basename(file_path).lower().startswith('x'):
            # Mapping cho file X (file xuất)
            column_mapping = {
                'Ngày xuất': 'Ngày đăng ký',
                'Đơn vị đối tác': 'Đơn vị đối tác', 
                'Mã hs xuất': 'Mã hàng khai báo',
                'Tên hàng': 'Tên hàng',
                'Tờ khai': 'Tờ khai',  # Thêm cột Tờ khai để sử dụng cho mapping Loại hình
                'Loại hình': 'PP khai báo',
                'Đơn vị tính': 'Đơn vị tính',
                'Xuất xứ': 'Tên nuớc xuất xứ',
                'Điều kiện giao hàng': 'Điều kiện giao hàng',
                'Thuế suất XNK': 'Thuế suất XNK',
                'Thuế suất TTĐB': 'Thuế suất TTĐB',
                'Thuế suất VAT': 'Thuế suất VAT',
                'Thuế suất TVCBPG': 'Thuế suất tự vệ',
                'Thuế suất BVMT': 'Thuế môi trường'
            }
        else:
            # Mapping cho file N (file nhập)
            column_mapping = {
                'Ngày nhập': 'Ngày đăng ký',
                'Nhà cung cấp': 'Đơn vị đối tác', 
                'Mã hs nhập': 'Mã hàng khai báo',
                'Tên hàng': 'Tên hàng',
                'Tờ khai': 'Tờ khai',  # Thêm cột Tờ khai để sử dụng cho mapping Loại hình
                'Loại hình': 'PP khai báo',
                'Đơn vị tính': 'Đơn vị tính',
                'Xuất xứ': 'Tên nuớc xuất xứ',
                'Điều kiện giao hàng': 'Điều kiện giao hàng',
                'Thuế suất XNK': 'Thuế suất XNK',
                'Thuế suất TTĐB': 'Thuế suất TTĐB',
                'Thuế suất VAT': 'Thuế suất VAT',
                'Thuế suất TVCBPG': 'Thuế suất tự vệ',
                'Thuế suất BVMT': 'Thuế môi trường'
            }
        
        # Chuyển đổi dữ liệu theo mapping
        mapped_columns = []
        unmapped_columns = []
        fuzzy_matches = []
        
        for new_col, old_col in column_mapping.items():
            if old_col in df_original.columns:
                # Xử lý đặc biệt cho cột Xuất xứ trong file X
                if new_col == 'Xuất xứ' and os.path.basename(file_path).lower().startswith('x'):
                    # Lấy xuất xứ từ cột Tên hàng cho file X
                    if 'Tên hàng' in df_original.columns:
                        df_processed[new_col] = pd.Series(df_original['Tên hàng']).apply(extract_xuat_xu_from_x_file)
                    else:
                        df_processed[new_col] = df_original[old_col]
                else:
                    df_processed[new_col] = df_original[old_col]
                mapped_columns.append(f"'{old_col}' -> '{new_col}'")
            else:
                # Thử tìm cột tương tự
                found_similar = False
                for actual_col in df_original.columns:
                    # Chuyển đổi tên cột thành string để xử lý
                    actual_col_str = str(actual_col)
                    old_col_str = str(old_col)
                    
                                    # Kiểm tra nếu tên cột chứa từ khóa - cải thiện logic để tránh nhầm lẫn
                if new_col == 'Xuất xứ':
                    # Đối với cột Xuất xứ, chỉ map chính xác với "Tên nước xuất xứ" hoặc các biến thể gần đúng
                    if (actual_col_str.lower() == 'tên nước xuất xứ' or 
                        actual_col_str.lower() == 'tên nuớc xuất xứ' or
                        actual_col_str.lower() == 'nước xuất xứ' or
                        actual_col_str.lower() == 'xuất xứ'):
                        df_processed[new_col] = df_original[actual_col]
                        fuzzy_matches.append(f"'{actual_col_str}' -> '{new_col}' (fuzzy match cho '{old_col_str}')")
                        found_similar = True
                        break
                else:
                    # Đối với các cột khác, sử dụng logic fuzzy matching cũ
                    if (old_col_str.lower() in actual_col_str.lower() or 
                        actual_col_str.lower() in old_col_str.lower() or
                        any(keyword in actual_col_str.lower() for keyword in old_col_str.lower().split())):
                        df_processed[new_col] = df_original[actual_col]
                        fuzzy_matches.append(f"'{actual_col_str}' -> '{new_col}' (fuzzy match cho '{old_col_str}')")
                        found_similar = True
                        break
                
                if not found_similar:
                    unmapped_columns.append(old_col)
                    df_processed[new_col] = np.nan
        
        # Kiểm tra các cột quan trọng có dữ liệu không
        if os.path.basename(file_path).lower().startswith('x'):
            # File X: kiểm tra các cột quan trọng
            important_columns = ['Ngày xuất', 'Đơn vị đối tác', 'Mã hs xuất', 'Tên hàng']
        else:
            # File N: kiểm tra các cột quan trọng
            important_columns = ['Ngày nhập', 'Nhà cung cấp', 'Mã hs nhập', 'Tên hàng']
        
        # Xử lý dữ liệu ngày tháng
        if os.path.basename(file_path).lower().startswith('x'):
            # File X: xử lý ngày xuất
            if 'Ngày xuất' in df_processed.columns:
                df_processed['Ngày xuất'] = pd.to_datetime(df_processed['Ngày xuất'], dayfirst=True, errors='coerce')
                # Format ngày theo định dạng dd/mm/yyyy
                df_processed['Ngày xuất'] = df_processed['Ngày xuất'].dt.strftime('%d/%m/%Y')
        else:
            # File N: xử lý ngày nhập
            if 'Ngày nhập' in df_processed.columns:
                df_processed['Ngày nhập'] = pd.to_datetime(df_processed['Ngày nhập'], dayfirst=True, errors='coerce')
                # Format ngày theo định dạng dd/mm/yyyy
                df_processed['Ngày nhập'] = df_processed['Ngày nhập'].dt.strftime('%d/%m/%Y')
        
        # Xử lý mã HS - loại bỏ dấu nháy đơn nếu có
        if os.path.basename(file_path).lower().startswith('x'):
            # File X: xử lý mã hs xuất
            if 'Mã hs xuất' in df_processed.columns:
                df_processed['Mã hs xuất'] = df_processed['Mã hs xuất'].astype(str).str.replace("'", "")
        else:
            # File N: xử lý mã hs nhập
            if 'Mã hs nhập' in df_processed.columns:
                df_processed['Mã hs nhập'] = df_processed['Mã hs nhập'].astype(str).str.replace("'", "")
        
        # Xử lý thuế suất BVMT - chuyển đổi 0/1 thành KCT/CT
        if 'Thuế suất BVMT' in df_processed.columns:
            df_processed['Thuế suất BVMT'] = df_processed['Thuế suất BVMT'].apply(
                lambda x: 'KCT' if pd.notna(x) and str(x).strip() == '0' else ('CT' if pd.notna(x) and str(x).strip() == '1' else x)
            )
        
        # Xử lý cột Loại hình dựa trên cột Tờ khai
        if 'Loại hình' in df_processed.columns and 'Tờ khai' in df_processed.columns:
            # Lưu giá trị gốc của cột Loại hình trước khi xử lý
            original_loai_hinh = df_processed['Loại hình'].copy()
            
            # Áp dụng logic mapping dựa trên loại file
            if os.path.basename(file_path).lower().startswith('x'):
                # Sử dụng logic mapping cho file X
                df_processed['Loại hình'] = df_processed['Tờ khai'].apply(map_loai_hinh_X)
            else:
                # Sử dụng logic mapping cho file N
                df_processed['Loại hình'] = df_processed['Tờ khai'].apply(map_loai_hinh)
        
        # Ẩn thông tin nhạy cảm trong cột Tên hàng
        if 'Tên hàng' in df_processed.columns:
            # Áp dụng masking dựa trên loại file
            if os.path.basename(file_path).lower().startswith('x'):
                # Sử dụng masking logic cho file X
                df_processed['Tên hàng'] = df_processed['Tên hàng'].apply(lambda x: mask_sensitive_info(x, file_path))
            else:
                # Sử dụng masking logic cho file N
                df_processed['Tên hàng'] = df_processed['Tên hàng'].apply(lambda x: mask_sensitive_info(x, file_path))
        
        # Sắp xếp theo ngày (nhập hoặc xuất tùy theo loại file)
        if os.path.basename(file_path).lower().startswith('x'):
            # File X: sắp xếp theo ngày xuất
            if 'Ngày xuất' in df_processed.columns:
                df_processed = df_processed.sort_values('Ngày xuất')
        else:
            # File N: sắp xếp theo ngày nhập
            if 'Ngày nhập' in df_processed.columns:
                df_processed = df_processed.sort_values('Ngày nhập')
        
        # Reset index
        df_processed = df_processed.reset_index(drop=True)
        

        
        # Xóa cột 'Tờ khai' khỏi file sau khi xử lý
        if 'Tờ khai' in df_processed.columns:
            df_processed = df_processed.drop(columns=['Tờ khai'])
        
        return df_processed
        
    except Exception as e:
        raise e

def get_files_in_input_folder():
    """Lấy danh sách file Excel trong thư mục input"""
    excel_files = []
    for ext in ['*.xls', '*.xlsx']:
        excel_files.extend(glob.glob(os.path.join(INPUT_FOLDER, ext)))
    return excel_files

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # Kiểm tra action từ form
        action = request.form.get('action', '')
        
        if action == 'upload':
            # Xử lý upload file
            if 'file' not in request.files:
                flash('Không có file được chọn', 'error')
                return redirect(request.url)
            
            file = request.files['file']
            
            if file.filename == '':
                flash('Vui lòng chọn file để upload', 'error')
                return redirect(request.url)
            
            if file and file.filename and allowed_file(file.filename):
                filename = secure_filename(file.filename)
                file_path = os.path.join(INPUT_FOLDER, filename)
                
                try:
                    file.save(file_path)
                    flash(f'File "{filename}" đã được upload thành công!', 'success')
                    return redirect(request.url)
                except Exception as e:
                    flash(f'Lỗi khi upload file: {str(e)}', 'error')
                    return redirect(request.url)
            else:
                flash('Định dạng file không được phép. Vui lòng chọn file .xls hoặc .xlsx', 'error')
                return redirect(request.url)
        
        elif action == 'handle':
            # Xử lý handle file
            excel_files = get_files_in_input_folder()
            
            if not excel_files:
                flash('Không tìm thấy file Excel nào trong thư mục fileBeforHandle')
                return redirect(request.url)
            
            try:
                # Xử lý file đầu tiên tìm thấy
                input_file = excel_files[0]
                filename = os.path.basename(input_file)
                
                # Xử lý file
                df_processed = process_data(input_file)
                
                # Tạo tên file output: tên gốc + "_done"
                name_without_ext = os.path.splitext(filename)[0]
                output_filename = f"{name_without_ext}_done.xlsx"
                output_path = os.path.join(OUTPUT_FOLDER, output_filename)
                
                df_processed.to_excel(output_path, index=False, engine='openpyxl')
                
                # Lưu thông tin file để sử dụng cho download
                session['processed_file'] = output_path
                session['input_file'] = input_file
                session['output_filename'] = output_filename
                
                return redirect(url_for('download_choice'))
                
            except Exception as e:
                flash(f'Lỗi khi xử lý file: {str(e)}')
                return redirect(request.url)
    
    # Hiển thị danh sách file trong thư mục input
    excel_files = get_files_in_input_folder()
    return render_template('index.html', files=excel_files)



@app.route('/delete_file', methods=['POST'])
def delete_file():
    """Xóa file từ thư mục input"""
    try:
        data = request.get_json()
        filename = data.get('filename')
        
        if not filename:
            return {'error': 'Không có tên file'}, 400
        
        file_path = os.path.join(INPUT_FOLDER, filename)
        
        if os.path.exists(file_path):
            os.remove(file_path)
            return {'success': True}, 200
        else:
            return {'error': 'File không tồn tại'}, 404
            
    except Exception as e:
        return {'error': str(e)}, 500

@app.route('/download_choice')
def download_choice():
    """Trang hỏi người dùng có muốn download file không"""
    if 'processed_file' not in session:
        flash('Không có file nào được xử lý', 'error')
        return redirect(url_for('index'))
    
    return render_template('download_choice.html')

@app.route('/download_and_cleanup')
def download_and_cleanup():
    """Download file và xóa cả file input và output"""
    if 'processed_file' not in session or 'input_file' not in session:
        flash('Không có file nào để download', 'error')
        return redirect(url_for('index'))
    
    try:
        output_path = session['processed_file']
        input_file = session['input_file']
        output_filename = session['output_filename']
        
        # Lưu thông tin để xóa sau
        session['files_to_delete'] = [input_file, output_path]
        
        # Trả về file để download
        return send_file(output_path, as_attachment=True, download_name=output_filename)
        
    except Exception as e:
        flash(f'Lỗi khi download file: {str(e)}', 'error')
        return redirect(url_for('index'))

@app.route('/cleanup_after_download')
def cleanup_after_download():
    """Xóa file sau khi download (được gọi bởi JavaScript)"""
    if 'files_to_delete' in session:
        files_to_delete = session['files_to_delete']
        
        for file_path in files_to_delete:
            if os.path.exists(file_path):
                try:
                    os.remove(file_path)
                except Exception as e:
                    pass
        
        # Xóa session
        session.pop('files_to_delete', None)
        session.pop('processed_file', None)
        session.pop('input_file', None)
        session.pop('output_filename', None)
        
        flash('Đã download và xóa file thành công!', 'success')
    
    return redirect(url_for('index'))

@app.route('/skip_download')
def skip_download():
    """Bỏ qua download và quay về trang chính"""
    # Xóa session
    session.pop('processed_file', None)
    session.pop('input_file', None)
    session.pop('output_filename', None)
    
    flash('Đã bỏ qua download. File vẫn được lưu trong thư mục fileAfterHandle', 'success')
    return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=8086) 