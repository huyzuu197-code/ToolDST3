import fitz  # PyMuPDF
import os

def xu_ly_pdf_A4_A3_Fix_Triet_De():
    input_folder = os.getcwd()
    output_folder = os.path.join(input_folder, "KET_QUA_GOM_FILE")

    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # Lọc tất cả file PDF không phân biệt chữ hoa/thường
    files = [f for f in os.listdir(input_folder) if f.lower().endswith(".pdf")]
    
    # Sắp xếp để file A4 được xử lý trước, sau đó tới A3
    files.sort(key=lambda x: ("(A4)" not in x, "(A3)" not in x))

    page_counter = 1
    blue = (0, 0, 1)

    for filename in files:
        # Mở file
        doc = fitz.open(os.path.join(input_folder, filename))
        is_a4_file = "(A4)" in filename
        
        for page in doc:
            # 1. BƯỚC XOAY: Nếu là file A4 và trang đang nằm ngang (W > H)
            # Xoay ngược chiều kim đồng hồ (Counter-clockwise)
            if is_a4_file and page.rect.width > page.rect.height:
                page.set_rotation((page.rotation + 90) % 360) #
            
            # Lấy thông số trang sau khi xoay
            rect = page.rect
            w, h = rect.width, rect.height

            # 2. CHỌN VỊ TRÍ A, B, C DỰA TRÊN SIZE TRANG THỰC TẾ
            # Đơn vị tính là Point (1/72 inch)
            if w < 610 and h > 800:  # Khổ A4 Dọc (Vị trí A)
                pos = fitz.Point(w - 60, h - 35)
                f_size = 11
            elif w < 900 and h > 1100:  # Khổ A3 Dọc (Vị trí B)
                pos = fitz.Point(w - 80, h - 45)
                f_size = 13
            elif w > 1100 and w > h:  # Khổ A3 Ngang (Vị trí C)
                pos = fitz.Point(w - 120, h - 60)
                f_size = 14
            else:  # Khổ bất kỳ hoặc bản vẽ lớn (Vị trí D)
                pos = fitz.Point(w * 0.94, h * 0.97)
                f_size = h * 0.015 # Chữ chiếm 1.5% chiều cao trang

            # 3. GHI SỐ TRANG
            page.insert_text(
                pos,
                f"Page {page_counter}",
                fontsize=f_size,
                fontname="helv",
                color=blue,
                align=fitz.TEXT_ALIGN_RIGHT, # Đảm bảo số trang luôn sát lề phải
                overlay=True
            )
            page_counter += 1

        # Lưu file với tiền tố Checked_ vào thư mục KET_QUA_GOM_FILE
        output_path = os.path.join(output_folder, f"Checked_{filename}")
        doc.save(output_path)
        doc.close()
        print(f"Đã xử lý xong: {filename}")

if __name__ == "__main__":
    xu_ly_pdf_A4_A3_Fix_Triet_De()
