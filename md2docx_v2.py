import os
import pypandoc
import sys

def preprocess_markdown(content):
    """
    Hàm tiền xử lý nội dung Markdown để sửa các lỗi cú pháp LaTeX đã biết.
    """
    print("...Thực hiện tiền xử lý nội dung Markdown để sửa lỗi cú pháp LaTeX...")

    # Sửa lỗi 1: Dư dấu gạch chéo ngược trong công thức gamma
    content = content.replace('$\\\\gamma', '$\\gamma')
    print("    - Đã sửa lỗi '\\\\gamma'.")

    # Sửa lỗi 2: Chuyển đổi công thức hàm g(x,y) không hợp lệ sang cú pháp "cases" của LaTeX
    invalid_g_func = r'$$g\left(x,y\right)=\left{max_val0​if f\left(x,y\right)>Totherwise​$$'
    valid_g_func = r'$$g(x,y) = \begin{cases} \text{max\_val} & \text{if } f(x,y) > T \\ 0 & \text{otherwise} \end{cases}$$'
    if invalid_g_func in content:
        content = content.replace(invalid_g_func, valid_g_func)
        print("    - Đã sửa lỗi công thức hàm g(x,y).")
        
    # Sửa lỗi 3: Chuyển đổi ma trận K không hợp lệ sang cú pháp "pmatrix" của LaTeX
    invalid_k_matrix = r'$$K=\left​f_{x}00​0f_{y}0​c_{x}c_{y}1​\right​$$'
    valid_k_matrix = r'$$K = \begin{pmatrix} f_x & 0 & c_x \\ 0 & f_y & c_y \\ 0 & 0 & 1 \end{pmatrix}$$'
    if invalid_k_matrix in content:
        content = content.replace(invalid_k_matrix, valid_k_matrix)
        print("    - Đã sửa lỗi công thức ma trận K.")

    print("...Tiền xử lý hoàn tất.")
    return content


def convert_md_to_docx(md_file_path, docx_file_path):
    """
    Đọc file Markdown, tiền xử lý nội dung để sửa lỗi, sau đó chuyển đổi
    sang DOCX với công thức toán học.
    """
    print(f"▶️  Bắt đầu chuyển đổi file: {md_file_path}")
    
    # Bước 1: Đọc toàn bộ nội dung của file markdown
    try:
        with open(md_file_path, 'r', encoding='utf-8') as f:
            md_content = f.read()
    except Exception as e:
        print(f"❌ Lỗi nghiêm trọng khi đọc file Markdown: {e}", file=sys.stderr)
        return

    # Bước 2: Tiền xử lý nội dung để sửa các lỗi đã biết
    processed_content = preprocess_markdown(md_content)

    # --- Các thiết lập cho Pandoc (giữ nguyên) ---
    input_format = 'markdown+tex_math_dollars'
    extra_args = ['--standalone', '--mathml']

    print("⚙️  Đang gọi Pandoc để thực hiện chuyển đổi...")
    print(f"    - Định dạng đầu vào: {input_format}")
    print(f"    - Đối số thêm: {extra_args}")

    try:
        # Bước 3: Chuyển đổi nội dung đã được xử lý (thay vì file)
        pypandoc.convert_text(
            source=processed_content,
            to='docx',
            format=input_format,
            outputfile=docx_file_path,
            extra_args=extra_args
        )
        print(f"✅🏆 Thành công! File DOCX đã được lưu tại: {docx_file_path}")

    except Exception as e:
        print(f"❌ Đã xảy ra lỗi trong quá trình chuyển đổi bằng Pandoc.", file=sys.stderr)
        print(f"Lỗi: {e}", file=sys.stderr)
        print("\n--- GỢI Ý SỬA LỖI ---", file=sys.stderr)
        print("1. Hãy chắc chắn rằng bạn đã cài đặt Pandoc trên máy tính.", file=sys.stderr)
        print("   (Truy cập: https://pandoc.org/installing.html)", file=sys.stderr)
        print("2. Kiểm tra xem đường dẫn file đầu vào có chính xác không.", file=sys.stderr)
        print("--------------------", file=sys.stderr)


if __name__ == '__main__':
    # --- THIẾT LẬP CÁC ĐƯỜNG DẪN FILE ---
    input_file = 'document\ComputerVision\Gemini-Giáo Trình Thị Giác Máy Tính.md'
    output_file = 'document\ComputerVision\Gemini-Giáo Trình Thị Giác Máy Tính_Output_v2.docx'

    # --- BẮT ĐẦU QUÁ TRÌNH XỬ LÝ ---
    print("\n--- CHƯƠNG TRÌNH CHUYỂN ĐỔI MARKDOWN SANG DOCX (v2 - Sửa lỗi tự động) ---")
    if not os.path.exists(input_file):
        print(f"Lỗi: File đầu vào '{input_file}' không tồn tại.", file=sys.stderr)
    else:
        convert_md_to_docx(input_file, output_file)
    print("---------------------------------------------------------------------")
    