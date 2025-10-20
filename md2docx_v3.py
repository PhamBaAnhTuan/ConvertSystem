import os
import pypandoc
import sys
import re # Import thư viện Biểu thức chính quy

def preprocess_markdown_v3(content):
    """
    Hàm tiền xử lý nội dung Markdown để sửa các lỗi cú pháp LaTeX một cách hệ thống
    bằng Biểu thức chính quy (Regex).
    """
    print("...Thực hiện tiền xử lý (v3) bằng Biểu thức chính quy...")

    # 1. Sửa lỗi ba dấu đô la ($$$) thành hai dấu ($$)
    # Thao tác này an toàn để thực hiện trên toàn bộ nội dung.
    if '$$$' in content:
        content = content.replace('$$$', '$$')
        print("    - Đã chuẩn hóa '$$$' thành '$$'.")

    # 2. Sửa lỗi dư dấu gạch chéo ngược cho các lệnh LaTeX (ví dụ: \\sum -> \sum)
    # Regex: Tìm chuỗi '\\' theo sau là một hoặc nhiều chữ cái (a-z, A-Z)
    # và thay thế bằng '\' + chuỗi chữ cái đó.
    original_pattern_count = len(re.findall(r'\\\\([a-zA-Z]+)', content))
    if original_pattern_count > 0:
        content = re.sub(r'\\\\([a-zA-Z]+)', r'\\\1', content)
        print(f"    - Đã sửa {original_pattern_count} lỗi lệnh LaTeX dư dấu '\\' (ví dụ: \\sum, \\lambda, \\cdot).")
        
    # 3. Sửa lỗi dư dấu gạch chéo ngược cho ký tự gạch dưới (ví dụ: \_ -> _)
    # Lỗi này thường xảy ra khi muốn viết chỉ số dưới (subscript)
    original_underscore_count = content.count(r'\_')
    if original_underscore_count > 0:
        content = content.replace(r'\_', '_')
        print(f"    - Đã sửa {original_underscore_count} lỗi ký tự gạch dưới '\\_'.")

    # 4. Sửa lại các trường hợp cụ thể còn lại (nếu có)
    # Đây là các lỗi từ phiên bản v2, vẫn giữ lại để đảm bảo tính toàn diện
    invalid_g_func = r'$$g\left(x,y\right)=\left{max_val0​if f\left(x,y\right)>Totherwise$$'
    valid_g_func = r'$$g(x,y) = \begin{cases} \text{max_val} & \text{if } f(x,y) > T \\ 0 & \text{otherwise} \end{cases}$$'
    if invalid_g_func in content:
        content = content.replace(invalid_g_func, valid_g_func)
        print("    - Đã sửa lỗi công thức hàm g(x,y) không hợp lệ.")
        
    invalid_k_matrix = r'$$K=\left​f_{x}00​0f_{y}0​c_{x}c_{y}1​\right​$$'
    valid_k_matrix = r'$$K = \begin{pmatrix} f_x & 0 & c_x \\ 0 & f_y & c_y \\ 0 & 0 & 1 \end{pmatrix}$$'
    if invalid_k_matrix in content:
        content = content.replace(invalid_k_matrix, valid_k_matrix)
        print("    - Đã sửa lỗi công thức ma trận K không hợp lệ.")

    print("...Tiền xử lý hoàn tất.")
    return content


def convert_md_to_docx(md_file_path, docx_file_path):
    """
    Đọc file Markdown, tiền xử lý nội dung để sửa lỗi, sau đó chuyển đổi
    sang DOCX với công thức toán học.
    """
    print(f"▶️  Bắt đầu chuyển đổi file: {md_file_path}")
    
    try:
        with open(md_file_path, 'r', encoding='utf-8') as f:
            md_content = f.read()
    except Exception as e:
        print(f"❌ Lỗi nghiêm trọng khi đọc file Markdown: {e}", file=sys.stderr)
        return

    # Sử dụng hàm tiền xử lý phiên bản mới (v3)
    processed_content = preprocess_markdown_v3(md_content)

    input_format = 'markdown+tex_math_dollars'
    extra_args = ['--standalone', '--mathml']

    print("⚙️  Đang gọi Pandoc để thực hiện chuyển đổi...")
    print(f"    - Định dạng đầu vào: {input_format}")
    print(f"    - Đối số thêm: {extra_args}")

    try:
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
    input_file = f'document\Kỹ thuật lập trình\Gemini-Giáo Trình Java_ Nhập Môn_6_7.md'
    output_file = f'document\Kỹ thuật lập trình\Gemini-Giáo Trình Java_ Nhập Môn_6_7_Output_v1.docx'

    # --- BẮT ĐẦU QUÁ TRÌNH XỬ LÝ ---
    print("\n--- CHƯƠNG TRÌNH CHUYỂN ĐỔI MARKDOWN SANG DOCX (v3 - Sửa lỗi hệ thống) ---")
    if not os.path.exists(input_file):
        print(f"Lỗi: File đầu vào '{input_file}' không tồn tại.", file=sys.stderr)
    else:
        convert_md_to_docx(input_file, output_file)
    print("-------------------------------------------------------------------------")
