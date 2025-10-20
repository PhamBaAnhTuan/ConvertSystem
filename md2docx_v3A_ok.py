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
    if '$$$' in content:
        content = content.replace('$$$', '$$')
        print("    - Đã chuẩn hóa '$$$' thành '$$'.")

    # 2. Sửa lỗi dư dấu gạch chéo ngược cho các lệnh LaTeX
    original_pattern_count = len(re.findall(r'\\\\([a-zA-Z]+)', content))
    if original_pattern_count > 0:
        content = re.sub(r'\\\\([a-zA-Z]+)', r'\\\1', content)
        print(f"    - Đã sửa {original_pattern_count} lỗi lệnh LaTeX dư dấu '\\' (ví dụ: \\sum, \\lambda, \\cdot).")
        
    # 3. Sửa lỗi dư dấu gạch chéo ngược cho ký tự gạch dưới
    original_underscore_count = content.count(r'\_')
    if original_underscore_count > 0:
        content = content.replace(r'\_', '_')
        print(f"    - Đã sửa {original_underscore_count} lỗi ký tự gạch dưới '\\_'.")

    print("...Tiền xử lý hoàn tất.")
    return content


def convert_md_to_docx(md_file_path, docx_file_path):
    """
    Đọc file Markdown, tiền xử lý nội dung để sửa lỗi, sau đó chuyển đổi
    sang DOCX với công thức toán học và màu sắc cú pháp cho code.
    """
    print(f"▶️  Bắt đầu chuyển đổi file: {md_file_path}")
    
    try:
        with open(md_file_path, 'r', encoding='utf-8') as f:
            md_content = f.read()
    except Exception as e:
        print(f"❌ Lỗi nghiêm trọng khi đọc file Markdown: {e}", file=sys.stderr)
        return

    processed_content = preprocess_markdown_v3(md_content)

    input_format = 'markdown+tex_math_dollars'
    
    # --- THAY ĐỔI QUAN TRỌNG Ở ĐÂY ---
    # Thêm '--highlight-style=tango' để Pandoc tự động tô màu cho các khối mã.
    # Bạn có thể thử các kiểu khác như: pygments, kate, espresso, zenburn...
    extra_args = ['--standalone', '--mathml', '--highlight-style=tango'] 

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
    # Hãy đảm bảo file Markdown đầu vào đã được định dạng với khối mã ```java
    input_file = r'C:\Users\thanh\Downloads\Gemini-Hướng Dẫn Thanh Toán Thương Mại Điện Tử (1).md'
    #output_file = f'document\Kỹ thuật lập trình\Gemini-Giáo Trình Java_ Nhập Môn_all_Output_v1A.docx'
    # Tạo tên file output .docx cùng tên và cùng thư mục với input_file
    base_name = os.path.splitext(os.path.basename(input_file))[0]
    folder = os.path.dirname(input_file)
    output_file = os.path.join(folder, f"{base_name}.docx")
    # --- BẮT ĐẦU QUÁ TRÌNH XỬ LÝ ---
    print("\n--- CHƯƠNG TRÌNH CHUYỂN ĐỔI MARKDOWN SANG DOCX (v4 - Hỗ trợ màu sắc) ---")
    if not os.path.exists(input_file):
        print(f"Lỗi: File đầu vào '{input_file}' không tồn tại.", file=sys.stderr)
    else:
        convert_md_to_docx(input_file, output_file)
    print("-------------------------------------------------------------------------")
    
