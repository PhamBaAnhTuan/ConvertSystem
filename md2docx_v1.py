import os
import pypandoc
import sys

def convert_md_to_docx(md_file_path, docx_file_path):
    """
    Chuyển đổi file Markdown (.md) sang file DOCX (.docx), xử lý các công thức
    toán học được định dạng bằng LaTeX thành công thức gốc trong Word.

    Args:
        md_file_path (str): Đường dẫn đến file Markdown đầu vào.
        docx_file_path (str): Đường dẫn để lưu file DOCX đầu ra.
    """
    print(f"▶️  Bắt đầu chuyển đổi file: {md_file_path}")

    # --- Định dạng đầu vào và các đối số cho Pandoc ---
    # 'markdown+tex_math_dollars' cho phép Pandoc nhận diện công thức toán
    # được bao quanh bởi dấu $...$ và $$...$$
    input_format = 'markdown+tex_math_dollars'

    # --standalone: Tạo một file docx hoàn chỉnh.
    # --mathml: Sử dụng MathML để render công thức toán, cho chất lượng
    #           tốt nhất và có thể chỉnh sửa được trong Word.
    extra_args = ['--standalone', '--mathml']

    print("⚙️  Đang gọi Pandoc để thực hiện chuyển đổi...")
    print(f"    - Định dạng đầu vào: {input_format}")
    print(f"    - Đối số thêm: {extra_args}")

    try:
        # Thực hiện chuyển đổi
        pypandoc.convert_file(
            source_file=md_file_path,
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

    # Đường dẫn đến file Markdown bạn muốn chuyển đổi
    # (Trong trường hợp này là file bạn đã cung cấp)
    input_file = 'document\ComputerVision\Gemini-Giáo Trình Thị Giác Máy Tính.md'

    # Tên file DOCX bạn muốn tạo ra
    output_file = 'document\ComputerVision\Gemini-Giáo Trình Thị Giác Máy Tính_Output.docx'

    # --- BẮT ĐẦU QUÁ TRÌNH XỬ LÝ ---
    print("\n--- CHƯƠNG TRÌNH CHUYỂN ĐỔI MARKDOWN SANG DOCX ---")
    if not os.path.exists(input_file):
        print(f"Lỗi: File đầu vào '{input_file}' không tồn tại.", file=sys.stderr)
    else:
        convert_md_to_docx(input_file, output_file)
    print("-------------------------------------------------")