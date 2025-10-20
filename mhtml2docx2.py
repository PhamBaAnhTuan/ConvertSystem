import email
import base64
import os
import pypandoc
from bs4 import BeautifulSoup
import uuid
import shutil

def convert_mhtml_to_docx_large(mhtml_file_path, docx_file_path):
    """
    Chuyển đổi file MHTML lớn sang DOCX bằng cách sử dụng file tạm
    để tránh các vấn đề về bộ nhớ.
    """
    print(f"Bắt đầu xử lý file lớn: {mhtml_file_path}")
    
    temp_resource_dir = "temp_resources_" + str(uuid.uuid4())
    os.makedirs(temp_resource_dir, exist_ok=True)
    print(f"Đã tạo thư mục tài nguyên tạm: {temp_resource_dir}")

    try:
        with open(mhtml_file_path, 'rb') as f:
            msg = email.message_from_bytes(f.read())
    except Exception as e:
        print(f"Lỗi nghiêm trọng khi đọc file MHTML: {e}")
        shutil.rmtree(temp_resource_dir)
        return

    html_part = None
    resources = {}

    # --- SỬA ĐỔI QUAN TRỌNG BẮT ĐẦU TỪ ĐÂY ---
    # Tìm phần HTML chính một cách linh hoạt hơn
    for part in msg.walk():
        content_type = part.get_content_type()
        content_disposition = str(part.get("Content-Disposition"))

        # Điều kiện mới: Kiểm tra content_type có tồn tại và BẮT ĐẦU BẰNG 'text/html'
        # Đồng thời kiểm tra content_disposition không phải là 'attachment'
        if content_type and content_type.startswith('text/html') and 'attachment' not in content_disposition:
            html_part = part
            # Khi đã tìm thấy, chúng ta có thể dừng sớm để tránh ghi đè bởi một phần khác (nếu có)
            # break # Tạm thời bình luận dòng này để đảm bảo nó lấy phần cuối cùng, thường là phần đúng.
        
        # Lưu các tài nguyên
        if part.get('Content-ID'):
            cid = part.get('Content-ID').strip('<>')
            resources[cid] = part
    # --- KẾT THÚC SỬA ĐỔI QUAN TRỌNG ---

    if not html_part:
        print("❌ Lỗi: Không tìm thấy nội dung HTML trong file MHTML ngay cả khi đã dùng điều kiện linh hoạt hơn.")
        shutil.rmtree(temp_resource_dir)
        # THÊM BƯỚC GỠ LỖI: In ra cấu trúc file nếu không tìm thấy
        print("--- BẮT ĐẦU IN CẤU TRÚC FILE ĐỂ GỠ LỖI ---")
        i = 0
        with open(mhtml_file_path, 'rb') as f:
            msg_debug = email.message_from_bytes(f.read())
        for part_debug in msg_debug.walk():
            i += 1
            print(f"\n--- Phần #{i} ---")
            print(f"  Content-Type: {part_debug.get_content_type()}")
            print(f"  Content-Disposition: {part_debug.get('Content-Disposition')}")
            print(f"  Content-ID: {part_debug.get('Content-ID')}")
        print("--- KẾT THÚC IN CẤU TRÚC FILE ---")
        return

    charset = html_part.get_content_charset() or 'utf-8'
    html_content = html_part.get_payload(decode=True).decode(charset, 'replace')
    soup = BeautifulSoup(html_content, 'lxml')

    for img_tag in soup.find_all('img'):
        src = img_tag.get('src')
        if src and src.startswith('cid:'):
            cid_key = src[4:]
            if cid_key in resources:
                resource_part = resources[cid_key]
                filename = resource_part.get_filename() or f"{cid_key.split('@')[0]}.png"
                filepath = os.path.join(temp_resource_dir, filename)
                
                with open(filepath, 'wb') as f:
                    f.write(resource_part.get_payload(decode=True))
                
                img_tag['src'] = filepath
                print(f"Đã trích xuất và lưu tài nguyên: {filepath}")

    temp_html_path = 'temp_main.html'
    with open(temp_html_path, 'w', encoding='utf-8') as f:
        f.write(str(soup))
    
    print(f"File HTML tạm đã được lưu tại: {temp_html_path}")
    print("Đang gọi Pandoc để chuyển đổi từ file...")
    
    try:
        pypandoc.convert_file(
            temp_html_path,
            'docx',
            outputfile=docx_file_path,
            extra_args=['--standalone']
        )
        print(f"✅ Thành công! File DOCX đã được lưu tại: {docx_file_path}")
    except Exception as e:
        print(f"❌ Đã xảy ra lỗi trong quá trình chuyển đổi bằng Pandoc: {e}")

    finally:
        print("Đang dọn dẹp các file tạm...")
        if os.path.exists(temp_html_path):
            os.remove(temp_html_path)
        if os.path.exists(temp_resource_dir):
            shutil.rmtree(temp_resource_dir)
        print("Dọn dẹp hoàn tất.")

if __name__ == '__main__':
    input_mhtml_file = 'Gemini.mhtml' 
    output_docx_file = 'Gemini_output2a.docx'   

    if not os.path.exists(input_mhtml_file):
        print(f"Lỗi: File '{input_mhtml_file}' không tồn tại.")
    else:
        convert_mhtml_to_docx_large(input_mhtml_file, output_docx_file)