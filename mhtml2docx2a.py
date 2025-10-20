import email
import base64
import os
import pypandoc
from bs4 import BeautifulSoup
import uuid
import shutil

def convert_mhtml_to_docx_final(mhtml_file_path, docx_file_path):
    """
    Chuyển đổi file MHTML lớn sang DOCX, sử dụng logic tìm kiếm phần HTML
    lớn nhất để đảm bảo độ chính xác.
    """
    print(f"Bắt đầu xử lý file: {mhtml_file_path}")
    
    temp_resource_dir = "temp_resources_" + str(uuid.uuid4())
    os.makedirs(temp_resource_dir, exist_ok=True)

    try:
        with open(mhtml_file_path, 'rb') as f:
            msg = email.message_from_bytes(f.read())
    except Exception as e:
        print(f"Lỗi nghiêm trọng khi đọc file MHTML: {e}")
        shutil.rmtree(temp_resource_dir)
        return

    html_parts_candidates = []
    resources = {}

    for part in msg.walk():
        content_type = part.get_content_type()
        content_disposition = str(part.get("Content-Disposition"))

        # 1. Thu thập tất cả các phần HTML tiềm năng vào một danh sách
        if content_type and content_type.startswith('text/html') and 'attachment' not in content_disposition:
            html_parts_candidates.append(part)
        
        if part.get('Content-ID'):
            cid = part.get('Content-ID').strip('<>')
            resources[cid] = part

    # 2. Chọn phần HTML lớn nhất từ danh sách ứng cử viên
    html_part = None
    if html_parts_candidates:
        # Dùng hàm max với key là độ dài của payload để tìm phần lớn nhất
        html_part = max(html_parts_candidates, key=lambda p: len(p.get_payload(decode=True)))
        print(f"✅ Đã tìm thấy {len(html_parts_candidates)} phần HTML. Đã chọn phần lớn nhất (kích thước: {len(html_part.get_payload(decode=True))} bytes).")
    
    if not html_part:
        print("❌ Lỗi: Không tìm thấy bất kỳ phần HTML hợp lệ nào trong file MHTML.")
        shutil.rmtree(temp_resource_dir)
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

    temp_html_path = 'temp_main_final.html'
    with open(temp_html_path, 'w', encoding='utf-8') as f:
        f.write(str(soup))
    
    print("Đang gọi Pandoc để chuyển đổi...")
    
    try:
        pypandoc.convert_file(
            temp_html_path,
            'docx',
            outputfile=docx_file_path,
            extra_args=['--standalone']
        )
        print(f"🏆 Thành công! File DOCX đã được lưu tại: {docx_file_path}")
    except Exception as e:
        print(f"❌ Đã xảy ra lỗi trong quá trình chuyển đổi bằng Pandoc: {e}")

    finally:
        # Dọn dẹp các file tạm sau khi hoàn tất
        print("Đang dọn dẹp các file tạm...")
        if os.path.exists(temp_html_path):
            os.remove(temp_html_path)
        if os.path.exists(temp_resource_dir):
            shutil.rmtree(temp_resource_dir)
        print("Dọn dẹp hoàn tất.")

if __name__ == '__main__':
    input_mhtml_file = 'gptb2.mhtml' 
    output_docx_file = 'gptb2_docx.docx'   

    if not os.path.exists(input_mhtml_file):
        print(f"Lỗi: File '{input_mhtml_file}' không tồn tại.")
    else:
        convert_mhtml_to_docx_final(input_mhtml_file, output_docx_file)