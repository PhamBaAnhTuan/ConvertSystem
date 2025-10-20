import email
import base64
import os
import pypandoc
from bs4 import BeautifulSoup

def convert_mhtml_to_docx(mhtml_file_path, docx_file_path):
    """
    Chuyển đổi một file MHTML sang DOCX, cố gắng giữ lại công thức toán học.
    Sử dụng Pandoc làm công cụ chuyển đổi chính.

    Args:
        mhtml_file_path (str): Đường dẫn đến file MHTML đầu vào.
        docx_file_path (str): Đường dẫn để lưu file DOCX đầu ra.
    """
    print(f"Bắt đầu xử lý file: {mhtml_file_path}")

    try:
        # Mở file MHTML bằng encoding utf-8 là phổ biến nhất
        with open(mhtml_file_path, 'r', encoding='utf-8') as f:
            msg = email.message_from_file(f)
    except Exception:
        # Nếu lỗi, thử mở ở dạng binary để thư viện email tự xử lý
        print("Không đọc được với UTF-8, thử đọc file ở dạng binary...")
        try:
             with open(mhtml_file_path, 'rb') as f:
                msg = email.message_from_bytes(f.read())
        except Exception as e_inner:
            print(f"Lỗi nghiêm trọng: Không thể đọc file MHTML. Chi tiết: {e_inner}")
            return

    html_part = None
    resources = {}

    # Phân tích file MHTML để lấy ra phần HTML và các tài nguyên (ảnh, css...)
    for part in msg.walk():
        content_type = part.get_content_type()
        
        # Tìm phần nội dung HTML chính
        if content_type == 'text/html' and "attachment" not in part.get("Content-Disposition", ""):
            html_part = part
        
        # Lưu các tài nguyên khác vào một dictionary với key là Content-ID (cid)
        if part.get('Content-ID'):
            cid = part.get('Content-ID').strip('<>')
            resources[cid] = part

    if not html_part:
        print("Lỗi: Không tìm thấy nội dung HTML trong file MHTML.")
        return

    # Lấy nội dung HTML và decode
    charset = html_part.get_content_charset() or 'utf-8'
    html_content = html_part.get_payload(decode=True).decode(charset, 'replace')
    
    # Sử dụng BeautifulSoup để xử lý HTML
    soup = BeautifulSoup(html_content, 'lxml')

    # Thay thế các link tài nguyên cid: bằng dữ liệu base64 nhúng trực tiếp
    # Điều này giúp Pandoc có thể tìm thấy và nhúng hình ảnh vào file DOCX
    for img_tag in soup.find_all('img'):
        src = img_tag.get('src')
        if src and src.startswith('cid:'):
            cid_key = src[4:]
            if cid_key in resources:
                resource_part = resources[cid_key]
                mime_type = resource_part.get_content_type()
                img_data = resource_part.get_payload(decode=True)
                base64_data = base64.b64encode(img_data).decode('utf-8')
                
                # Tạo data URI và thay thế src của thẻ img
                data_uri = f"data:{mime_type};base64,{base64_data}"
                img_tag['src'] = data_uri
                print(f"Đã xử lý và nhúng tài nguyên: {cid_key}")

    # Lấy lại chuỗi HTML đã được cập nhật
    final_html = str(soup)

    print("Đang gọi Pandoc để chuyển đổi...")
    
    try:
        # Sử dụng pypandoc để chuyển đổi HTML sang DOCX
        # Pandoc sẽ tự động nhận diện MathML/LaTeX và chuyển thành công thức Word
        pypandoc.convert_text(
            final_html,
            'docx',
            format='html+smart', # '+smart' để xử lý các kí tự đặc biệt
            outputfile=docx_file_path,
            extra_args=['--standalone']
        )
        print(f"✅ Thành công! File DOCX đã được lưu tại: {docx_file_path}")
    except Exception as e:
        print(f"❌ Đã xảy ra lỗi trong quá trình chuyển đổi bằng Pandoc: {e}")
        print("👉 Gợi ý: Hãy chắc chắn rằng bạn đã cài đặt Pandoc và thêm nó vào PATH hệ thống.")
        print("   Bạn có thể kiểm tra bằng cách mở terminal và gõ 'pandoc --version'")

# --- CÁCH SỬ DỤNG ---
if __name__ == '__main__':
    # 1. Đặt file MHTML của bạn vào cùng thư mục với script này
    # 2. Thay đổi tên file dưới đây cho đúng với file của bạn
    input_mhtml_file = 'Gemini.mhtml' 
    output_docx_file = 'Gemini_output.docx'   

    # Kiểm tra xem file input có tồn tại không
    if not os.path.exists(input_mhtml_file):
        print(f"Lỗi: File '{input_mhtml_file}' không tồn tại.")
        # Nếu không có, tạo 1 file mẫu để người dùng thử nghiệm
        print("Đang tạo một file MHTML mẫu 'example.mhtml'...")
        sample_mhtml_content = r"""MIME-Version: 1.0
Content-Type: multipart/related; boundary="----=_NextPart_Sample"

------=_NextPart_Sample
Content-Type: text/html; charset="utf-8"
Content-Transfer-Encoding: quoted-printable

<!DOCTYPE html><html><head><title>Test Math</title><script type=3D"text/ja=
vascript" async src=3D"https://cdnjs.cloudflare.com/ajax/libs/mathjax/2.7.=
7/MathJax.js?config=3DTex-MML-AM_CHTML"></script></head><body>
<h1>Test chuyển đổi công thức toán</h1>
<p>Công thức inline: \(E =3D mc^2\).</p>
<p>Phương trình dạng khối:</p>
$$ \int_a^b f(x) \, dx =3D F(b) - F(a) $$
<p>Phương trình bậc hai:</p>
$$ x =3D \frac{-b \pm \sqrt{b^2 - 4ac}}{2a} $$
</body></html>
------=_NextPart_Sample--
"""
        with open(input_mhtml_file, 'w', encoding='utf-8') as f:
            f.write(sample_mhtml_content)
        print("File 'example.mhtml' đã được tạo. Hãy chạy lại script để chuyển đổi nó.")
    else:
        convert_mhtml_to_docx(input_mhtml_file, output_docx_file)