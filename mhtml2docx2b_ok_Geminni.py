import email
import base64
import os
import pypandoc
from bs4 import BeautifulSoup
import uuid
import shutil
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from PIL import Image
import io # Cần import thư viện io

# --- Các hàm get_katex_html và katex_to_image giữ nguyên như trước ---

def get_katex_html(latex_string):
    """Tạo một file HTML đầy đủ để render một công thức KaTeX."""
    # Thêm ký tự $$ để KaTeX hiểu là chế độ hiển thị (display mode)
    # Điều này giúp công thức lớn và rõ ràng hơn trong ảnh
    display_latex = f"$${latex_string}$$"
    return f"""
    <!DOCTYPE html>
    <html>
    <head>
        <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/katex@0.16.9/dist/katex.min.css">
        <script defer src="https://cdn.jsdelivr.net/npm/katex@0.16.9/dist/katex.min.js"></script>
        <style>
            body {{ margin: 0; padding: 5px; display: inline-block; background-color: white; }}
        </style>
    </head>
    <body>
        <span id="formula">{display_latex}</span>
        <script>
            document.addEventListener("DOMContentLoaded", function() {{
                katex.render(document.getElementById('formula').textContent, document.getElementById('formula'), {{
                    throwOnError: false,
                    displayMode: true
                }});
            }});
        </script>
    </body>
    </html>
    """

def katex_to_image(latex_string, output_path, temp_dir):
    """Render một chuỗi LaTeX thành hình ảnh sử dụng Selenium và KaTeX."""
    html_content = get_katex_html(latex_string)
    # Sử dụng một tên file html tạm thời duy nhất để tránh xung đột
    temp_html_path = os.path.join(temp_dir, f'temp_formula_{uuid.uuid4()}.html')

    with open(temp_html_path, 'w', encoding='utf-8') as f:
        f.write(html_content)

    options = webdriver.ChromeOptions()
    options.add_argument('--headless')
    options.add_argument('--disable-gpu')
    options.add_argument('--no-sandbox')
    # Thêm các tùy chọn để đảm bảo trình duyệt chạy ổn định
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument("window-size=1200,800")

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    
    try:
        driver.get(f'file://{os.path.abspath(temp_html_path)}')
        # Có thể cần tăng thời gian chờ nếu công thức phức tạp
        time.sleep(1.5) 
        
        # Chụp ảnh phần tử chứa công thức đã render
        element = driver.find_element(By.CLASS_NAME, 'katex-html')
        
        # Chụp ảnh màn hình của riêng element đó
        png = element.screenshot_as_png
        
        im = Image.open(io.BytesIO(png))
        im.save(output_path)
        
    finally:
        driver.quit()
        # Xóa file html tạm sau khi dùng xong
        if os.path.exists(temp_html_path):
            os.remove(temp_html_path)

def convert_mhtml_to_docx_universal(mhtml_file_path, docx_file_path):
    """
    Chuyển đổi file MHTML sang DOCX, xử lý KaTeX bằng cách chuyển thành hình ảnh.
    """
    print(f"🚀 Bắt đầu xử lý file: {mhtml_file_path}")
    
    temp_resource_dir = "temp_resources_" + str(uuid.uuid4())
    os.makedirs(temp_resource_dir, exist_ok=True)
    
    try:
        # Đọc file dưới dạng nhị phân để tránh lỗi encoding
        with open(mhtml_file_path, 'rb') as f:
            msg = email.message_from_binary_file(f)

        html_part = None
        for part in msg.walk():
            if part.get_content_type() == 'text/html':
                html_part = part
                break

        if not html_part:
            print("❌ Không tìm thấy phần HTML trong file MHTML.")
            return

        html_content = html_part.get_payload(decode=True).decode(
            html_part.get_content_charset() or 'utf-8', 'ignore'
        )
        soup = BeautifulSoup(html_content, 'html.parser')

        print("🔍 Đang tìm và chuyển đổi các công thức KaTeX...")
        
        # --- THAY ĐỔI QUAN TRỌNG ---
        # Tìm tất cả các thẻ 'span' có class là 'katex' hoặc 'katex-display'
        katex_elements = soup.find_all('span', class_=['katex', 'katex-display'])
        
        if not katex_elements:
            print("ℹ️ Không tìm thấy công thức KaTeX nào.")
        
        for i, element in enumerate(katex_elements):
            # Tránh xử lý các thẻ lồng nhau
            if element.find_parent(class_=['katex', 'katex-display']):
                continue

            latex_source = element.find('annotation', encoding='application/x-tex')
            if latex_source:
                latex_string = latex_source.get_text()
                # Tạo tên file ảnh duy nhất
                img_name = f'formula_{uuid.uuid4()}.png'
                img_path = os.path.join(temp_resource_dir, img_name)
                
                print(f"  - ⚙️ Đang render công thức {i+1}/{len(katex_elements)}...")
                try:
                    katex_to_image(latex_string, img_path, temp_resource_dir)
                    
                    if os.path.exists(img_path):
                        # Thay thế thẻ span bằng thẻ img
                        # Sử dụng đường dẫn tuyệt đối để pandoc có thể tìm thấy file
                        img_tag = soup.new_tag('img', src=os.path.abspath(img_path), style="vertical-align: middle;")
                        element.replace_with(img_tag)
                        print(f"    ✅ Đã chuyển công thức thành ảnh: {img_name}")
                    else:
                        print(f"    ⚠️ Không thể tạo file ảnh cho công thức.")

                except Exception as e:
                    print(f"    ❌ Lỗi khi render công thức: {latex_string[:30]}... | Lỗi: {e}")

        modified_html_path = os.path.join(temp_resource_dir, 'modified.html')
        with open(modified_html_path, 'w', encoding='utf-8') as f:
            f.write(str(soup))
        
        print("🔄 Đang chuyển đổi HTML sang DOCX với Pandoc...")
        output = pypandoc.convert_file(
            modified_html_path,
            'docx',
            format='html',
            outputfile=docx_file_path,
            extra_args=['--standalone', '--reference-doc=reference.docx']
        )

        if output == "":
            print(f"🎉 Chuyển đổi thành công! File DOCX đã được lưu tại: {docx_file_path}")
        else:
            print("❗️ Có lỗi xảy ra trong quá trình chuyển đổi Pandoc.")

    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        # Dọn dẹp thư mục tạm
        if os.path.exists(temp_resource_dir):
            shutil.rmtree(temp_resource_dir)
            print(f"🗑️ Đã xóa thư mục tạm: {temp_resource_dir}")

# --- Cách sử dụng ---
if __name__ == '__main__':
    # Đảm bảo file MHTML/TXT nằm cùng thư mục hoặc cung cấp đường dẫn đầy đủ
    mhtml_file = 'document/ComputerVision/Thị giác máy tính.mhtml'
    docx_file = 'document/ComputerVision/Thị giác máy tính_output2.docx'

    if os.path.exists(mhtml_file):
        # Tạo một file reference.docx trống để Pandoc sử dụng (tùy chọn nhưng nên có)
        if not os.path.exists('reference.docx'):
            print("Tạo file reference.docx mặc định...")
            os.system('pandoc -o reference.docx --print-default-data-file reference.docx')
            
        convert_mhtml_to_docx_universal(mhtml_file, docx_file)
    else:
        print(f"Lỗi: Không tìm thấy file '{mhtml_file}'. Vui lòng kiểm tra lại đường dẫn.")

