import os
import re
from gtts import gTTS

# --- Cấu hình ---
INPUT_FILE = r'D:\DongAUniversity\TÀI LIỆU DẠY HỌC_2024-2025\Máy học_UDA\Các buổi học\Chương 3 - Hồi quy – Regression\slide_video\script1.txt'
# Tự động lấy đường dẫn thư mục từ file input
OUTPUT_FOLDER = os.path.dirname(INPUT_FILE) 
LANGUAGE = 'vi' # Tiếng Việt

# --- Bắt đầu xử lý ---
print(f"Bắt đầu quá trình chuyển đổi với gTTS...")
print(f"Các file audio sẽ được lưu tại: {OUTPUT_FOLDER}")

try:
    with open(INPUT_FILE, 'r', encoding='utf-8') as f:
        content = f.read()

    slides_content = re.split(r'\n-– SLIDE \d+ -–\n', content)
    slide_number = 1
    for text_chunk in slides_content:
        cleaned_text = text_chunk.strip()
        if cleaned_text:
            print(f"Đang xử lý Slide {slide_number}...")
            tts = gTTS(text=cleaned_text, lang=LANGUAGE, slow=False)
            output_filename = os.path.join(OUTPUT_FOLDER, f'slide_{slide_number}.mp3')
            tts.save(output_filename)
            print(f"-> Đã tạo thành công file: {output_filename}")
            slide_number += 1

    print("\nHoàn tất!")

except FileNotFoundError:
    print(f"Lỗi: Không tìm thấy file '{INPUT_FILE}'. Vui lòng kiểm tra lại đường dẫn.")
except Exception as e:
    print(f"Đã xảy ra lỗi không mong muốn: {e}")