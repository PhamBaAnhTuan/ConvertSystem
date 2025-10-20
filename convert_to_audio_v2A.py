import os
import re
import asyncio
import edge_tts

# --- CẤU HÌNH ---
# !!! QUAN TRỌNG: Hãy thay đổi đường dẫn này đến đúng file script của bạn !!!
INPUT_FILE = r'D:\DongAUniversity\TÀI LIỆU DẠY HỌC_2024-2025\Máy học_UDA\Các buổi học\Chương 3 - Hồi quy – Regression\slide_video\script1.txt' 

# [CẬP NHẬT] Tạo một thư mục con tên là "output1" để chứa file audio
parent_directory = os.path.dirname(INPUT_FILE)
OUTPUT_FOLDER = os.path.join(parent_directory, "output1") # <-- THAY ĐỔI TẠI ĐÂY

VOICE = 'vi-VN-HoaiMyNeural' # Giọng đọc nữ Miền Nam
RATE = "+15%" # Tốc độ nói
MAX_RETRIES = 3 # Số lần thử lại tối đa cho mỗi slide nếu gặp lỗi

# --- BẮT ĐẦU XỬ LÝ ---

async def main():
    """
    Hàm chính để thực hiện việc chuyển đổi văn bản thành âm thanh.
    """
    print(f"▶️  Bắt đầu quá trình chuyển đổi với Edge-TTS (phiên bản v4)...")
    
    # [CẬP NHẬT] Tạo thư mục output nếu nó chưa tồn tại
    if not os.path.exists(OUTPUT_FOLDER):
        os.makedirs(OUTPUT_FOLDER)
        print(f"✅ Đã tạo thư mục mới tại: {OUTPUT_FOLDER}")

    print(f"🗣️  Tốc độ nói được đặt thành: {RATE}")
    print(f"🔁 Số lần thử lại tối đa khi gặp lỗi: {MAX_RETRIES}")
    print(f"📁 Các file audio sẽ được lưu tại: {OUTPUT_FOLDER}")
    print("-" * 40)

    try:
        with open(INPUT_FILE, 'r', encoding='utf-8') as f:
            content = f.read()

        slides_content = re.split(r'\n[-–—]+\s*SLIDE\s*\d+\s*[-–—]+\n', content)

        slide_number = 1
        processed_count = 0
        failed_slides = []

        for text_chunk in slides_content:
            cleaned_text = text_chunk.strip()
            if not cleaned_text:
                continue

            script_to_read = cleaned_text
            lines = cleaned_text.split('\n')
            if len(lines) > 1 and len(lines[0]) < 70 and not lines[0].strip().endswith(('.', '!', '?', ':', ';', ',')):
                script_to_read = '\n'.join(lines[1:]).strip()
            
            print(f"▶️  Đang xử lý Slide {slide_number}...")
            
            output_filename = os.path.join(OUTPUT_FOLDER, f'slide_{slide_number}.mp3')
            
            success = False
            for attempt in range(MAX_RETRIES):
                try:
                    communicate = edge_tts.Communicate(script_to_read, VOICE, rate=RATE)
                    await communicate.save(output_filename)
                    print(f"  ✅ Thành công sau {attempt + 1} lần thử.")
                    success = True
                    processed_count += 1
                    break
                except Exception as e:
                    print(f"  ⚠️ Lần thử {attempt + 1}/{MAX_RETRIES} thất bại: {str(e).strip()}")
                    if attempt < MAX_RETRIES - 1:
                        print("     -> Đang chờ 2 giây để thử lại...")
                        await asyncio.sleep(2)
            
            if not success:
                print(f"  ❌ Xử lý Slide {slide_number} thất bại sau {MAX_RETRIES} lần thử.")
                failed_slides.append(slide_number)

            slide_number += 1

        print("-" * 40)
        print(f"🎉 Hoàn tất! Đã xử lý và tạo ra thành công {processed_count} file audio.")
        if failed_slides:
            print(f"⚠️ Các slide sau đã bị lỗi: {failed_slides}")

    except FileNotFoundError:
        print(f"❌ Lỗi: Không tìm thấy file '{INPUT_FILE}'. Vui lòng kiểm tra lại đường dẫn.")
    except Exception as e:
        print(f"❌ Đã xảy ra lỗi không mong muốn trong quá trình đọc file: {e}")

if __name__ == "__main__":
    asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())
    asyncio.run(main())