import os
import re
import asyncio
import edge_tts

# --- CẤU HÌNH ---
# !!! QUAN TRỌNG: Hãy thay đổi đường dẫn này đến đúng file script của bạn !!!
INPUT_FILE = r'D:\DongAUniversity\TÀI LIỆU DẠY HỌC_2024-2025\Kỹ thuật lập trình\Các buổi học\CHƯƠNG 3 - KIỂU DỮ LIỆU, BIẾN VÀ MẢNG\Script Chương 3-Hướng dẫn lập trình Java cơ bản.txt' 

# Thư mục output để chứa file audio
parent_directory = os.path.dirname(INPUT_FILE)
OUTPUT_FOLDER = os.path.join(parent_directory, "audio_chat_luong_cao") # Lưu vào thư mục mới

# === [CẬP NHẬT] Các thông số âm thanh mới cho chất lượng cao hơn ===
VOICE = 'vi-VN-HoaiMyNeural' 
RATE = "+5%"      # Tốc độ nói nhanh hơn
PITCH = "+10Hz"     # Cao độ giọng cao hơn, trong hơn
VOLUME = "+100%"    # Âm lượng tối đa
# =================================================================

MAX_RETRIES = 3 # Số lần thử lại tối đa

# --- BẮT ĐẦU XỬ LÝ ---

async def main():
    """
    Hàm chính để thực hiện việc chuyển đổi văn bản thành âm thanh.
    """
    print(f"▶️  Bắt đầu quá trình tạo audio chất lượng cao...")
    
    if not os.path.exists(OUTPUT_FOLDER):
        os.makedirs(OUTPUT_FOLDER)
        print(f"✅ Đã tạo thư mục mới tại: {OUTPUT_FOLDER}")

    print(f"🗣️  Giọng đọc: {VOICE} | Tốc độ: {RATE} | Cao độ: {PITCH} | Âm lượng: {VOLUME}")
    print(f"📁 Các file audio sẽ được lưu tại: {OUTPUT_FOLDER}")
    print("-" * 40)

    try:
        with open(INPUT_FILE, 'r', encoding='utf-8') as f:
            content = f.read()

        #slides_content = re.split(r'\n[-–—]+\s*SLIDE\s*\d+\s*[-–—]+\n', content)
        slides_content = re.split(r'(?=Slide \d+:)', content)
        valid_slides = [chunk for chunk in slides_content if chunk.strip()]
        
        print(f"🔍 Đã tìm thấy {len(valid_slides)} slide có nội dung trong file script.")
        print("-" * 40)

        processed_count = 0
        failed_slides = []

        for i, text_chunk in enumerate(valid_slides):
            slide_number = i + 1
            script_to_read = text_chunk.strip()
            
            print(f"▶️  Đang xử lý Slide {slide_number}...")
            
            output_filename = os.path.join(OUTPUT_FOLDER, f'slide_{slide_number}.mp3')
            
            success = False
            for attempt in range(MAX_RETRIES):
                try:
                    # [CẬP NHẬT] Thêm các tham số pitch và volume vào đây
                    communicate = edge_tts.Communicate(
                        script_to_read, 
                        VOICE, 
                        rate=RATE, 
                        pitch=PITCH, 
                        volume=VOLUME
                    )
                    await communicate.save(output_filename)
                    print(f"  ✅ Thành công sau {attempt + 1} lần thử.")
                    success = True
                    processed_count += 1
                    break
                except Exception as e:
                    print(f"  ⚠️ Lần thử {attempt + 1}/{MAX_RETRIES} thất bại.")
                    if attempt < MAX_RETRIES - 1:
                        await asyncio.sleep(2)
            
            if not success:
                failed_slides.append(slide_number)

        print("-" * 40)
        print(f"🎉 Hoàn tất! Đã xử lý và tạo ra thành công {processed_count} file audio.")
        if failed_slides:
            print(f"⚠️ Các slide sau đã bị lỗi: {failed_slides}")

    except FileNotFoundError:
        print(f"❌ Lỗi: Không tìm thấy file '{INPUT_FILE}'.")
    except Exception as e:
        print(f"❌ Đã xảy ra lỗi không mong muốn: {e}")

if __name__ == "__main__":
    asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())
    asyncio.run(main())