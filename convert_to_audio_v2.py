import os
import re
import asyncio
import edge_tts

# --- Cấu hình ---
INPUT_FILE = r'D:\DongAUniversity\TÀI LIỆU DẠY HỌC_2024-2025\Máy học_UDA\Các buổi học\Chương 3 - Hồi quy – Regression\slide_video\script1.txt' 
OUTPUT_FOLDER = os.path.dirname(INPUT_FILE)
VOICE = 'vi-VN-HoaiMyNeural'

# TỐC ĐỘ NÓI: "+0%" là bình thường. "+20%" là nhanh hơn 20%. "-10%" là chậm hơn 10%.
# Giá trị "+15%" là một khởi đầu tốt để nói nhanh hơn một chút.
RATE = "+15%" 

# --- Bắt đầu xử lý ---

async def main():
    print(f"Bắt đầu quá trình chuyển đổi với edge-tts...")
    print(f"Tốc độ nói được đặt thành: {RATE}")
    print(f"Các file audio sẽ được lưu tại: {OUTPUT_FOLDER}")

    try:
        with open(INPUT_FILE, 'r', encoding='utf-8') as f:
            content = f.read()

        # [CẢI TIẾN] Regex linh hoạt hơn để xử lý các loại dấu gạch ngang và khoảng trắng khác nhau
        # Nó sẽ tìm các dòng có dạng --- SLIDE 1 ---, -- SLIDE 2 --, v.v.
        slides_content = re.split(r'\n-+[–—]\s*SLIDE\s*\d+\s*[–—]-+\n', content)

        slide_number = 1
        for text_chunk in slides_content:
            # Bỏ qua các đoạn text trống và chỉ xử lý nội dung slide thực sự
            cleaned_text = text_chunk.strip()
            # [CẢI TIẾN] Loại bỏ dòng tiêu đề nếu nó còn sót lại trong nội dung
            # Ví dụ: "Hiểu Rõ “Bên Dưới” Mô Hình\nTại sao chúng ta phải..."
            lines = cleaned_text.split('\n')
            if len(lines) > 1:
                 # Giả định dòng đầu tiên có thể là tiêu đề, các dòng sau là script
                 script_to_read = '\n'.join(lines) 
            else:
                 script_to_read = cleaned_text

            if script_to_read:
                print(f"Đang xử lý Slide {slide_number}...")
                
                output_filename = os.path.join(OUTPUT_FOLDER, f'slide_{slide_number}.mp3')
                
                # [CẢI TIẾN] Thêm tham số 'rate' để điều chỉnh tốc độ nói
                communicate = edge_tts.Communicate(script_to_read, VOICE, rate=RATE)
                await communicate.save(output_filename)
                
                print(f"-> Đã tạo thành công file: {output_filename}")
                slide_number += 1

        print("\nHoàn tất!")

    except FileNotFoundError:
        print(f"Lỗi: Không tìm thấy file '{INPUT_FILE}'. Vui lòng kiểm tra lại đường dẫn.")
    except Exception as e:
        print(f"Đã xảy ra lỗi không mong muốn: {e}")

if __name__ == "__main__":
    asyncio.run(main())