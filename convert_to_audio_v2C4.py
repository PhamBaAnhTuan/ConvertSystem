import os
import re
import asyncio
import edge_tts

# --- CẤU HÌNH ---
INPUT_FILE = r'D:\DongAUniversity\Khoa CNTT\AUN_28.6.2025\ThuyetTrinh_26.9.2025\slide_17.10.2025\Information Technology_17.10.2025_notes.txt'
parent_directory = os.path.dirname(INPUT_FILE)
OUTPUT_FOLDER = os.path.join(parent_directory, "audio")

# Các giọng nữ (female voice) phổ biến:
# vi-VN-HoaiMyNeural → nữ, tự nhiên, hơi trẻ trung (bạn đang dùng).
# vi-VN-HoaiAnNeural → nữ, giọng miền Bắc, trong trẻo.
# vi-VN-HoaiBaoNeural → nữ, giọng miền Trung, ấm hơn.
# vi-VN-HoaiNamNeural → nữ, nhẹ nhàng, hơi “formal”.
# vi-VN-HoaiSuongNeural → nữ, giọng miền Nam, mềm mại.

# Giọng nam (male voice) để tham khảo:
# vi-VN-NamMinhNeural → nam, chuẩn miền Bắc.
# vi-VN-NamQuanNeural → nam, miền Trung, chậm rãi.
# vi-VN-NamPhongNeural → nam, miền Nam, trẻ trung.

VOICE = 'vi-VN-HoaiMyNeural' #'vi-VN-NamQuanNeural' #
#VOICE  = 'en-US-GuyNeural'
#RATE = "+5%"
RATE = "0%"
PITCH = "+10Hz"
VOLUME = "+100%"


# --- VOICE CONFIG (English, Male) ---
# Chọn 1 trong 2:
# VOICE  = 'en-US-GuyNeural'   # English (US), male
# # VOICE  = 'en-GB-RyanNeural'  # English (UK), male

# RATE   = "+0%"
# PITCH  = "+10Hz"
# VOLUME = "+100%"                # tránh clip; có thể bỏ hẳn tham số volume

MAX_RETRIES = 3
BASE_RETRY_SLEEP = 2.0      # giây
BETWEEN_SLIDES_SLEEP = 0.8  # giây

def strip_bom(s: str) -> str:
    return s.lstrip("\ufeff")

def parse_slides(text: str):
    """
    Trả về list các dict: [{'index': 1, 'title': 'Slide 1: ...', 'body': '...'}, ...]
    Bỏ qua các slide có body rỗng.
    """
    text = strip_bom(text).replace("\r\n", "\n").replace("\r", "\n")
    # Bắt tiêu đề 'Slide X: ...' và lấy phần nội dung cho đến trước 'Slide Y:' tiếp theo
    pattern = r"(Slide\s+\d+:\s*[^\n]*)(?:\n)(.*?)(?=\nSlide\s+\d+:|\Z)"
    matches = re.findall(pattern, text, flags=re.S)
    slides = []
    for idx, (title, body) in enumerate(matches, start=1):
        body = body.strip()
        if not body:
            # skip slide rỗng
            continue
        slides.append({"index": idx, "title": title.strip(), "body": body})
    return slides

async def tts_one(text: str, outfile: str):
    comm = edge_tts.Communicate(
        text,
        VOICE,
        rate=RATE,
        pitch=PITCH,
        volume=VOLUME
    )
    await comm.save(outfile)

async def main():
    print("▶️  Bắt đầu quá trình tạo audio chất lượng cao...")

    os.makedirs(OUTPUT_FOLDER, exist_ok=True)
    print(f"🗣️  Giọng đọc: {VOICE} | Tốc độ: {RATE} | Cao độ: {PITCH} | Âm lượng: {VOLUME}")
    print(f"📁 Các file audio sẽ được lưu tại: {OUTPUT_FOLDER}")
    print("-" * 40)

    try:
        with open(INPUT_FILE, "r", encoding="utf-8") as f:
            raw = f.read()

        slides = parse_slides(raw)
        print(f"🔍 Đã tìm thấy {len(slides)} slide hợp lệ có nội dung trong file script.")
        print("-" * 40)

        processed = 0
        failed = []

        for s in slides:
            slide_no = s["index"]  # đánh số đúng với tiêu đề Slide X
            text_chunk = s["body"]

            print(f"▶️  Đang xử lý Slide {slide_no}...")

            # Tên file: slide_1.mp3, slide_2.mp3, ...
            outpath = os.path.join(OUTPUT_FOLDER, f"slide_{slide_no}.mp3")

            ok = False
            for attempt in range(1, MAX_RETRIES + 1):
                try:
                    await tts_one(text_chunk, outpath)
                    print(f"  ✅ Thành công sau {attempt} lần thử.")
                    ok = True
                    processed += 1
                    break
                except Exception as e:
                    print(f"  ⚠️ Lần thử {attempt}/{MAX_RETRIES} thất bại. Lý do: {repr(e)}")
                    if attempt < MAX_RETRIES:
                        # Exponential backoff
                        await asyncio.sleep(BASE_RETRY_SLEEP * (2 ** (attempt - 1)))

            if not ok:
                failed.append(slide_no)

            # Nghỉ ngắn giữa các slide để tránh rate-limit
            await asyncio.sleep(BETWEEN_SLIDES_SLEEP)

        print("-" * 40)
        print(f"🎉 Hoàn tất! Đã xử lý thành công {processed} file audio.")
        if failed:
            print(f"⚠️ Các slide sau bị lỗi: {failed}")

    except FileNotFoundError:
        print(f"❌ Lỗi: Không tìm thấy file '{INPUT_FILE}'.")
    except Exception as e:
        print(f"❌ Đã xảy ra lỗi không mong muốn: {repr(e)}")

if __name__ == "__main__":
    asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())
    asyncio.run(main())
