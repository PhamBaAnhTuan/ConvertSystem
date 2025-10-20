import os
import re

# --- CẤU HÌNH ---
# !!! QUAN TRỌNG: Dán đường dẫn đến thư mục chứa ảnh của bạn vào đây !!!
TARGET_DIRECTORY = r'D:\DongAUniversity\TÀI LIỆU DẠY HỌC_2024-2025\Kỹ thuật lập trình\Các buổi học\CHƯƠNG 5 - LUỒNG VÀO-RA CƠ BẢN (BASIC I-O)\Thực hành Chương 5\slide\images'

# --- BẮT ĐẦU XỬ LÝ ---

def rename_slide_images():
    """
    Tự động tìm và đổi tên các file ảnh slide trong thư mục chỉ định.
    Ví dụ: "1_ABC.png" sẽ được đổi thành "slide-1.png"
    """
    print(f"🔍 Bắt đầu quét thư mục: {TARGET_DIRECTORY}")
    
    if not os.path.isdir(TARGET_DIRECTORY):
        print(f"❌ Lỗi: Thư mục không tồn tại. Vui lòng kiểm tra lại đường dẫn.")
        return

    renamed_count = 0
    skipped_count = 0
    
    # Lấy danh sách tất cả các file trong thư mục
    try:
        filenames = os.listdir(TARGET_DIRECTORY)
    except OSError as e:
        print(f"❌ Lỗi: Không thể truy cập thư mục. Chi tiết: {e}")
        return

    for filename in filenames:
        # Tìm các file có dạng "số_tênfile.đuôi"
        match = re.match(r'^(\d+)_.*(\.\w+)$', filename)
        
        if match:
            # Tách số thứ tự và phần đuôi file
            number = match.group(1)
            extension = match.group(2)
            
            # Tạo tên file mới
            new_filename = f"slide-{number}{extension}"
            
            # Lấy đường dẫn đầy đủ của file cũ và file mới
            old_path = os.path.join(TARGET_DIRECTORY, filename)
            new_path = os.path.join(TARGET_DIRECTORY, new_filename)
            
            # Thực hiện đổi tên
            try:
                os.rename(old_path, new_path)
                print(f"✅ Đã đổi tên: '{filename}'  ->  '{new_filename}'")
                renamed_count += 1
            except OSError as e:
                print(f"❌ Lỗi khi đổi tên file '{filename}': {e}")
                skipped_count += 1
        else:
            skipped_count += 1

    print("-" * 40)
    print("🎉 Hoàn tất!")
    print(f"👍 Đã đổi tên thành công: {renamed_count} file.")
    print(f"⏩ Đã bỏ qua: {skipped_count} file (không khớp định dạng).")


# Chạy hàm chính
if __name__ == "__main__":
    rename_slide_images()