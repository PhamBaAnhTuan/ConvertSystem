import cv2
import os
import shutil
from skimage.metrics import structural_similarity as ssim
import numpy as np
import time

def extract_slides_optimized(video_path, output_folder, threshold=0.85):
    """
    phiên bản tối ưu của hàm trích xuất slide.
    - Xóa thư mục đầu ra cũ nếu tồn tại để đảm bảo kết quả mới nhất.
    - Bỏ qua các khung hình không cần thiết để tăng tốc.
    - Thay đổi kích thước khung hình trước khi so sánh để giảm tải tính toán.
    """
    if os.path.exists(output_folder):
        print(f"Thư mục '{output_folder}' đã tồn tại. Đang xóa để tạo mới...")
        try:
            shutil.rmtree(output_folder)
            print("Đã xóa thư mục cũ thành công.")
        except OSError as e:
            print(f"Lỗi khi xóa thư mục {output_folder}: {e}")
            return

    os.makedirs(output_folder)
    print(f"Đã tạo thư mục mới: {output_folder}")

    cap = cv2.VideoCapture(video_path)
    if not cap.isOpened():
        print("Lỗi: Không thể mở file video.")
        return

    fps = int(cap.get(cv2.CAP_PROP_FPS))
    skip_frames_interval = max(1, fps // 2)

    print(f"Video FPS: {fps}. Sẽ kiểm tra mỗi {skip_frames_interval} khung hình.")

    prev_frame_gray_resized = None
    saved_slide_count = 0
    frame_number = 0
    
    start_time = time.time()

    while True:
        ret, frame = cap.read()
        if not ret:
            break

        frame_number += 1
        if frame_number % skip_frames_interval != 0:
            continue
            
        current_frame_gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)

        scale_percent = 640 / current_frame_gray.shape[1]
        width = int(current_frame_gray.shape[1] * scale_percent)
        height = int(current_frame_gray.shape[0] * scale_percent)
        dim = (width, height)
        current_frame_gray_resized = cv2.resize(current_frame_gray, dim, interpolation=cv2.INTER_AREA)

        if prev_frame_gray_resized is None:
            saved_slide_count += 1
            output_path = os.path.join(output_folder, f"slide_{saved_slide_count:03d}.png")
            cv2.imwrite(output_path, frame)
            print(f"Đã lưu slide đầu tiên: {output_path}")
            prev_frame_gray_resized = current_frame_gray_resized
            continue

        score = ssim(prev_frame_gray_resized, current_frame_gray_resized, data_range=255)

        if score < threshold:
            saved_slide_count += 1
            output_path = os.path.join(output_folder, f"slide_{saved_slide_count:03d}.png")
            cv2.imwrite(output_path, frame)
            print(f"Phát hiện slide mới tại khung hình {frame_number}. Đã lưu: {output_path} (SSIM: {score:.2f})")
            
            prev_frame_gray_resized = current_frame_gray_resized
    
    end_time = time.time()
    cap.release()
    print("\nHoàn tất!")
    print(f"Đã trích xuất tổng cộng {saved_slide_count} slides.")
    print(f"Thời gian xử lý: {end_time - start_time:.2f} giây.")

# --- CẤU HÌNH VÀ CHẠY SCRIPT ---
if __name__ == "__main__":
    # 1. BẠN CHỈ CẦN THAY ĐỔI ĐƯỜNG DẪN VIDEO TẠI ĐÂY
    #input_video_path = "C:/Users/Admin/Videos/BaiGiangTuan1.mp4" 
    input_video_path = r'C:\Users\thanh\Downloads\1 (6).mp4'

    # *** MỚI: Tự động tạo đường dẫn thư mục output dựa trên file input ***
    if not os.path.exists(input_video_path):
        print(f"Lỗi: Không tìm thấy file video tại '{input_video_path}'")
    else:
        # Lấy thư mục chứa file video
        video_directory = os.path.dirname(input_video_path)
        
        # Lấy tên file video không có đuôi .mp4 để làm tên thư mục
        video_filename_without_ext = os.path.splitext(os.path.basename(input_video_path))[0]
        
        # Tạo tên thư mục output, ví dụ: "BaiGiangTuan1_slides"
        output_folder_name = f"{video_filename_without_ext}_slides"
        
        # Nối chúng lại để có đường dẫn đầy đủ, an toàn trên mọi hệ điều hành
        output_slides_folder = os.path.join(video_directory, output_folder_name)

        # 2. (Tùy chỉnh) Thay đổi ngưỡng nếu cần
        similarity_threshold = 0.85

        # Gọi hàm với đường dẫn output đã được tạo tự động
        extract_slides_optimized(input_video_path, output_slides_folder, threshold=similarity_threshold)