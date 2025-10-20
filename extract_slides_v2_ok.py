import cv2
import os
from skimage.metrics import structural_similarity as ssim
import numpy as np
import time

def extract_slides_optimized(video_path, output_folder, threshold=0.85):
    """
     phiên bản tối ưu của hàm trích xuất slide.
    - Bỏ qua các khung hình không cần thiết để tăng tốc.
    - Thay đổi kích thước khung hình trước khi so sánh để giảm tải tính toán.
    """
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
        print(f"Đã tạo thư mục: {output_folder}")

    cap = cv2.VideoCapture(video_path)
    if not cap.isOpened():
        print("Lỗi: Không thể mở file video.")
        return

    # Lấy số khung hình mỗi giây (FPS) của video
    fps = int(cap.get(cv2.CAP_PROP_FPS))
    # *** TỐI ƯU 1: CHỈ KIỂM TRA 2 KHUNG HÌNH MỖI GIÂY ***
    # Nếu fps = 30, chúng ta sẽ bỏ qua 14 khung hình rồi mới kiểm tra khung hình tiếp theo.
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
        # *** TỐI ƯU 1: BỎ QUA KHUNG HÌNH ***
        if frame_number % skip_frames_interval != 0:
            continue
            
        current_frame_gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)

        # *** TỐI ƯU 2: THAY ĐỔI KÍCH THƯỚC ẢNH TRƯỚC KHI SO SÁNH ***
        # Giảm kích thước xuống chiều rộng 640px, giữ nguyên tỷ lệ
        scale_percent = 640 / current_frame_gray.shape[1]
        width = int(current_frame_gray.shape[1] * scale_percent)
        height = int(current_frame_gray.shape[0] * scale_percent)
        dim = (width, height)
        current_frame_gray_resized = cv2.resize(current_frame_gray, dim, interpolation=cv2.INTER_AREA)

        if prev_frame_gray_resized is None:
            saved_slide_count += 1
            output_path = os.path.join(output_folder, f"slide_{saved_slide_count:03d}.png")
            cv2.imwrite(output_path, frame) # Lưu ảnh chất lượng gốc
            print(f"Đã lưu slide đầu tiên: {output_path}")
            prev_frame_gray_resized = current_frame_gray_resized
            continue

        score = ssim(prev_frame_gray_resized, current_frame_gray_resized, data_range=255)

        if score < threshold:
            saved_slide_count += 1
            # QUAN TRỌNG: Vẫn lưu khung hình gốc, chưa resize
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
    #input_video_path = "path/to/your/bai_giang.mp4" 
    input_video_path = r'C:\Users\thanh\Downloads\1 (5).mp4' #"path/to/your/bai_giang.mp4" 
    output_slides_folder = "extracted_slides_optimized"
    similarity_threshold = 0.85

    extract_slides_optimized(input_video_path, output_slides_folder, threshold=similarity_threshold)