import cv2
import os
from skimage.metrics import structural_similarity as ssim
import numpy as np

def extract_slides(video_path, output_folder, threshold=0.85):
    """
    Trích xuất các slide tĩnh từ video dựa trên sự thay đổi khung hình.

    :param video_path: Đường dẫn tới file video MP4 đầu vào.
    :param output_folder: Thư mục để lưu các ảnh slide được trích xuất.
    :param threshold: Ngưỡng tương đồng SSIM. Nếu độ tương đồng < ngưỡng này,
                      coi như là một slide mới. Giá trị từ 0 đến 1.
                      Giá trị thấp hơn sẽ nhạy hơn với các thay đổi nhỏ.
    """
    # Tạo thư mục đầu ra nếu nó chưa tồn tại
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
        print(f"Đã tạo thư mục: {output_folder}")

    # Mở file video
    cap = cv2.VideoCapture(video_path)
    if not cap.isOpened():
        print("Lỗi: Không thể mở file video.")
        return

    prev_frame_gray = None
    saved_slide_count = 0
    frame_number = 0

    while True:
        # Đọc từng khung hình
        ret, frame = cap.read()
        if not ret:
            break # Kết thúc video

        frame_number += 1
        # Chuyển khung hình sang ảnh xám để so sánh hiệu quả hơn
        current_frame_gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)

        # Khung hình đầu tiên luôn được lưu lại làm slide đầu tiên
        if prev_frame_gray is None:
            saved_slide_count += 1
            output_path = os.path.join(output_folder, f"slide_{saved_slide_count:03d}.png")
            cv2.imwrite(output_path, frame)
            print(f"Đã lưu slide đầu tiên: {output_path}")
            prev_frame_gray = current_frame_gray
            continue

        # Tính toán độ tương đồng cấu trúc (SSIM) giữa khung hình trước và hiện tại
        # Thêm data_range để đảm bảo tính toán chính xác trên ảnh 8-bit
        score = ssim(prev_frame_gray, current_frame_gray, data_range=current_frame_gray.max() - current_frame_gray.min())

        # Nếu độ tương đồng thấp hơn ngưỡng, đây là một slide mới
        if score < threshold:
            saved_slide_count += 1
            output_path = os.path.join(output_folder, f"slide_{saved_slide_count:03d}.png")
            cv2.imwrite(output_path, frame)
            print(f"Phát hiện slide mới tại khung hình {frame_number}. Đã lưu: {output_path} (SSIM: {score:.2f})")
            
            # Cập nhật khung hình tham chiếu
            prev_frame_gray = current_frame_gray

    # Giải phóng tài nguyên
    cap.release()
    print("\nHoàn tất! Đã trích xuất tổng cộng", saved_slide_count, "slides.")

# --- CẤU HÌNH VÀ CHẠY SCRIPT ---
if __name__ == "__main__":
    # 1. Thay đổi đường dẫn đến file video của bạn
    input_video_path = r'C:\Users\thanh\Downloads\1 (5).mp4' #"path/to/your/bai_giang.mp4" 

    # 2. Đặt tên thư mục để lưu các slide
    output_slides_folder = "extracted_slides"

    # 3. (Tùy chỉnh) Thay đổi ngưỡng nếu cần. 
    #    - Tăng lên gần 1.0 (ví dụ: 0.9) nếu code bắt nhầm các chuyển động nhỏ.
    #    - Giảm xuống (ví dụ: 0.8) nếu code bỏ lỡ một vài slide.
    similarity_threshold = 0.85

    extract_slides(input_video_path, output_slides_folder, threshold=similarity_threshold)