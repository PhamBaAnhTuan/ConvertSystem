import os
import threading
import customtkinter as ctk
from tkinter import filedialog, messagebox
from moviepy import ImageClip, AudioFileClip, concatenate_videoclips
import re


# ====== Hàm xử lý logic video ======
def extract_number(filename):
    match = re.search(r"(\d+)", filename)
    return int(match.group(1)) if match else None


def make_video_from_slide_audio(base_dir, log_callback, fps=24):
    """
    base_dir: thư mục chứa 2 thư mục con 'images' và 'audio'
    log_callback: hàm log ra GUI
    """
    images_dir = os.path.join(base_dir, "images")
    audio_dir = os.path.join(base_dir, "audio")
    output_dir = os.path.join(base_dir, "video_output")

    if not os.path.exists(images_dir) or not os.path.exists(audio_dir):
        log_callback("❌ Thiếu thư mục 'images' hoặc 'audio' trong đường dẫn đã chọn.")
        return

    os.makedirs(output_dir, exist_ok=True)

    # Lấy danh sách ảnh và audio
    images = [
        (extract_number(f), os.path.join(images_dir, f))
        for f in os.listdir(images_dir)
        if extract_number(f) and f.lower().endswith((".png", ".jpg"))
    ]
    audios = [
        (extract_number(f), os.path.join(audio_dir, f))
        for f in os.listdir(audio_dir)
        if extract_number(f) and f.lower().endswith((".mp3", ".wav", ".m4a"))
    ]

    images.sort(key=lambda x: x[0])
    audios.sort(key=lambda x: x[0])
    audio_map = {num: path for num, path in audios}

    log_callback(f"🖼️ {len(images)} ảnh, 🔊 {len(audios)} audio tìm thấy.")

    clips = []
    for num, img_path in images:
        if num not in audio_map:
            log_callback(f"⚠️ Bỏ slide {num} vì thiếu audio tương ứng.")
            continue
        try:
            log_callback(f"🎬 Đang xử lý slide {num}...")
            audio_clip = AudioFileClip(audio_map[num])
            img_clip = ImageClip(img_path).with_duration(audio_clip.duration)
            img_clip = img_clip.with_audio(audio_clip)
            clips.append(img_clip)
        except Exception as e:
            log_callback(f"❌ Lỗi tại slide {num}: {e}")

    if not clips:
        log_callback("❌ Không có clip hợp lệ để ghép.")
        return

    log_callback("🧩 Đang ghép video tổng...")
    final_video = concatenate_videoclips(clips, method="compose")
    out_path = os.path.join(output_dir, "video_tong.mp4")

    final_video.write_videofile(out_path, fps=fps, codec="libx264", audio_codec="aac")
    log_callback(f"✅ Đã xuất video: {out_path}")
    messagebox.showinfo("Hoàn tất", f"🎉 Video tổng đã được tạo:\n{out_path}")


# ====== Giao diện GUI ======
class VideoApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("🎬 Tool Tạo Video Slide từ Ảnh & Audio")
        self.geometry("700x500")
        ctk.set_appearance_mode("system")
        ctk.set_default_color_theme("blue")

        # --- Frame chính ---
        self.main_frame = ctk.CTkFrame(self, corner_radius=15)
        self.main_frame.pack(padx=20, pady=20, fill="both", expand=True)

        # Tiêu đề
        self.title_label = ctk.CTkLabel(
            self.main_frame,
            text="TẠO VIDEO SLIDE TỪ ẢNH & ÂM THANH",
            font=("Arial", 20, "bold"),
        )
        self.title_label.pack(pady=10)

        # Nút chọn thư mục
        self.dir_path = ctk.StringVar(value="")
        self.browse_btn = ctk.CTkButton(
            self.main_frame,
            text="📁 Chọn thư mục chứa 'images' & 'audio'",
            command=self.browse_folder,
        )
        self.browse_btn.pack(pady=10)

        # Vùng placeholder
        self.placeholder = ctk.CTkFrame(self.main_frame, height=150, fg_color="#222")
        self.placeholder.pack(fill="both", padx=30, pady=10, expand=True)

        self.placeholder_label = ctk.CTkLabel(
            self.placeholder,
            text="📂 Hãy chọn thư mục chứa:\n - images/\n - audio/",
            font=("Arial", 14),
            text_color="gray",
            justify="center",
        )
        self.placeholder_label.place(relx=0.5, rely=0.5, anchor="center")

        # Nút chạy
        self.run_btn = ctk.CTkButton(
            self.main_frame, text="▶️ Tạo Video", command=self.run_make_video
        )
        self.run_btn.pack(pady=10)

        # Vùng log
        self.log_box = ctk.CTkTextbox(self.main_frame, height=120)
        self.log_box.pack(padx=20, pady=10, fill="both")

    def log(self, message):
        """In log ra textbox"""
        self.log_box.insert("end", message + "\n")
        self.log_box.see("end")
        self.update_idletasks()

    def browse_folder(self):
        folder = filedialog.askdirectory(title="Chọn thư mục chính")
        if folder:
            self.dir_path.set(folder)
            self.placeholder_label.configure(
                text=f"📁 Thư mục được chọn:\n{folder}", text_color="white"
            )

    def run_make_video(self):
        base_dir = self.dir_path.get().strip()
        if not base_dir:
            messagebox.showwarning("Thiếu thông tin", "Vui lòng chọn thư mục trước.")
            return

        # Chạy trong thread riêng tránh treo UI
        threading.Thread(
            target=make_video_from_slide_audio, args=(base_dir, self.log)
        ).start()


if __name__ == "__main__":
    app = VideoApp()
    app.mainloop()
