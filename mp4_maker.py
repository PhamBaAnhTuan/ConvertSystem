import os
import threading
import re
import customtkinter as ctk
from tkinter import filedialog, messagebox
from moviepy import ImageClip, AudioFileClip, concatenate_videoclips


def extract_number(filename):
    match = re.search(r"(\d+)", filename)
    return int(match.group(1)) if match else None


class VideoMakerApp(ctk.CTkFrame):
    """Component tạo video từ ảnh + audio"""

    def __init__(self, master=None):
        super().__init__(master)
        self.pack(fill="both", expand=True, padx=20, pady=20)

        self.base_dir = ctk.StringVar()
        self.status = ctk.StringVar(value="Chưa chọn thư mục...")

        ctk.CTkLabel(self, text="🎬 Tạo Video Slide", font=("Arial", 20, "bold")).pack(
            pady=10
        )
        ctk.CTkEntry(
            self,
            textvariable=self.base_dir,
            width=600,
            placeholder_text="Chọn thư mục chính...",
        ).pack(pady=5)
        ctk.CTkButton(self, text="📁 Chọn Thư Mục", command=self.browse_folder).pack(
            pady=5
        )
        ctk.CTkButton(self, text="▶️ Tạo Video", command=self.run_make_video).pack(
            pady=10
        )
        ctk.CTkLabel(self, textvariable=self.status, text_color="gray").pack(pady=5)

    def browse_folder(self):
        folder = filedialog.askdirectory(title="Chọn thư mục chứa images & audio")
        if folder:
            self.base_dir.set(folder)
            self.status.set(f"📁 Đã chọn: {folder}")

    def run_make_video(self):
        base_dir = self.base_dir.get().strip()
        if not base_dir:
            messagebox.showwarning("Thiếu thông tin", "Vui lòng chọn thư mục trước.")
            return
        self.status.set("⏳ Đang xử lý...")
        threading.Thread(target=self._make_video_thread, args=(base_dir,)).start()

    def _make_video_thread(self, base_dir):
        try:
            images_dir = os.path.join(base_dir, "images")
            audio_dir = os.path.join(base_dir, "audio")
            output_dir = os.path.join(base_dir, "video")
            os.makedirs(output_dir, exist_ok=True)

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

            audios_map = {n: p for n, p in audios}
            clips = []
            for num, img_path in images:
                if num in audios_map:
                    audio_clip = AudioFileClip(audios_map[num])
                    img_clip = (
                        ImageClip(img_path)
                        .with_duration(audio_clip.duration)
                        .with_audio(audio_clip)
                    )
                    clips.append(img_clip)

            if not clips:
                self.status.set("❌ Không có clip hợp lệ.")
                return

            final = concatenate_videoclips(clips, method="compose")
            out_path = os.path.join(output_dir)
            final.write_videofile(out_path, fps=24, codec="libx264", audio_codec="aac")
            self.status.set(f"✅ Đã lưu: {out_path}")
            messagebox.showinfo("Hoàn tất", f"Đã tạo video tại:\n{out_path}")

        except Exception as e:
            self.status.set(f"❌ Lỗi: {e}")
