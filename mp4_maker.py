import os
import re
import threading
import customtkinter as ctk
from tkinter import filedialog, messagebox
from moviepy import ImageClip, AudioFileClip, concatenate_videoclips


class VideoBuilderTool(ctk.CTkFrame):
    """Component: Tạo video từ thư mục chứa images/ và audio/"""

    def __init__(self, master=None):
        super().__init__(master)
        self.pack(fill="both", expand=True, padx=20, pady=20)

        # --- Biến ---
        self.project_dir = ctk.StringVar()
        self.audio_count = ctk.IntVar(value=0)
        self.image_count = ctk.IntVar(value=0)
        self.total_slides = ctk.IntVar(value=0)
        self.status = ctk.StringVar(value="Chưa có tác vụ...")

        # --- Tiêu đề ---
        ctk.CTkLabel(
            self,
            text="🎬 Video Builder (Audio + Images → MP4)",
            font=("Arial", 18, "bold"),
        ).pack(pady=10)

        # --- Chọn thư mục project ---
        frm = ctk.CTkFrame(self)
        frm.pack(fill="x", pady=10)
        ctk.CTkEntry(
            frm,
            textvariable=self.project_dir,
            placeholder_text="Chọn thư mục chứa project...",
        ).pack(side="left", padx=10, fill="x", expand=True)
        ctk.CTkButton(frm, text="📂 Browse", command=self.select_folder).pack(
            side="right", padx=10
        )

        # --- Thông tin file ---
        info = ctk.CTkFrame(self)
        info.pack(fill="x", pady=5)
        ctk.CTkLabel(info, text="🎵 Audio Files:").grid(
            row=0, column=0, padx=10, pady=5, sticky="w"
        )
        ctk.CTkLabel(info, textvariable=self.audio_count).grid(
            row=0, column=1, sticky="w"
        )
        ctk.CTkLabel(info, text="🖼️ Image Files:").grid(
            row=1, column=0, padx=10, pady=5, sticky="w"
        )
        ctk.CTkLabel(info, textvariable=self.image_count).grid(
            row=0, column=3, sticky="w"
        )

        ctk.CTkLabel(info, text="📊 Tổng số slide:").grid(
            row=2, column=0, padx=10, pady=5, sticky="w"
        )
        ctk.CTkLabel(info, textvariable=self.total_slides).grid(
            row=1, column=1, sticky="w"
        )

        # --- Nút điều khiển ---
        btns = ctk.CTkFrame(self)
        btns.pack(pady=10)
        ctk.CTkButton(
            btns, text="🔍 Kiểm tra thư mục", command=self.check_folders
        ).pack(side="left", padx=10)
        ctk.CTkButton(btns, text="▶️ Tạo Video MP4", command=self.run_build).pack(
            side="left", padx=10
        )

        # --- Log ---
        self.log_box = ctk.CTkTextbox(self, height=200)
        self.log_box.pack(fill="both", expand=True, pady=10)
        ctk.CTkLabel(self, textvariable=self.status, text_color="gray").pack(pady=5)

    # ==============================
    # 🗂️ Chọn thư mục Project
    # ==============================
    def select_folder(self):
        folder = filedialog.askdirectory(title="Chọn thư mục Project")
        if folder:
            self.project_dir.set(folder)

    # ==============================
    # 📁 Kiểm tra và đếm file
    # ==============================
    def check_folders(self):
        self.log_box.delete("1.0", "end")
        project = self.project_dir.get()

        if not os.path.isdir(project):
            self.log("❌ Thư mục không hợp lệ.")
            return

        # tạo subfolders nếu thiếu
        images_dir = os.path.join(project, "images")
        audio_dir = os.path.join(project, "audio")
        for sub in [images_dir, audio_dir]:
            if not os.path.exists(sub):
                os.makedirs(sub)
                self.log(f"📁 Tạo thư mục con: {os.path.basename(sub)}")

        # đếm file
        img_files = [
            f
            for f in os.listdir(images_dir)
            if re.match(r"^audio[_-]\d+\.(png|jpg|jpeg)$", f, re.I)
        ]
        audio_files = [
            f
            for f in os.listdir(audio_dir)
            if re.match(r"^slide[_-]\d+\.(mp3|wav|m4a)$", f, re.I)
        ]

        self.image_count.set(len(img_files))
        self.audio_count.set(len(audio_files))
        self.total_slides.set(max(len(img_files), len(audio_files)))

        self.log(f"🖼️ Ảnh: {len(img_files)}, 🎵 Audio: {len(audio_files)}")
        self.log(f"🔢 Tổng số slide dự kiến: {self.total_slides.get()}")

    # ==============================
    # ▶️ Tạo Video
    # ==============================
    def run_build(self):
        project = self.project_dir.get()
        if not os.path.isdir(project):
            messagebox.showerror("Lỗi", "Vui lòng chọn thư mục Project hợp lệ.")
            return

        threading.Thread(
            target=self._build_thread, args=(project,), daemon=True
        ).start()

    def _build_thread(self, project):
        try:
            self.status.set("⏳ Đang tạo video...")
            self.log_box.delete("1.0", "end")

            images_dir = os.path.join(project, "images")
            audio_dir = os.path.join(project, "audio")
            output_path = os.path.join(project, "final_video.mp4")

            # lấy danh sách file (sắp xếp theo số)
            img_files = sorted(
                [
                    f
                    for f in os.listdir(images_dir)
                    if re.match(r"^slide[_-](\d+)\.(png|jpg|jpeg)$", f, re.I)
                ],
                key=lambda x: int(re.findall(r"\d+", x)[0]),
            )
            audio_files = sorted(
                [
                    f
                    for f in os.listdir(audio_dir)
                    if re.match(r"^slide[_-](\d+)\.(mp3|wav|m4a)$", f, re.I)
                ],
                key=lambda x: int(re.findall(r"\d+", x)[0]),
            )

            if not img_files or not audio_files:
                self.log("❌ Thiếu ảnh hoặc audio để tạo video.")
                return

            clips = []
            for img, aud in zip(img_files, audio_files):
                num = re.findall(r"\d+", img)[0]
                img_path = os.path.join(images_dir, img)
                aud_path = os.path.join(audio_dir, aud)
                self.log(f"🎞️ Xử lý slide {num}...")

                audio_clip = AudioFileClip(aud_path)
                image_clip = ImageClip(img_path).with_duration(audio_clip.duration)
                image_clip = image_clip.with_audio(audio_clip)
                clips.append(image_clip)

            final = concatenate_videoclips(clips)
            final.write_videofile(
                output_path, fps=24, codec="libx264", audio_codec="aac"
            )

            self.status.set("✅ Hoàn tất.")
            self.log(f"✅ Đã lưu video: {output_path}")
            messagebox.showinfo("Thành công", f"Video đã được tạo tại:\n{output_path}")

        except Exception as e:
            self.status.set("❌ Lỗi khi tạo video.")
            self.log(f"❌ Lỗi: {e}")

    # ==============================
    # 🧾 Ghi log tiện dụng
    # ==============================
    def log(self, text):
        self.log_box.insert("end", text + "\n")
        self.log_box.see("end")
