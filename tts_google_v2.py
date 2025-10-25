import os
import re
import threading
import customtkinter as ctk
from tkinter import filedialog, messagebox
from gtts import gTTS


class TextToSpeechTool(ctk.CTkFrame):
    def __init__(self, master=None):
        super().__init__(master)
        self.pack(fill="both", expand=True, padx=20, pady=20)

        # --- Biến ---
        self.txt_path = ctk.StringVar()
        self.output_dir = ctk.StringVar()
        self.lang = ctk.StringVar(value="vi")
        self.speed = ctk.DoubleVar(value=1.0)

        # --- UI chính ---
        ctk.CTkLabel(self, text="🗣️Text to Speech", font=("Serif", 20, "bold")).pack(
            pady=10
        )

        # --- Chọn file txt ---
        text_file_frame = ctk.CTkFrame(self)
        text_file_frame.pack(fill="x", pady=10)
        ctk.CTkLabel(text_file_frame, text="File TXT:").pack(side="left", padx=10)
        ctk.CTkEntry(
            text_file_frame,
            textvariable=self.txt_path,
            placeholder_text="Chọn file .txt...",
            placeholder_text_color="#888",
        ).pack(side="left", padx=5, fill="x", expand=True)
        ctk.CTkButton(
            text_file_frame, text="📂 Chọn File", command=self.select_txt
        ).pack(side="right", padx=10)

        # --- Chọn thư mục output ---
        output_file_frame = ctk.CTkFrame(self)
        output_file_frame.pack(fill="x", pady=10)
        ctk.CTkLabel(output_file_frame, text="Thư mục lưu:").pack(side="left", padx=10)
        ctk.CTkEntry(
            output_file_frame,
            textvariable=self.output_dir,
            placeholder_text="Chọn hoặc tạo thư mục audio...",
            placeholder_text_color="#888",
        ).pack(side="left", padx=5, fill="x", expand=True)
        ctk.CTkButton(
            output_file_frame, text="📂 Chọn thư mục", command=self.select_output
        ).pack(side="right", padx=10)

        # --- Frame chứa ngôn ngữ và tốc độ ---
        lang_n_speed_frame = ctk.CTkFrame(self)
        lang_n_speed_frame.pack(fill="x", pady=10)

        # Cấu hình lưới để chia bố cục đẹp
        lang_n_speed_frame.grid_columnconfigure(0, weight=1)
        lang_n_speed_frame.grid_columnconfigure(1, weight=1)
        lang_n_speed_frame.grid_columnconfigure(2, weight=1)
        lang_n_speed_frame.grid_columnconfigure(3, weight=1)

        # --- Ngôn ngữ ---
        ctk.CTkLabel(lang_n_speed_frame, text="Ngôn ngữ đọc:").grid(
            row=0, column=0, padx=10, sticky="w"
        )
        ctk.CTkOptionMenu(
            lang_n_speed_frame,
            values=["vi", "en", "fr", "ja", "zh-cn"],
            variable=self.lang,
        ).grid(row=0, column=1, padx=5, sticky="w")

        # --- Tốc độ ---
        ctk.CTkLabel(lang_n_speed_frame, text="🎚️ Tốc độ nói:").grid(
            row=0, column=2, padx=10, sticky="e"
        )

        self.speed_slider = ctk.CTkSlider(
            lang_n_speed_frame,
            from_=0.5,
            to=2.0,
            number_of_steps=15,
            variable=self.speed,
        )
        self.speed_slider.grid(row=0, column=3, padx=10, sticky="ew")

        # --- Nhãn hiển thị tốc độ ---
        self.speed_label = ctk.CTkLabel(lang_n_speed_frame, text="1.0x")
        self.speed_label.grid(row=0, column=4, padx=10, sticky="w")

        # Cập nhật nhãn khi kéo slider
        self.speed_slider.configure(
            command=lambda v: self.speed_label.configure(text=f"{float(v):.1f}x")
        )

        # --- Log box ---
        self.log_box = ctk.CTkTextbox(self, height=100)
        self.log_box.pack(fill="both", expand=True, pady=10)
        self.log(
            "📢 Lưu ý, File TXT phải có cấu trúc: \nSlide n: Tiêu đề slide n \n<Nội dung slide n>"
        )

        # --- Nút thực thi ---
        ctk.CTkButton(self, text="▶️ Tạo Audio", command=self.run_tts).pack(pady=10)
        # ctk.CTkLabel(self, textvariable=self.status, text_color="gray").pack(pady=5)

    # ==============================
    # 🗂️ Chọn file và thư mục
    # ==============================
    def select_txt(self):
        path = filedialog.askopenfilename(
            title="Chọn file TXT", filetypes=[("Text Files", "*.txt")]
        )
        if path:
            self.txt_path.set(path)
            self.log("✅ Đã chọn file TXT.")
            output_dir = os.path.join(os.path.dirname(path), "audio")
            self.output_dir.set(output_dir)
            self.log(
                f"✅ Thư mục lưu audio tương tự đã được điền.{self.output_dir.get()}"
            )

    def select_output(self):
        folder = filedialog.askdirectory(title="Thư mục lưu:")
        if folder:
            self.output_dir.set(folder)
            self.log("✅ Đã chọn thư mục lưu.")

    # ==============================
    # ▶️ Chạy TTS
    # ==============================
    def run_tts(self):
        txt_path = self.txt_path.get()
        if not os.path.isfile(txt_path):
            messagebox.showerror("Lỗi", "Vui lòng chọn file TXT hợp lệ.")
            return

        out_dir = self.output_dir.get() or os.path.join(
            os.path.dirname(txt_path), "audio"
        )
        if not os.path.exists(out_dir):
            os.makedirs(out_dir)

        threading.Thread(
            target=self._tts_thread,
            args=(txt_path, out_dir, self.lang.get(), self.speed.get()),
            daemon=True,
        ).start()

    # ==============================
    # 🧠 Luồng tạo audio
    # ==============================
    def _tts_thread(self, txt_path, out_dir, lang, speed):
        try:
            self.log("⏳ Đang xử lý file...")
            self.log_box.delete("1.0", "end")

            with open(txt_path, "r", encoding="utf-8") as f:
                content = f.read()

            # tách nội dung từng slide
            slides = re.split(r"(?i)Slide\s+(\d+)\s*:", content)
            # slides sẽ có dạng: ['', '1', 'Nội dung 1', '2', 'Nội dung 2', ...]

            slide_data = []
            for i in range(1, len(slides), 2):
                num = slides[i]
                text = slides[i + 1].strip()
                if text:
                    slide_data.append((num, text))

            if not slide_data:
                self.log("❌ Không tìm thấy định dạng 'Slide n:' trong file.")
                return

            self.log(f"📄 Phát hiện {len(slide_data)} slides.")
            for num, text in slide_data:
                out_path = os.path.join(out_dir, f"audio_{num}.mp3")
                self.log(f"🎙️ Đang tạo audio_{num}.mp3 ...")
                try:
                    tts = gTTS(text=text, lang=lang, slow=(speed < 1.0))
                    temp_path = out_path.replace(".mp3", "_temp.mp3")
                    tts.save(temp_path)

                    # Nếu tốc độ khác 1.0 → chỉnh tốc độ qua ffmpeg
                    if abs(speed - 1.0) > 0.01:
                        os.system(
                            f'ffmpeg -y -i "{temp_path}" -filter:a "atempo={speed}" "{out_path}"'
                        )
                        os.remove(temp_path)
                    else:
                        os.rename(temp_path, out_path)
                except Exception as e:
                    self.log(f"❌ Lỗi slide {num}: {e}")
                    continue

            self.log(f"✅ Hoàn tất! Đã lưu {len(slide_data)} file trong: {out_dir}")

        except Exception as e:
            self.log(f"❌ Lỗi: {e}")

    # ==============================
    # 📜 Ghi log ra UI
    # ==============================
    def log(self, text):
        self.log_box.insert("end", text + "\n")
        self.log_box.see("end")
