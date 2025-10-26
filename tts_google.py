import os
import threading
import customtkinter as ctk
from tkinter import filedialog, messagebox
from gtts import gTTS


class TextToSpeechGoogle(ctk.CTkFrame):
    def __init__(self, master=None):
        super().__init__(master)
        self.pack(fill="both", expand=True, padx=20, pady=20)

        # Biến
        self.text_input = ctk.StringVar()
        self.txt_file = ctk.StringVar()
        self.output_path = ctk.StringVar()
        self.language = ctk.StringVar(value="vi")
        self.speed = ctk.DoubleVar(value=1.0)
        self.status = ctk.StringVar(value="Chưa có tác vụ...")

        # --- Tiêu đề ---
        ctk.CTkLabel(
            self, text="🔊 Text → Speech (Google)", font=("Arial", 20, "bold")
        ).pack(pady=10)

        # --- Vùng nhập văn bản ---
        self.textbox = ctk.CTkTextbox(self, width=800, height=150)
        self.textbox.pack(pady=10)
        self.textbox.insert("1.0", "Nhập văn bản hoặc chọn file .txt...")

        # --- Nút chọn file txt ---
        file_frame = ctk.CTkFrame(self)
        file_frame.pack(pady=5, fill="x")
        ctk.CTkEntry(
            file_frame,
            textvariable=self.txt_file,
            placeholder_text="Chọn file .txt...",
            width=400,
        ).pack(side="left", padx=10, pady=5, expand=True, fill="x")
        ctk.CTkButton(
            file_frame, text="📂 Chọn File", command=self.select_txt_file
        ).pack(side="right", padx=10)

        # --- Lựa chọn ngôn ngữ ---
        lang_frame = ctk.CTkFrame(self)
        lang_frame.pack(pady=10)
        ctk.CTkLabel(lang_frame, text="🌍 Chọn ngôn ngữ:").grid(
            row=0, column=0, padx=10
        )
        lang_menu = ctk.CTkOptionMenu(
            lang_frame,
            variable=self.language,
            values=["vi", "en", "ja", "ko", "fr", "de", "zh-CN"],
        )
        lang_menu.grid(row=0, column=1, padx=10)
        # --- Thanh trượt tốc độ ---
        speed_frame = ctk.CTkFrame(self)
        speed_frame.pack(pady=10)
        ctk.CTkLabel(speed_frame, text="🎚️ Tốc độ nói:").grid(row=0, column=0, padx=10)
        self.speed_slider = ctk.CTkSlider(
            speed_frame, from_=0.5, to=2.0, number_of_steps=15, variable=self.speed
        )
        self.speed_slider.grid(row=0, column=1, padx=10)
        self.speed_label = ctk.CTkLabel(speed_frame, text="1.0x")
        self.speed_label.grid(row=0, column=2)
        self.speed_slider.configure(
            command=lambda v: self.speed_label.configure(text=f"{float(v):.1f}x")
        )

        # --- Chọn nơi lưu file ---
        ctk.CTkEntry(
            self,
            textvariable=self.output_path,
            width=400,
            placeholder_text="Tên file MP3 xuất ra...",
        ).pack(pady=10)
        ctk.CTkButton(
            self, text="💾 Chọn nơi lưu file", command=self.select_output
        ).pack(pady=5)

        # --- Nút chạy ---
        ctk.CTkButton(self, text="▶️ Chuyển Đổi", command=self.run_tts).pack(pady=10)

        # --- Trạng thái ---
        ctk.CTkLabel(self, textvariable=self.status, text_color="gray").pack(pady=5)

    def select_txt_file(self):
        path = filedialog.askopenfilename(filetypes=[("Text files", "*.txt")])
        if path:
            self.txt_file.set(path)
            try:
                with open(path, "r", encoding="utf-8") as f:
                    content = f.read()
                self.textbox.delete("1.0", "end")
                self.textbox.insert("1.0", content)
            except Exception as e:
                messagebox.showerror("Lỗi", f"Không thể đọc file: {e}")

    def select_output(self):
        dir_path = filedialog.askdirectory()
        if not dir_path:
            return
        lang = self.language.get()
        file_name = (
            self.txt_file.get().split("/")[-1].replace(".txt", "")
        )  # Bỏ đuôi .txt
        self.output_path.set(os.path.join(dir_path, f"{file_name}_{lang}.mp3"))
        # print("output_path:", self.output_path.get())

    def run_tts(self):
        text = self.textbox.get("1.0", "end").strip()
        lang = self.language.get().strip()
        out_path = self.output_path.get().strip()
        # print("out_path:", out_path)
        speed = self.speed.get()

        if not text:
            messagebox.showwarning(
                "Thiếu nội dung", "Vui lòng nhập văn bản hoặc chọn file .txt."
            )
            return
        if not out_path:
            messagebox.showwarning("Thiếu thông tin", "Vui lòng chọn nơi lưu file.")
            return

        self.status.set("⏳ Đang chuyển đổi văn bản thành giọng nói...")
        threading.Thread(
            target=self._tts_thread,
            args=(text, lang, speed, out_path),
            daemon=True,
        ).start()

    def _tts_thread(self, text, lang, speed, out_path):
        try:
            # Tạo file giọng nói
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

            self.status.set(f"✅ Đã lưu file giọng nói: {out_path}")
            messagebox.showinfo("Hoàn tất", f"🎧 File MP3 đã được tạo:\n{out_path}")
        except Exception as e:
            self.status.set(f"❌ Lỗi: {e}")
            messagebox.showerror("Lỗi", f"Không thể tạo file TTS:\n{e}")
