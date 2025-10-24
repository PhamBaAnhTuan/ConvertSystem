import os
import threading
import customtkinter as ctk
from tkinter import filedialog, messagebox
from gtts import gTTS


class TextToSpeechGoogle(ctk.CTkFrame):
    """Component: Chuyển đổi văn bản hoặc file .txt thành giọng nói (MP3)"""

    def __init__(self, master=None):
        super().__init__(master)
        self.pack(fill="both", expand=True, padx=20, pady=20)

        # Biến
        self.text_input = ctk.StringVar()
        self.txt_file = ctk.StringVar()
        self.output_path = ctk.StringVar()
        self.language = ctk.StringVar(value="vi")  # Mặc định tiếng Việt
        self.speed = ctk.StringVar(value="1.0")  # Mặc định tốc độ bình thường
        self.status = ctk.StringVar(value="Chưa có tác vụ...")

        # --- Tiêu đề ---
        ctk.CTkLabel(
            self, text="🔊 Text → Speech (TTS)", font=("Arial", 20, "bold")
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

        # --- Dòng chọn ngôn ngữ & tốc độ ---
        lang_frame = ctk.CTkFrame(self)
        lang_frame.pack(pady=10)
        ctk.CTkLabel(lang_frame, text="Ngôn ngữ:").grid(row=0, column=0, padx=5)
        ctk.CTkEntry(lang_frame, textvariable=self.language, width=80).grid(
            row=0, column=1, padx=5
        )
        ctk.CTkLabel(lang_frame, text="Tốc độ (0.5 - 2.0):").grid(
            row=0, column=2, padx=5
        )
        ctk.CTkEntry(lang_frame, textvariable=self.speed, width=80).grid(
            row=0, column=3, padx=5
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
        self.output_path.set(os.path.join(dir_path, "speech.mp3"))

    def run_tts(self):
        text = self.textbox.get("1.0", "end").strip()
        lang = self.language.get().strip()
        out_path = self.output_path.get().strip()
        speed_str = self.speed.get().strip()

        if not text:
            messagebox.showwarning(
                "Thiếu nội dung", "Vui lòng nhập văn bản hoặc chọn file .txt."
            )
            return
        if not out_path:
            messagebox.showwarning("Thiếu thông tin", "Vui lòng chọn nơi lưu file.")
            return

        try:
            speed = float(speed_str)
            if speed <= 0:
                raise ValueError
        except ValueError:
            messagebox.showwarning(
                "Giá trị không hợp lệ", "Tốc độ phải là số lớn hơn 0."
            )
            return

        self.status.set("⏳ Đang chuyển đổi văn bản thành giọng nói...")
        threading.Thread(
            target=self._tts_thread, args=(text, lang, speed, out_path)
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
