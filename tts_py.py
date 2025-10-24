import os
import threading
import customtkinter as ctk
from tkinter import filedialog, messagebox
from gtts import gTTS
import re
from time import time


class TextToSpeechPy(ctk.CTkFrame):
    """Component: chuyển đổi văn bản thành giọng nói (MP3)"""

    def __init__(self, master=None):
        super().__init__(master)
        self.pack(fill="both", expand=True, padx=20, pady=20)

        # Biến lưu text và output
        self.text_input = ctk.StringVar()
        self.output_path = ctk.StringVar()
        self.language = ctk.StringVar(value="vi")
        self.status = ctk.StringVar(value="Chưa có tác vụ...")

        # --- Tiêu đề ---
        ctk.CTkLabel(
            self, text="🔊 Text → Speech (TTS)", font=("Arial", 20, "bold")
        ).pack(pady=10)

        # --- Ô nhập văn bản ---
        self.textbox = ctk.CTkTextbox(self, width=600, height=150)
        self.textbox.pack(pady=10)
        self.textbox.insert("1.0", "Nhập nội dung văn bản cần đọc...")
        self.placeholder_active = True

        # --- Chọn ngôn ngữ ---
        ctk.CTkLabel(self, text="Ngôn ngữ (mã ISO, ví dụ: 'vi' hoặc 'en')").pack()
        self.lang_entry = ctk.CTkEntry(self, textvariable=self.language, width=100)
        self.lang_entry.pack(pady=5)

        # --- Nút chọn file đầu ra ---
        ctk.CTkEntry(
            self,
            textvariable=self.output_path,
            width=400,
            placeholder_text="Tên file MP3 xuất ra...",
        ).pack(pady=5)
        ctk.CTkButton(
            self, text="💾 Chọn nơi lưu file", command=self.select_output
        ).pack(pady=5)

        # --- Nút chạy ---
        ctk.CTkButton(self, text="▶️ Chuyển Đổi", command=self.run_tts).pack(pady=10)

        # --- Trạng thái ---
        ctk.CTkLabel(self, textvariable=self.status, text_color="gray").pack(pady=5)

    def select_output(self):
        # chọn thư mục đích (thay vì file) và tự tạo tên file .mp3
        dir_path = filedialog.askdirectory()
        if not dir_path:
            return
        self.output_path.set(os.path.join(dir_path, "speech.mp3"))

    def run_tts(self):
        text = self.textbox.get("1.0", "end").strip()
        lang = self.language.get().strip()
        out_path = self.output_path.get()

        if not text:
            messagebox.showwarning("Thiếu nội dung", "Vui lòng nhập văn bản.")
            return
        if not out_path:
            messagebox.showwarning("Thiếu thông tin", "Vui lòng chọn nơi lưu file.")
            return

        self.status.set("⏳ Đang chuyển đổi văn bản thành giọng nói...")
        threading.Thread(target=self._tts_thread, args=(text, lang, out_path)).start()

    def _tts_thread(self, text, lang, out_path):
        try:
            tts = gTTS(text=text, lang=lang)
            tts.save(out_path)
            self.status.set(f"✅ Đã lưu file giọng nói: {out_path}")
            messagebox.showinfo("Hoàn tất", f"🎧 File MP3 đã được tạo:\n{out_path}")
        except Exception as e:
            self.status.set(f"❌ Lỗi: {e}")
