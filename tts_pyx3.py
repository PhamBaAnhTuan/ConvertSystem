import os
import threading
import pyttsx3
import customtkinter as ctk
from tkinter import filedialog, messagebox


class TextToSpeechPyx3(ctk.CTkFrame):
    def __init__(self, master=None):
        super().__init__(master)
        self.pack(fill="both", expand=True, padx=20, pady=20)

        # --- Biến ---
        self.text_input = ctk.StringVar()
        self.txt_file = ctk.StringVar()
        self.output_path = ctk.StringVar()
        self.selected_voice = ctk.StringVar(value="Mặc định")
        self.speed = ctk.DoubleVar(value=1.0)
        self.language = ctk.StringVar(value="vi")
        self.status = ctk.StringVar(value="Chưa có tác vụ...")

        # --- Tiêu đề ---
        ctk.CTkLabel(
            self, text="🗣️ Text → Speech (Pyx3)", font=("Arial", 20, "bold")
        ).pack(pady=10)

        # --- Vùng nhập văn bản ---
        self.textbox = ctk.CTkTextbox(self, width=600, height=150)
        self.textbox.pack(pady=10)
        self.textbox.insert("1.0", "Nhập nội dung hoặc chọn file .txt...")

        # --- Nút chọn file .txt ---
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

        # --- Lựa chọn giọng đọc ---
        voice_frame = ctk.CTkFrame(self)
        voice_frame.pack(pady=5)
        ctk.CTkLabel(voice_frame, text="🎤 Chọn giọng đọc:").grid(
            row=0, column=0, padx=10
        )
        self.voice_option = ctk.CTkOptionMenu(
            voice_frame, variable=self.selected_voice, values=["Mặc định", "Nam", "Nữ"]
        )
        self.voice_option.grid(row=0, column=1, padx=10)

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
        ).pack(pady=5)
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
        voice_choice = self.selected_voice.get()
        rate = int(self.speed.get() * 150)  # tốc độ nói chuẩn khoảng 150 wpm

        if not text:
            messagebox.showwarning(
                "Thiếu nội dung", "Vui lòng nhập hoặc chọn file .txt."
            )
            return
        if not out_path:
            messagebox.showwarning("Thiếu thông tin", "Vui lòng chọn nơi lưu file.")
            return

        self.status.set("⏳ Đang tạo giọng nói...")
        threading.Thread(
            target=self._tts_thread, args=(text, out_path, lang, rate, voice_choice)
        ).start()

    def _tts_thread(self, text, out_path, lang, rate, voice_choice):
        try:
            engine = pyttsx3.init()
            voices = engine.getProperty("voices")
            print("voices: ", voices)

            # Chọn giọng đọc
            if voice_choice == "Nam":
                for v in voices:
                    if "male" in v.name.lower():
                        engine.setProperty("voice", v.id)
                        break
            elif voice_choice == "Nữ":
                for v in voices:
                    if "female" in v.name.lower():
                        engine.setProperty("voice", v.id)
                        break

            engine.setProperty("rate", rate)
            engine.save_to_file(text, out_path)
            engine.runAndWait()

            self.status.set(f"✅ Đã lưu file giọng nói: {out_path}")
            messagebox.showinfo("Hoàn tất", f"🎧 File MP3 đã được tạo:\n{out_path}")
        except Exception as e:
            self.status.set(f"❌ Lỗi: {e}")
            messagebox.showerror("Lỗi", f"Không thể tạo TTS:\n{e}")
