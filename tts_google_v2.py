import os
import re
import threading
import customtkinter as ctk
from tkinter import filedialog, messagebox
from gtts import gTTS


class TextToSpeechToolV2(ctk.CTkFrame):
    """Component: Chuyển file TXT thành nhiều file audio (TTS từng Slide)"""

    def __init__(self, master=None):
        super().__init__(master)
        self.pack(fill="both", expand=True, padx=20, pady=20)

        # --- Biến ---
        self.txt_path = ctk.StringVar()
        self.output_dir = ctk.StringVar()
        self.lang = ctk.StringVar(value="vi")
        self.speed = ctk.DoubleVar(value=1.0)
        self.status = ctk.StringVar(value="Chưa thực hiện...")

        # --- UI chính ---
        ctk.CTkLabel(
            self, text="🗣️ Text to Speech (Multi-Slide)", font=("Arial", 18, "bold")
        ).pack(pady=10)

        # --- Chọn file txt ---
        frm1 = ctk.CTkFrame(self)
        frm1.pack(fill="x", pady=10)
        ctk.CTkLabel(frm1, text="File TXT:").pack(side="left", padx=10)
        ctk.CTkEntry(
            frm1, textvariable=self.txt_path, placeholder_text="Chọn file .txt..."
        ).pack(side="left", padx=5, fill="x", expand=True)
        ctk.CTkButton(frm1, text="📂 Browse", command=self.select_txt).pack(
            side="right", padx=10
        )

        # --- Chọn thư mục output ---
        frm2 = ctk.CTkFrame(self)
        frm2.pack(fill="x", pady=10)
        ctk.CTkLabel(frm2, text="Thư mục Output:").pack(side="left", padx=10)
        ctk.CTkEntry(
            frm2,
            textvariable=self.output_dir,
            placeholder_text="Chọn hoặc tạo thư mục audio...",
        ).pack(side="left", padx=5, fill="x", expand=True)
        ctk.CTkButton(frm2, text="📂 Browse", command=self.select_output).pack(
            side="right", padx=10
        )

        # --- Chọn ngôn ngữ ---
        frm3 = ctk.CTkFrame(self)
        frm3.pack(fill="x", pady=10)
        ctk.CTkLabel(frm3, text="Ngôn ngữ đọc:").pack(side="left", padx=10)
        ctk.CTkOptionMenu(
            frm3,
            values=["vi", "en", "fr", "ja", "zh-cn"],
            variable=self.lang,
        ).pack(side="left", padx=5)

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

        # --- Log box ---
        self.log_box = ctk.CTkTextbox(self, height=150)
        self.log_box.pack(fill="both", expand=True, pady=10)

        # --- Nút thực thi ---
        ctk.CTkButton(
            self, text="▶️ Chuyển TXT → Audio Slides", command=self.run_tts
        ).pack(pady=10)
        ctk.CTkLabel(self, textvariable=self.status, text_color="gray").pack(pady=5)

    # ==============================
    # 🗂️ Chọn file và thư mục
    # ==============================
    def select_txt(self):
        path = filedialog.askopenfilename(
            title="Chọn file TXT", filetypes=[("Text Files", "*.txt")]
        )
        if path:
            self.txt_path.set(path)

    def select_output(self):
        folder = filedialog.askdirectory(title="Chọn thư mục Output")
        if folder:
            self.output_dir.set(folder)

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
            self.status.set("⏳ Đang xử lý file...")
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
                self.status.set("❌ Không có slide hợp lệ.")
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

            self.status.set("✅ Hoàn tất tạo audio.")
            self.log(f"✅ Hoàn tất! Đã lưu {len(slide_data)} file trong: {out_dir}")
            messagebox.showinfo(
                "Thành công", f"Đã tạo {len(slide_data)} file audio tại:\n{out_dir}"
            )

        except Exception as e:
            self.status.set("❌ Lỗi khi xử lý.")
            self.log(f"❌ Lỗi: {e}")

    # ==============================
    # 📜 Ghi log ra UI
    # ==============================
    def log(self, text):
        self.log_box.insert("end", text + "\n")
        self.log_box.see("end")
