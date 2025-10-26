import os
import threading
import whisper
import customtkinter as ctk
from tkinter import filedialog, messagebox
from datetime import timedelta


class WhisperConverterApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("🎧 Voice → Text Converter (Whisper)")
        self.geometry("700x500")
        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("green")  # 🌿 xanh lá chủ đạo

        # Biến lưu file
        self.input_path = ctk.StringVar()
        self.output_path = ctk.StringVar()
        self.status = ctk.StringVar(value="Chưa có tác vụ...")

        self.build_ui()

    def build_ui(self):
        """Khởi tạo giao diện chính"""
        frame = ctk.CTkFrame(self)
        frame.pack(padx=20, pady=20, fill="both", expand=True)

        ctk.CTkLabel(
            frame, text="🎵 Chọn file MP3 hoặc MP4:", font=("Arial", 16, "bold")
        ).pack(pady=10)
        ctk.CTkEntry(frame, textvariable=self.input_path, width=450).pack(pady=5)
        ctk.CTkButton(frame, text="📂 Chọn file", command=self.select_file).pack(pady=5)

        ctk.CTkLabel(
            frame, text="💾 File TXT đầu ra:", font=("Arial", 16, "bold")
        ).pack(pady=10)
        ctk.CTkEntry(frame, textvariable=self.output_path, width=450).pack(pady=5)

        ctk.CTkButton(
            frame,
            text="▶️ Bắt đầu chuyển đổi",
            command=self.run_conversion,
            fg_color="#2e7d32",
        ).pack(pady=10)

        # Trạng thái
        ctk.CTkLabel(frame, textvariable=self.status, text_color="gray").pack(pady=10)

        # Log output
        self.log_box = ctk.CTkTextbox(frame, width=600, height=200)
        self.log_box.pack(pady=10)

    def select_file(self):
        path = filedialog.askopenfilename(
            title="Chọn file MP3 hoặc MP4",
            filetypes=[("Media files", "*.mp3 *.mp4"), ("All files", "*.*")],
        )
        if path:
            self.input_path.set(path)
            base, _ = os.path.splitext(path)
            self.output_path.set(base + ".txt")

    def run_conversion(self):
        input_path = self.input_path.get().strip()
        output_path = self.output_path.get().strip()

        if not os.path.exists(input_path):
            messagebox.showwarning("Lỗi", "Vui lòng chọn file hợp lệ.")
            return

        self.status.set("⏳ Đang chuyển đổi... Vui lòng chờ.")
        self.log_box.delete("1.0", "end")
        threading.Thread(
            target=self._convert_thread, args=(input_path, output_path)
        ).start()

    def _convert_thread(self, input_path, output_path):
        try:
            self.log(f"🎧 Đang tải model Whisper (small)...")
            model = whisper.load_model("small")

            self.log("🔍 Đang nhận dạng giọng nói...")
            result = model.transcribe(input_path, verbose=False, language="vi")

            with open(output_path, "w", encoding="utf-8") as f:
                for segment in result["segments"]:
                    start = str(timedelta(seconds=segment["start"]))
                    end = str(timedelta(seconds=segment["end"]))
                    text = segment["text"].strip()
                    f.write(f"[{start} → {end}] {text}\n")

            self.status.set("✅ Hoàn tất! File đã được tạo.")
            self.log(f"✅ Đã lưu file: {output_path}")
            messagebox.showinfo(
                "Hoàn tất", f"Đã chuyển đổi xong!\nFile TXT: {output_path}"
            )

        except Exception as e:
            self.status.set("❌ Có lỗi xảy ra.")
            self.log(f"Lỗi: {e}")

    def log(self, msg):
        self.log_box.insert("end", msg + "\n")
        self.log_box.see("end")


if __name__ == "__main__":
    app = WhisperConverterApp()
    app.mainloop()
