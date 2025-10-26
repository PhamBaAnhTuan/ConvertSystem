import os, time
import threading
import whisper
import customtkinter as ctk
from tkinter import filedialog, messagebox
from datetime import timedelta


class AudioToTextTool(ctk.CTkFrame):
    def __init__(self, master=None):
        super().__init__(master)
        self.pack(fill="both", expand=True, padx=20, pady=20)

        # Biến lưu file
        self.input_path = ctk.StringVar()
        self.output_path = ctk.StringVar()
        self.progress_value = ctk.DoubleVar(value=0)
        # self.status = ctk.StringVar(value="Chưa có tác vụ...")

        # UI
        ctk.CTkLabel(
            self, text="🎧 Voice → Text Converter (Whisper)", font=("Arial", 20, "bold")
        ).pack(pady=10)

        self.input_file_frame = ctk.CTkFrame(self)
        self.input_file_frame.pack(pady=10, fill="x")
        ctk.CTkLabel(self.input_file_frame, text="🎵 File MP3 hoặc MP4:").pack(
            side="left", padx=10
        )
        ctk.CTkEntry(self.input_file_frame, textvariable=self.input_path).pack(
            side="left", padx=5, fill="x", expand=True
        )
        self.input_btn = ctk.CTkButton(
            self.input_file_frame, text="📂 Chọn File", command=self.select_audio_file
        )
        self.input_btn.pack(side="right", padx=10)

        self.output_file_frame = ctk.CTkFrame(self)
        self.output_file_frame.pack(pady=10, fill="x")
        ctk.CTkLabel(self.output_file_frame, text="🎵 Thư mục lưu:").pack(
            side="left", padx=10
        )
        ctk.CTkEntry(self.output_file_frame, textvariable=self.output_path).pack(
            side="left", padx=5, fill="x", expand=True
        )
        self.output_btn = ctk.CTkButton(
            self.output_file_frame,
            text="📂 Chọn thư mục",
            command=self.select_audio_file,
        )
        self.output_btn.pack(side="right", padx=10)

        # Thanh tiến trình
        self.progress_bar = ctk.CTkFrame(self)
        self.progress_bar.pack(pady=10, fill="x")
        self.slider_label = ctk.CTkLabel(self.progress_bar, text="Tiến trình: ").pack(
            side="left", padx=10
        )
        self.slider = ctk.CTkProgressBar(
            self.progress_bar, variable=self.progress_value
        )
        self.slider.pack(pady=10, fill="x")
        self.slider.set(0)

        # Log output
        self.log_box = ctk.CTkTextbox(self)
        self.log_box.pack(fill="x", pady=10)

        # Btn chuyển đổi
        self.convert_btn = ctk.CTkButton(
            self, text="▶️ Chuyển đổi", command=self.run_conversion
        )
        self.convert_btn.pack(padx=10)

    def select_audio_file(self):
        path = filedialog.askopenfilename(
            title="Chọn file MP3 hoặc MP4",
            filetypes=[("Media files", "*.mp3 *.mp4"), ("All files", "*.*")],
        )
        if path:
            self.input_path.set(path)
            base, ext = os.path.splitext(path)
            self.output_path.set(base + ".txt")
            self.log(
                f"✅ Đã chọn file: {ext} \n✅ File .txt sẽ tự động lưu cùng thư mục"
            )

    def select_target_folder(self):
        folder = filedialog.askdirectory(
            title="Chọn thư mục lưu",
            initialdir=os.getcwd(),
        )
        if folder:
            self.output_path.set(folder)
            self.log(f"✅ Đã chọn thư mục: {folder}")

    def disable_buttons(self, state=True):
        state_val = "disabled" if state else "normal"
        for widget in [self.input_btn, self.output_btn, self.convert_btn]:
            widget.configure(state=state_val)

    def run_conversion(self):
        input_path = self.input_path.get().strip()
        output_path = self.output_path.get().strip()

        if not os.path.exists(input_path):
            self.log("❌ File đầu vào không tồn tại.")
            return

        # Reset trạng thái
        self.progress_value.set(0)
        self.slider.set(0)
        self.log_box.delete("1.0", "end")
        self.log("⏳ Đang chuyển đổi... ")
        self.disable_buttons(True)

        threading.Thread(
            target=self._convert_thread, args=(input_path, output_path)
        ).start()

    def _convert_thread(self, input_path, output_path):
        # model = "large-v2"
        model = "small"
        try:
            self.log(f"🎧 Đang tải model Whisper ({model})...")
            model = whisper.load_model(model)

            self.log("🔍 Đang nhận dạng giọng nói...")
            result = model.transcribe(input_path, verbose=False, language="vi")
            total_segments = len(result["segments"])
            self.log(f"📊 Tổng số đoạn thoại: {total_segments}")

            with open(output_path, "w", encoding="utf-8") as f:
                for i, segment in enumerate(result["segments"], start=1):
                    start = str(timedelta(seconds=segment["start"]))
                    end = str(timedelta(seconds=segment["end"]))
                    text = segment["text"].strip()
                    f.write(f"[{start} → {end}] {text}\n")

                    # Cập nhật tiến trình
                    percent = i / total_segments
                    self.progress_value.set(percent)
                    self.slider.set(percent)
                    self.log(f"⏳ Đang xử lý: {int(percent*100)}%")
                    self.update_idletasks()
                    time.sleep(0.05)

            self.log("✅ Hoàn tất! File đã được tạo.")
            self.log(f"✅ Đã lưu file: {output_path}")
            self.progress_value.set(1)
            self.slider.set(1)

        except Exception as e:
            self.log("❌ Có lỗi xảy ra.")
            self.log(f"Lỗi: {e}")
        finally:
            self.disable_buttons(False)

    def log(self, msg):
        self.log_box.insert("end", msg + "\n")
        self.log_box.see("end")


# if __name__ == "__main__":
#     app = AudioToTextTool()
#     app.mainloop()
