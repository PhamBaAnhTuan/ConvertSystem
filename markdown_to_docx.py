import os
import threading
import customtkinter as ctk
from tkinter import filedialog, messagebox
import pypandoc


class MarkdownToDocxConverter(ctk.CTkFrame):
    def __init__(self, master=None):
        super().__init__(master)
        self.pack(fill="both", expand=True, padx=20, pady=20)

        # Biến lưu đường dẫn
        self.md_input = ctk.StringVar()
        self.docx_output = ctk.StringVar()
        # self.log = ctk.StringVar(value="Chưa chọn file...")

        # --- Giao diện ---
        title = ctk.CTkLabel(
            self, text="📄 Markdown → DOCX", font=("Arial", 20, "bold")
        )
        title.pack(pady=10)

        # --- Chọn file Markdown ---
        md_file_frame = ctk.CTkFrame(self)
        md_file_frame.pack(fill="x", pady=10)
        ctk.CTkLabel(md_file_frame, text="File Markdown:").pack(side="left", padx=10)
        ctk.CTkEntry(
            md_file_frame,
            textvariable=self.md_input,
            # width=600,
            placeholder_text="Chọn file .md...",
            placeholder_text_color="#888",
        ).pack(side="left", padx=5, fill="x", expand=True)
        self.input_btn = ctk.CTkButton(
            md_file_frame, text="📂 Chọn File", command=self.select_md_file
        )
        self.input_btn.pack(side="right", padx=10)

        # --- File DOCX Output ---
        docx_file_frame = ctk.CTkFrame(self)
        docx_file_frame.pack(fill="x", pady=10)
        ctk.CTkLabel(docx_file_frame, text="File Docx:").pack(side="left", padx=10)
        ctk.CTkEntry(
            docx_file_frame,
            textvariable=self.docx_output,
            # width=600,
            placeholder_text="Đường dẫn file .docx xuất ra...",
            placeholder_text_color="#888",
        ).pack(side="left", padx=5, fill="x", expand=True)
        self.output_btn = ctk.CTkButton(
            docx_file_frame, text="💾 Chọn thư mục", command=self.select_docx_file
        )
        self.output_btn.pack(side="right", padx=10)

        # --- Log box ---
        self.log_box = ctk.CTkTextbox(self, height=100)
        self.log_box.pack(fill="both", pady=10)

        # --- Nút chạy ---
        self.convert_btn = ctk.CTkButton(
            self, text="▶️ Chuyển Đổi", command=self.run_convert
        )
        self.convert_btn.pack(pady=10)

    # --- Hàm chọn file ---
    def select_md_file(self):
        path = filedialog.askopenfilename(filetypes=[("Markdown Files", "*.md")])
        if path:
            self.md_input.set(path)
            suggested_docx = path.replace(".md", ".docx")
            self.docx_output.set(suggested_docx)
            self.log(
                "✅ Đã chọn file Markdown. \n✅ Đường dẫn DOCX tương tự đã tự động điền."
            )

    def select_docx_file(self):
        path = filedialog.askdirectory()
        if path:
            self.docx_output.set(path)
            self.log("✅ Đã chọn thư mục lưu file docx")

    def disable_buttons(self, state=True):
        state_val = "disabled" if state else "normal"
        for widget in [self.input_btn, self.output_btn, self.convert_btn]:
            widget.configure(state=state_val)

    # --- Hàm xử lý chính ---
    def run_convert(self):
        md_path = self.md_input.get()
        docx_path = self.docx_output.get()

        if not md_path or not os.path.exists(md_path):
            self.log("❌ File Markdown không tồn tại.")
            return
        if not docx_path:
            self.log("❌ Chưa chọn nơi lưu DOCX.")
            return

        self.log("⏳ Đang chuyển đổi...")
        threading.Thread(target=self._convert_thread, args=(md_path, docx_path)).start()

    def _convert_thread(self, md_path, docx_path):
        try:
            self.log("⏳ Đang xử lý file...")
            self.log_box.delete("1.0", "end")
            pypandoc.convert_file(md_path, "docx", outputfile=docx_path)
            self.log(f"✅ Đã lưu file DOCX tại: {docx_path}")
            # messagebox.showinfo("Hoàn tất", f"Đã tạo file DOCX:\n{docx_path}")
        except Exception as e:
            self.log(f"❌ Lỗi khi chuyển đổi: {e}")

    def log(self, text):
        self.log_box.insert("end", text + "\n")
        self.log_box.see("end")
