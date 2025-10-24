import os
import threading
import customtkinter as ctk
from tkinter import filedialog, messagebox
import pypandoc


class MarkdownToDocxConverter(ctk.CTkFrame):
    """Component chuyển đổi Markdown → DOCX"""

    def __init__(self, master=None):
        super().__init__(master)
        self.pack(fill="both", expand=True, padx=20, pady=20)

        # Biến lưu đường dẫn
        self.md_input = ctk.StringVar()
        self.docx_output = ctk.StringVar()
        self.status = ctk.StringVar(value="Chưa chọn file...")

        # --- Giao diện ---
        title = ctk.CTkLabel(
            self, text="📄 Markdown → DOCX", font=("Arial", 20, "bold")
        )
        title.pack(pady=10)

        # --- Chọn file Markdown ---
        self.md_entry = ctk.CTkEntry(
            self,
            textvariable=self.md_input,
            width=600,
            placeholder_text="Chọn file .md...",
        )
        self.md_entry.pack(pady=5)
        ctk.CTkButton(self, text="📂 Chọn Markdown", command=self.select_md_file).pack(
            pady=5
        )

        # --- File DOCX Output ---
        self.docx_entry = ctk.CTkEntry(
            self,
            textvariable=self.docx_output,
            width=600,
            placeholder_text="Đường dẫn file .docx xuất ra...",
        )
        self.docx_entry.pack(pady=5)
        ctk.CTkButton(
            self, text="💾 Lưu thành DOCX", command=self.select_docx_file
        ).pack(pady=5)

        # --- Nút chạy ---
        ctk.CTkButton(self, text="▶️ Chuyển Đổi", command=self.run_convert).pack(pady=10)

        # --- Trạng thái ---
        ctk.CTkLabel(self, textvariable=self.status, text_color="gray").pack(pady=5)

    # --- Hàm chọn file ---
    def select_md_file(self):
        path = filedialog.askopenfilename(filetypes=[("Markdown Files", "*.md")])
        if path:
            self.md_input.set(path)
            suggested_docx = path.replace(".md", ".docx")
            self.docx_output.set(suggested_docx)
            self.status.set("✅ Đã chọn file Markdown.")

    def select_docx_file(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".docx", filetypes=[("Word Document", "*.docx")]
        )
        if path:
            self.docx_output.set(path)

    # --- Hàm xử lý chính ---
    def run_convert(self):
        md_path = self.md_input.get()
        docx_path = self.docx_output.get()

        if not md_path or not os.path.exists(md_path):
            self.status.set("❌ File Markdown không tồn tại.")
            return
        if not docx_path:
            self.status.set("❌ Chưa chọn nơi lưu DOCX.")
            return

        self.status.set("⏳ Đang chuyển đổi...")
        threading.Thread(target=self._convert_thread, args=(md_path, docx_path)).start()

    def _convert_thread(self, md_path, docx_path):
        try:
            pypandoc.convert_file(md_path, "docx", outputfile=docx_path)
            self.status.set(f"✅ Đã lưu file DOCX: {docx_path}")
            messagebox.showinfo("Hoàn tất", f"Đã tạo file DOCX:\n{docx_path}")
        except Exception as e:
            self.status.set(f"❌ Lỗi khi chuyển đổi: {e}")
