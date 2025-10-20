import customtkinter as ctk
from customtkinter import filedialog, CTkScrollableFrame
import os, re, sys, pypandoc, asyncio, edge_tts, threading
from pdf2image import convert_from_path  # <— NEW

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Công Cụ Tự Động Hóa Giảng Dạy (v3.0)")
        self.geometry("880x720")

        self.tabview = ctk.CTkTabview(self, width=860, height=700)
        self.tabview.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")

        # Các tab
        self.tabview.add("1. Markdown -> DOCX")
        self.tabview.add("2. Text -> Audio (TTS)")
        self.tabview.add("3. Rename Slide Images")
        self.tabview.add("4. Tạo Video (Video Builder)")
        self.tabview.add("5. PDF -> Images (Slide Extractor)")  # NEW

        for name in self.tabview._tab_dict.keys():
            self.tabview.tab(name).grid_columnconfigure(0, weight=1)

        self.setup_md2docx_tab()
        self.setup_tts_tab()
        self.setup_rename_tab()
        self.setup_video_tab()
        self.setup_pdf_tab()  # NEW TAB

    # ============================================================
    # --- UTILITY ---
    # ============================================================
    def select_file(self, entry_var, extension="*"):
        path = filedialog.askopenfilename(
            defaultextension=f".{extension}",
            filetypes=[(f"{extension.upper()} files", f"*.{extension}"), ("All files", "*.*")]
        )
        if path:
            entry_var.set(path)

    def select_directory(self, entry_var):
        path = filedialog.askdirectory()
        if path:
            entry_var.set(path)

    # ============================================================
    # --- TAB 1: Markdown -> DOCX ---
    # ============================================================
    def setup_md2docx_tab(self):
        tab = self.tabview.tab("1. Markdown -> DOCX")
        self.md_input = ctk.StringVar()
        self.docx_output = ctk.StringVar()
        self.md_status = ctk.StringVar(value="...")

        frame = ctk.CTkFrame(tab)
        frame.grid(row=0, column=0, padx=20, pady=10, sticky="ew")
        frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(frame, text="File Markdown:").grid(row=0, column=0, padx=10, pady=5)
        ctk.CTkEntry(frame, textvariable=self.md_input).grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        ctk.CTkButton(frame, text="Browse", command=lambda: self.select_file(self.md_input, "md")).grid(row=0, column=2, padx=10)

        ctk.CTkLabel(frame, text="File DOCX Output:").grid(row=1, column=0, padx=10, pady=5)
        ctk.CTkEntry(frame, textvariable=self.docx_output).grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        ctk.CTkButton(frame, text="Browse", command=lambda: self.select_file(self.docx_output, "docx")).grid(row=1, column=2, padx=10)

        self.md_label = ctk.CTkLabel(tab, textvariable=self.md_status, text_color="gray")
        self.md_label.grid(row=1, column=0, padx=20, pady=10, sticky="w")

        ctk.CTkButton(tab, text="▶️ Chuyển Đổi", command=self.run_md).grid(row=2, column=0, pady=10)

    def run_md(self):
        src, dst = self.md_input.get(), self.docx_output.get()
        if not os.path.exists(src):
            self.md_status.set("❌ File không tồn tại!")
            return
        threading.Thread(target=self._md_thread, args=(src, dst)).start()

    def _md_thread(self, src, dst):
        try:
            with open(src, encoding="utf-8") as f:
                txt = f.read().replace("$$$", "$$")
            pypandoc.convert_text(txt, "docx", "markdown+tex_math_dollars", outputfile=dst)
            self.md_status.set(f"✅ Đã lưu: {dst}")
        except Exception as e:
            self.md_status.set(f"❌ Lỗi: {e}")

    # ============================================================
    # --- TAB 2: Text -> Audio ---
    # ============================================================
    VOICES = {
        # --- TIẾNG VIỆT (VIETNAMESE) ---
        
        # Giọng Nữ (Female)
        "VN - Hoai My (Nữ, Tự nhiên)": 'vi-VN-HoaiMyNeural',
        "VN - Hoai An (Nữ, Miền Bắc)": 'vi-VN-HoaiAnNeural',
        "VN - Hoai Bao (Nữ, Miền Trung)": 'vi-VN-HoaiBaoNeural',
        "VN - Hoai Suong (Nữ, Miền Nam)": 'vi-VN-HoaiSuongNeural', 
        
        # Giọng Nam (Male)
        "VN - Nam Minh (Nam, Chuẩn Bắc)": 'vi-VN-NamMinhNeural', 
        #"VN - Nam Quan (Nam, Miền Trung)": 'vi-VN-NamQuanNeural', 
        #"VN - Nam Phong (Nam, Miền Nam)": 'vi-VN-NamPhongNeural', 
        
        # --- TIẾNG ANH (ENGLISH) ---
        
        # Giọng Nữ (Female)
        "EN - Jenny (Nữ, US)": 'en-US-JennyNeural',
        "EN - Libby (Nữ, UK)": 'en-GB-LibbyNeural',
        
        # Giọng Nam (Male)
        "EN - Guy (Nam, US)": 'en-US-GuyNeural',
        "EN - Ryan (Nam, UK)": 'en-GB-RyanNeural',
    }

    def setup_tts_tab(self):
        tab = self.tabview.tab("2. Text -> Audio (TTS)")
        self.tts_input = ctk.StringVar()
        self.tts_voice = ctk.StringVar(value="VN - Nam Minh (Bắc)")
        self.tts_log = ctk.CTkTextbox(tab, height=200)

        ctk.CTkLabel(tab, text="File Script (.txt):").grid(row=0, column=0, padx=20, pady=5, sticky="w")
        ctk.CTkEntry(tab, textvariable=self.tts_input, width=400).grid(row=0, column=1, padx=5, pady=5)
        ctk.CTkButton(tab, text="Browse", command=lambda: self.select_file(self.tts_input, "txt")).grid(row=0, column=2, padx=10)

        ctk.CTkLabel(tab, text="Giọng đọc:").grid(row=1, column=0, padx=20, pady=5, sticky="w")
        ctk.CTkComboBox(tab, variable=self.tts_voice, values=list(self.VOICES.keys())).grid(row=1, column=1, columnspan=2, padx=5, pady=5, sticky="ew")

        ctk.CTkButton(tab, text="▶️ TẠO AUDIO", command=self.run_tts).grid(row=2, column=0, padx=20, pady=10)
        self.tts_log.grid(row=3, column=0, columnspan=3, padx=20, pady=10, sticky="nsew")

    def run_tts(self):
        path = self.tts_input.get()
        if not os.path.exists(path):
            self.tts_log.insert("end", "❌ File không tồn tại\n")
            return
        voice = self.VOICES[self.tts_voice.get()]
        threading.Thread(target=self._tts_thread, args=(path, voice)).start()

    def _tts_thread(self, path, voice):
        outdir = os.path.join(os.path.dirname(path), "audio")
        os.makedirs(outdir, exist_ok=True)
        async def job():
            with open(path, encoding="utf-8") as f: txt = f.read()
            slides = re.split(r"\nSlide\s+\d+:", txt)[1:]
            for i, s in enumerate(slides, 1):
                outfile = os.path.join(outdir, f"slide_{i}.mp3")
                comm = edge_tts.Communicate(s.strip(), voice)
                await comm.save(outfile)
                self.tts_log.insert("end", f"✅ Slide {i}\n")
        asyncio.run(job())

    # ============================================================
    # --- TAB 3: Rename Images ---
    # ============================================================
    def setup_rename_tab(self):
        tab = self.tabview.tab("3. Rename Slide Images")
        self.img_dir = ctk.StringVar()
        self.rename_log = ctk.CTkTextbox(tab, height=200)

        ctk.CTkLabel(tab, text="Thư mục ảnh:").grid(row=0, column=0, padx=20, pady=5)
        ctk.CTkEntry(tab, textvariable=self.img_dir, width=400).grid(row=0, column=1, padx=5, pady=5)
        ctk.CTkButton(tab, text="Browse", command=lambda: self.select_directory(self.img_dir)).grid(row=0, column=2, padx=10)
        ctk.CTkButton(tab, text="▶️ Đổi tên", command=self.rename_images).grid(row=1, column=0, padx=20, pady=10)
        self.rename_log.grid(row=2, column=0, columnspan=3, padx=20, pady=10, sticky="nsew")

    def rename_images(self):
        d = self.img_dir.get()
        if not os.path.isdir(d):
            self.rename_log.insert("end", "❌ Thư mục không hợp lệ\n")
            return
        for f in os.listdir(d):
            m = re.match(r"^(\d+)_.*(\.\w+)$", f)
            if m:
                new = f"slide-{m.group(1)}{m.group(2)}"
                os.rename(os.path.join(d, f), os.path.join(d, new))
                self.rename_log.insert("end", f"✅ {f} → {new}\n")

    # ============================================================
    # --- TAB 4: VIDEO BUILDER ---
    # ============================================================
    def setup_video_tab(self):
        tab = self.tabview.tab("4. Tạo Video (Video Builder)")
        self.video_dir = ctk.StringVar()
        self.audio_count = ctk.IntVar(value=0)
        self.img_count = ctk.IntVar(value=0)
        self.slide_count = ctk.IntVar(value=0)
        self.video_log = ctk.CTkTextbox(tab, height=200)

        frm = ctk.CTkFrame(tab)
        frm.grid(row=0, column=0, padx=20, pady=10, sticky="ew")
        frm.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(frm, text="Thư mục project:").grid(row=0, column=0, padx=10)
        ctk.CTkEntry(frm, textvariable=self.video_dir).grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        ctk.CTkButton(frm, text="Browse", command=lambda: self.select_directory(self.video_dir)).grid(row=0, column=2, padx=10)

        info = ctk.CTkFrame(tab)
        info.grid(row=1, column=0, padx=20, pady=10, sticky="ew")
        ctk.CTkLabel(info, text="Audio:").grid(row=0, column=0)
        ctk.CTkLabel(info, textvariable=self.audio_count).grid(row=0, column=1)
        ctk.CTkLabel(info, text="Images:").grid(row=1, column=0)
        ctk.CTkLabel(info, textvariable=self.img_count).grid(row=1, column=1)
        ctk.CTkLabel(info, text="Slides:").grid(row=2, column=0)
        ctk.CTkEntry(info, textvariable=self.slide_count, width=80).grid(row=2, column=1)

        btns = ctk.CTkFrame(tab)
        btns.grid(row=2, column=0, padx=20, pady=10)
        ctk.CTkButton(btns, text="🔍 Kiểm tra", command=self.check_folders).pack(side="left", padx=10)
        ctk.CTkButton(btns, text="▶️ Tạo index.html", command=self.create_index).pack(side="left", padx=10)
        self.video_log.grid(row=3, column=0, padx=20, pady=10, sticky="nsew")

    def check_folders(self):
        d = self.video_dir.get()
        self.video_log.delete("1.0", "end")
        if not os.path.isdir(d):
            self.video_log.insert("end", "❌ Thư mục không hợp lệ\n")
            return
        ad, im = os.path.join(d, "audio"), os.path.join(d, "images")
        for p in [ad, im]:
            if not os.path.exists(p):
                os.makedirs(p)
                self.video_log.insert("end", f"📁 Tạo {p}\n")
        mp3 = [f for f in os.listdir(ad) if f.endswith(".mp3")]
        png = [f for f in os.listdir(im) if f.endswith(".png")]
        self.audio_count.set(len(mp3))
        self.img_count.set(len(png))
        self.slide_count.set(max(len(mp3), len(png)))
        self.video_log.insert("end", f"🔢 Slides: {self.slide_count.get()}\n")

    def create_index(self):
        project_dir = self.video_dir.get()
        self.video_log.delete("1.0", "end")

        if not os.path.isdir(project_dir):
            self.video_log.insert("end", "❌ Vui lòng chọn thư mục Project hợp lệ\n")
            return

        # Đường dẫn file template gốc (index_backup.html)
        template_path = os.path.join(os.path.dirname(__file__), "index_backup.html")
        output_path = os.path.join(project_dir, "index.html")

        if not os.path.exists(template_path):
            self.video_log.insert("end", f"❌ Không tìm thấy file template: {template_path}\n")
            return

        try:
            with open(template_path, "r", encoding="utf-8") as f:
                html = f.read()

            # Cập nhật 3 dòng JS chính
            total = self.slide_count.get()
            html = re.sub(r"const TOTAL_SLIDES\s*=\s*\d+;", f"const TOTAL_SLIDES = {total};", html)
            html = re.sub(r"const IMAGE_PREFIX\s*=\s*['\"].*?['\"];", "const IMAGE_PREFIX = 'images/slide-';", html)
            html = re.sub(r"const AUDIO_PREFIX\s*=\s*['\"].*?['\"];", "const AUDIO_PREFIX = 'audio/slide_';", html)

            # Ghi file mới vào thư mục project
            with open(output_path, "w", encoding="utf-8") as f:
                f.write(html)

            self.video_log.insert("end", f"✅ Đã tạo file index.html tại: {output_path}\n")
            self.video_log.insert("end", f"🔢 Tổng số slide: {total}\n")

        except Exception as e:
            self.video_log.insert("end", f"❌ Lỗi khi tạo file index.html: {e}\n")

    # ============================================================
    # --- TAB 5: PDF -> IMAGES ---
    # ============================================================
    def setup_pdf_tab(self):
        tab = self.tabview.tab("5. PDF -> Images (Slide Extractor)")
        self.pdf_file = ctk.StringVar()
        self.pdf_log = ctk.CTkTextbox(tab, height=250)

        ctk.CTkLabel(tab, text="File PDF:").grid(row=0, column=0, padx=20, pady=5)
        ctk.CTkEntry(tab, textvariable=self.pdf_file, width=400).grid(row=0, column=1, padx=5, pady=5)
        ctk.CTkButton(tab, text="Browse", command=lambda: self.select_file(self.pdf_file, "pdf")).grid(row=0, column=2, padx=10)
        ctk.CTkButton(tab, text="▶️ Chuyển thành Ảnh", command=self.convert_pdf).grid(row=1, column=0, padx=20, pady=10)
        self.pdf_log.grid(row=2, column=0, columnspan=3, padx=20, pady=10, sticky="nsew")

    def convert_pdf(self):
        pdf_path = self.pdf_file.get()
        if not os.path.exists(pdf_path):
            self.pdf_log.insert("end", "❌ File không tồn tại\n")
            return
        out_dir = os.path.join(os.path.dirname(pdf_path), "images")
        os.makedirs(out_dir, exist_ok=True)
        self.pdf_log.insert("end", f"📁 Lưu ảnh tại: {out_dir}\n")

        def worker():
            try:
                pages = convert_from_path(pdf_path, dpi=200)
                self.pdf_log.insert("end", f"🔍 Tổng số trang: {len(pages)}\n")
                for i, page in enumerate(pages, 1):
                    out_file = os.path.join(out_dir, f"slide-{i}.png")
                    page.save(out_file, "PNG")
                    self.pdf_log.insert("end", f"✅ Trang {i} -> {out_file}\n")
                self.pdf_log.insert("end", "🎉 Hoàn tất!\n")
            except Exception as e:
                self.pdf_log.insert("end", f"❌ Lỗi: {e}\n")

        threading.Thread(target=worker).start()


if __name__ == "__main__":
    if sys.platform == "win32":
        try:
            asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())
        except Exception:
            pass
    app = App()
    app.mainloop()
