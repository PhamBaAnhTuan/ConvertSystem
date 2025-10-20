import customtkinter as ctk
from customtkinter import filedialog, CTkScrollableFrame
import os
import re
import sys
import pypandoc
import asyncio
import edge_tts
import threading

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")


class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Công Cụ Tự Động Hóa Giảng Dạy (v2.0)")
        self.geometry("850x700")

        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        # Tạo tabview
        self.tabview = ctk.CTkTabview(self, width=830, height=680)
        self.tabview.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")

        # Thêm các tab
        self.tabview.add("1. Markdown -> DOCX")
        self.tabview.add("2. Text -> Audio (TTS)")
        self.tabview.add("3. Rename Slide Images")
        self.tabview.add("4. Tạo Video (Video Builder)")

        # Cấu hình cột
        for name in self.tabview._tab_dict.keys():
            self.tabview.tab(name).grid_columnconfigure(0, weight=1)

        # Khởi tạo từng tab
        self.setup_md2docx_tab()
        self.setup_tts_tab()
        self.setup_rename_tab()
        self.setup_video_tab()

    # ===================================================================
    # --- TAB 1: MARKDOWN -> DOCX ---
    # ===================================================================
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

    def setup_md2docx_tab(self):
        tab = self.tabview.tab("1. Markdown -> DOCX")
        self.md_input_path = ctk.StringVar(value="")
        self.docx_output_path = ctk.StringVar(value="")
        self.highlight_style = ctk.StringVar(value="tango")

        frame = ctk.CTkFrame(tab)
        frame.grid(row=0, column=0, padx=20, pady=10, sticky="ew")
        frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(frame, text="File Markdown:").grid(row=0, column=0, padx=10, pady=5, sticky="w")
        ctk.CTkEntry(frame, textvariable=self.md_input_path).grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        ctk.CTkButton(frame, text="Browse", command=lambda: self.select_file(self.md_input_path, "md")).grid(row=0, column=2, padx=10)

        ctk.CTkLabel(frame, text="File DOCX Output:").grid(row=1, column=0, padx=10, pady=5, sticky="w")
        ctk.CTkEntry(frame, textvariable=self.docx_output_path).grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        ctk.CTkButton(frame, text="Browse", command=lambda: self.select_file(self.docx_output_path, "docx")).grid(row=1, column=2, padx=10)

        self.md_input_path.trace_add("write", self.update_docx_output_path)
        self.md_status = ctk.CTkLabel(tab, text="...", text_color="gray")
        self.md_status.grid(row=2, column=0, padx=20, pady=10, sticky="w")

        ctk.CTkButton(tab, text="▶️ BẮT ĐẦU CHUYỂN ĐỔI", command=self.run_md_conversion).grid(row=3, column=0, pady=10)

    def update_docx_output_path(self, *_):
        path = self.md_input_path.get()
        if path:
            folder = os.path.dirname(path)
            name = os.path.splitext(os.path.basename(path))[0]
            self.docx_output_path.set(os.path.join(folder, f"{name}_Output.docx"))

    def run_md_conversion(self):
        md = self.md_input_path.get()
        out = self.docx_output_path.get()
        style = self.highlight_style.get()
        if not os.path.exists(md):
            self.md_status.configure(text="❌ Không tìm thấy file Markdown", text_color="red")
            return
        threading.Thread(target=self._md_thread, args=(md, out, style)).start()

    def _md_thread(self, md, out, style):
        try:
            with open(md, encoding="utf-8") as f:
                text = f.read()
            text = text.replace("$$$", "$$")
            pypandoc.convert_text(
                text, "docx", format="markdown+tex_math_dollars",
                outputfile=out, extra_args=["--standalone", f"--highlight-style={style}"]
            )
            self.md_status.configure(text=f"✅ Đã tạo {out}", text_color="green")
        except Exception as e:
            self.md_status.configure(text=f"❌ Lỗi: {e}", text_color="red")

    # ===================================================================
    # --- TAB 2: TEXT -> AUDIO ---
    # ===================================================================
    # Giọng đọc thân thiện và mã
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
        self.tts_logbox = ctk.CTkTextbox(tab, height=200)

        ctk.CTkLabel(tab, text="File script (.txt):").grid(row=0, column=0, padx=20, pady=5, sticky="w")
        ctk.CTkEntry(tab, textvariable=self.tts_input, width=400).grid(row=0, column=1, padx=5, pady=5)
        ctk.CTkButton(tab, text="Browse", command=lambda: self.select_file(self.tts_input, "txt")).grid(row=0, column=2, padx=10)

        ctk.CTkLabel(tab, text="Giọng đọc:").grid(row=1, column=0, padx=20, pady=5, sticky="w")
        ctk.CTkComboBox(tab, variable=self.tts_voice, values=list(self.VOICES.keys())).grid(row=1, column=1, columnspan=2, padx=5, pady=5, sticky="ew")

        ctk.CTkButton(tab, text="▶️ TẠO AUDIO", command=self.run_tts).grid(row=2, column=0, padx=20, pady=10)
        self.tts_logbox.grid(row=3, column=0, columnspan=3, padx=20, pady=10, sticky="nsew")

    def run_tts(self):
        f = self.tts_input.get()
        if not os.path.exists(f):
            self.tts_logbox.insert("end", "❌ Không tìm thấy file\n")
            return
        voice = self.VOICES[self.tts_voice.get()]
        threading.Thread(target=self._tts_thread, args=(f, voice)).start()

    def _tts_thread(self, path, voice):
        outdir = os.path.join(os.path.dirname(path), "audio")
        os.makedirs(outdir, exist_ok=True)
        self.tts_logbox.insert("end", f"📁 Lưu tại: {outdir}\n")
        async def job():
            with open(path, encoding="utf-8") as f: txt = f.read()
            chunks = re.split(r"\nSlide\s+\d+:", txt)[1:]
            for i, chunk in enumerate(chunks, 1):
                file = os.path.join(outdir, f"slide_{i}.mp3")
                comm = edge_tts.Communicate(chunk.strip(), voice)
                await comm.save(file)
                self.tts_logbox.insert("end", f"✅ Slide {i}\n")
        asyncio.run(job())

    # ===================================================================
    # --- TAB 3: RENAME SLIDE IMAGES ---
    # ===================================================================
    def setup_rename_tab(self):
        tab = self.tabview.tab("3. Rename Slide Images")
        self.img_dir = ctk.StringVar()
        self.rename_log = ctk.CTkTextbox(tab, height=200)

        ctk.CTkLabel(tab, text="Thư mục ảnh:").grid(row=0, column=0, padx=20, pady=5, sticky="w")
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

    # ===================================================================
    # --- TAB 4: TẠO VIDEO (VIDEO BUILDER) ---
    # ===================================================================
    def setup_video_tab(self):
        tab = self.tabview.tab("4. Tạo Video (Video Builder)")
        self.video_dir = ctk.StringVar()
        self.audio_count = ctk.IntVar(value=0)
        self.img_count = ctk.IntVar(value=0)
        self.slide_count = ctk.IntVar(value=0)
        self.video_log = ctk.CTkTextbox(tab, height=200)

        # chọn thư mục
        frm = ctk.CTkFrame(tab)
        frm.grid(row=0, column=0, padx=20, pady=10, sticky="ew")
        frm.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(frm, text="Thư mục project:").grid(row=0, column=0, padx=10, pady=5)
        ctk.CTkEntry(frm, textvariable=self.video_dir).grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        ctk.CTkButton(frm, text="Browse", command=lambda: self.select_directory(self.video_dir)).grid(row=0, column=2, padx=10)

        # hiển thị số lượng
        info = ctk.CTkFrame(tab)
        info.grid(row=1, column=0, padx=20, pady=10, sticky="ew")
        ctk.CTkLabel(info, text="Audio (.mp3):").grid(row=0, column=0, sticky="w", padx=10)
        ctk.CTkLabel(info, textvariable=self.audio_count).grid(row=0, column=1, sticky="w")
        ctk.CTkLabel(info, text="Ảnh (.png):").grid(row=1, column=0, sticky="w", padx=10)
        ctk.CTkLabel(info, textvariable=self.img_count).grid(row=1, column=1, sticky="w")
        ctk.CTkLabel(info, text="Tổng Slide:").grid(row=2, column=0, sticky="w", padx=10)
        ctk.CTkEntry(info, textvariable=self.slide_count, width=80).grid(row=2, column=1, sticky="w")

        btns = ctk.CTkFrame(tab)
        btns.grid(row=2, column=0, padx=20, pady=10)
        ctk.CTkButton(btns, text="🔍 Kiểm tra thư mục", command=self.check_folders).pack(side="left", padx=10)
        ctk.CTkButton(btns, text="▶️ Tạo index.html", command=self.create_index).pack(side="left", padx=10)
        self.video_log.grid(row=3, column=0, padx=20, pady=10, sticky="nsew")

    def check_folders(self):
        d = self.video_dir.get()
        self.video_log.delete("1.0", "end")
        if not os.path.isdir(d):
            self.video_log.insert("end", "❌ Vui lòng chọn thư mục hợp lệ\n")
            return
        ad = os.path.join(d, "audio")
        im = os.path.join(d, "images")
        for p in [ad, im]:
            if not os.path.exists(p):
                os.makedirs(p)
                self.video_log.insert("end", f"📁 Tạo thư mục: {p}\n")
            else:
                self.video_log.insert("end", f"✅ Có sẵn: {p}\n")
        mp3 = [f for f in os.listdir(ad) if f.endswith(".mp3")]
        png = [f for f in os.listdir(im) if f.endswith(".png")]
        self.audio_count.set(len(mp3))
        self.img_count.set(len(png))
        self.slide_count.set(max(len(mp3), len(png)))
        self.video_log.insert("end", f"🔢 Tổng slide gợi ý: {self.slide_count.get()}\n")

    def create_index(self):
        d = self.video_dir.get()
        if not os.path.isdir(d):
            self.video_log.insert("end", "❌ Thư mục không hợp lệ\n")
            return
        template = os.path.join(os.path.dirname(__file__), "index.html")
        out = os.path.join(d, "index.html")
        try:
            with open(template, encoding="utf-8") as f:
                html = f.read()
            html = re.sub(r"const TOTAL_SLIDES = \d+;", f"const TOTAL_SLIDES = {self.slide_count.get()};", html)
            with open(out, "w", encoding="utf-8") as f:
                f.write(html)
            self.video_log.insert("end", f"✅ Tạo thành công: {out}\n")
        except Exception as e:
            self.video_log.insert("end", f"❌ Lỗi: {e}\n")


if __name__ == "__main__":
    if sys.platform == "win32":
        try:
            asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())
        except Exception:
            pass
    app = App()
    app.mainloop()
