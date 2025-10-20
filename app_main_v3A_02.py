import customtkinter as ctk
from customtkinter import filedialog, CTkScrollableFrame
import os, re, sys, pypandoc, asyncio, edge_tts, threading
from pdf2image import convert_from_path  

# ----------------------------------------------------------------------------------
# KHỐI CODE XỬ LÝ POPPLER (BẮT BUỘC CHO PYINSTALLER)
# Thiết lập đường dẫn Poppler khi ứng dụng chạy từ file EXE
# POPPLER_BIN_DIR sẽ được gán giá trị khi PyInstaller đóng gói Poppler
# ----------------------------------------------------------------------------------
POPPLER_BIN_DIR = None 

if getattr(sys, 'frozen', False):
    # Đường dẫn thư mục temp của PyInstaller
    base_path = sys._MEIPASS if hasattr(sys, '_MEIPASS') else os.path.abspath(".")
    
    # Thư mục 'poppler' là tên đích trong gói EXE (được tạo bởi --add-binary)
    POPPLER_BIN_DIR = os.path.join(base_path, "poppler")
    
    # Thiết lập biến môi trường PATH tạm thời để pdf2image/hệ thống tìm thấy EXE/DLL
    # Cần thiết cho các file DLL và EXE của Poppler
    os.environ["PATH"] += os.pathsep + POPPLER_BIN_DIR
# ----------------------------------------------------------------------------------


# --- CẤU HÌNH GIAO DIỆN ---
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Công Cụ Tự Động Hóa Giảng Dạy (v3.0)")
        # Cấu hình cửa sổ chính
        self.geometry("880x720")

        self.tabview = ctk.CTkTabview(self, width=860, height=700)
        self.tabview.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")

        # Các tab
        self.tabview.add("1. Markdown -> DOCX")
        self.tabview.add("2. Text -> Audio (TTS)")
        self.tabview.add("3. Rename Slide Images")
        self.tabview.add("4. Tạo Video (Video Builder)")
        self.tabview.add("5. PDF -> Images (Slide Extractor)")

        for name in self.tabview._tab_dict.keys():
            self.tabview.tab(name).grid_columnconfigure(0, weight=1)

        self.setup_md2docx_tab()
        self.setup_tts_tab()
        self.setup_rename_tab()
        self.setup_video_tab()
        self.setup_pdf_tab() 

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
            # Sửa lỗi $$$ thành $$ như yêu cầu trước đó
            with open(src, encoding="utf-8") as f:
                txt = f.read().replace("$$$", "$$") 
            # Giả định Pandoc đã được cài đặt
            pypandoc.convert_text(txt, "docx", "markdown+tex_math_dollars", outputfile=dst)
            self.md_status.set(f"✅ Đã lưu: {dst}")
        except Exception as e:
            self.md_status.set(f"❌ Lỗi Pandoc/Chuyển đổi: {e}")

    # ============================================================
    # --- TAB 2: Text -> Audio ---
    # ============================================================
    VOICES = {
        # --- TIẾNG VIỆT (VIETNAMESE) ---
        "VN - Hoai My (Nữ, Tự nhiên)": 'vi-VN-HoaiMyNeural',
        "VN - Hoai An (Nữ, Miền Bắc)": 'vi-VN-HoaiAnNeural',
        "VN - Hoai Bao (Nữ, Miền Trung)": 'vi-VN-HoaiBaoNeural',
        "VN - Hoai Suong (Nữ, Miền Nam)": 'vi-VN-HoaiSuongNeural', 
        "VN - Nam Minh (Nam, Chuẩn Bắc)": 'vi-VN-NamMinhNeural', 
        
        # --- TIẾNG ANH (ENGLISH) ---
        "EN - Jenny (Nữ, US)": 'en-US-JennyNeural',
        "EN - Libby (Nữ, UK)": 'en-GB-LibbyNeural',
        "EN - Guy (Nam, US)": 'en-US-GuyNeural',
        "EN - Ryan (Nam, UK)": 'en-GB-RyanNeural',
    }

    def setup_tts_tab(self):
        tab = self.tabview.tab("2. Text -> Audio (TTS)")
        self.tts_input = ctk.StringVar()
        self.tts_voice = ctk.StringVar(value="VN - Nam Minh (Chuẩn Bắc)")
        self.tts_log = ctk.CTkTextbox(tab, height=200)

        # Cấu hình tham số điều chỉnh giọng nói
        self.tts_rate = ctk.IntVar(value=0) # +0%
        self.tts_pitch = ctk.IntVar(value=0) # +0Hz

        # Khung Input/Voice
        input_frame = ctk.CTkFrame(tab)
        input_frame.grid(row=0, column=0, padx=20, pady=10, sticky="ew", columnspan=3)
        input_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(input_frame, text="File Script (.txt):").grid(row=0, column=0, padx=10, pady=5, sticky="w")
        ctk.CTkEntry(input_frame, textvariable=self.tts_input).grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        ctk.CTkButton(input_frame, text="Browse", command=lambda: self.select_file(self.tts_input, "txt")).grid(row=0, column=2, padx=10)

        ctk.CTkLabel(input_frame, text="Giọng đọc:").grid(row=1, column=0, padx=10, pady=5, sticky="w")
        ctk.CTkComboBox(input_frame, variable=self.tts_voice, values=list(self.VOICES.keys())).grid(row=1, column=1, columnspan=2, padx=5, pady=5, sticky="ew")

        # Khung Cấu hình chi tiết
        config_frame = ctk.CTkFrame(tab)
        config_frame.grid(row=1, column=0, padx=20, pady=10, sticky="ew", columnspan=3)
        config_frame.grid_columnconfigure(1, weight=1)

        # Rate
        ctk.CTkLabel(config_frame, text="Tốc độ (Rate):").grid(row=0, column=0, padx=10, pady=5, sticky="w")
        self.rate_label = ctk.CTkLabel(config_frame, text=f"{self.tts_rate.get():+d}%")
        self.rate_label.grid(row=0, column=2, padx=10, pady=5)
        ctk.CTkSlider(config_frame, from_=-20, to=20, variable=self.tts_rate, command=lambda v: self.rate_label.configure(text=f"{int(v):+d}%")).grid(row=0, column=1, padx=10, pady=5, sticky="ew")

        # Pitch
        ctk.CTkLabel(config_frame, text="Cao độ (Pitch):").grid(row=1, column=0, padx=10, pady=5, sticky="w")
        self.pitch_label = ctk.CTkLabel(config_frame, text=f"{self.tts_pitch.get():+d}Hz")
        self.pitch_label.grid(row=1, column=2, padx=10, pady=5)
        ctk.CTkSlider(config_frame, from_=-20, to=20, variable=self.tts_pitch, command=lambda v: self.pitch_label.configure(text=f"{int(v):+d}Hz")).grid(row=1, column=1, padx=10, pady=5, sticky="ew")


        ctk.CTkButton(tab, text="▶️ TẠO AUDIO", command=self.run_tts).grid(row=2, column=0, padx=20, pady=10, sticky="w")
        self.tts_log.grid(row=3, column=0, columnspan=3, padx=20, pady=10, sticky="nsew")
        tab.grid_rowconfigure(3, weight=1) # Log box expand

    def run_tts(self):
        path = self.tts_input.get()
        if not os.path.exists(path):
            self.tts_log.insert("end", "❌ File không tồn tại\n")
            return
        
        voice = self.VOICES[self.tts_voice.get()]
        rate = f"{self.tts_rate.get():+d}%"
        pitch = f"{self.tts_pitch.get():+d}Hz"

        self.tts_log.delete("1.0", "end")
        self.tts_log.insert("end", f"▶️ Bắt đầu tạo audio...\n")
        self.tts_log.insert("end", f"⚙️ Voice: {voice} | Rate: {rate} | Pitch: {pitch}\n")

        # Sử dụng thread cho quá trình bất đồng bộ
        threading.Thread(target=self._tts_thread, args=(path, voice, rate, pitch)).start()

    def _tts_thread(self, path, voice, rate, pitch):
        outdir = os.path.join(os.path.dirname(path), "audio")
        os.makedirs(outdir, exist_ok=True)

        # Hàm parse slides (để tách chính xác)
        def parse_slides(text: str):
            text = text.lstrip("\ufeff").replace("\r\n", "\n").replace("\r", "\n")
            # Bắt tiêu đề 'Slide X: ...' và lấy phần nội dung cho đến trước 'Slide Y:' tiếp theo
            pattern = r"(Slide\s+\d+:\s*[^\n]*)(?:\n)(.*?)(?=\nSlide\s+\d+:|\Z)"
            matches = re.findall(pattern, text, flags=re.S)
            slides = []
            for idx, (title, body) in enumerate(matches, start=1):
                body = body.strip()
                if not body: continue
                slides.append({"index": idx, "title": title.strip(), "body": body})
            return slides

        async def job():
            try:
                with open(path, encoding="utf-8") as f: 
                    txt = f.read()
                
                slides = parse_slides(txt)
                
                self.tts_log.insert("end", f"🔍 Tìm thấy {len(slides)} slide có nội dung.\n")

                for s in slides:
                    slide_no = s["index"]
                    text_chunk = s["body"]
                    outfile = os.path.join(outdir, f"slide_{slide_no}.mp3")
                    
                    self.tts_log.insert("end", f"--- Xử lý Slide {slide_no}...\n")
                    
                    comm = edge_tts.Communicate(text_chunk, voice, rate=rate, pitch=pitch)
                    await comm.save(outfile)
                    
                    self.tts_log.insert("end", f"  ✅ Thành công: slide_{slide_no}.mp3\n")

                self.tts_log.insert("end", "🎉 Hoàn tất!\n")

            except Exception as e:
                self.tts_log.insert("end", f"❌ Lỗi TTS: {repr(e)}\n")

        # Cần thiết lập loop policy cho Windows khi chạy trong thread
        if sys.platform == "win32":
            try:
                asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())
            except Exception:
                pass

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
        
        ctk.CTkLabel(tab, text="Quy tắc: 'SỐ_TÊN.ĐUÔI' → 'slide-SỐ.ĐUÔI'").grid(row=1, column=0, padx=20, pady=5, sticky="w")
        
        ctk.CTkButton(tab, text="▶️ Đổi tên", command=self.rename_images).grid(row=2, column=0, padx=20, pady=10)
        self.rename_log.grid(row=3, column=0, columnspan=3, padx=20, pady=10, sticky="nsew")
        tab.grid_rowconfigure(3, weight=1)

    def rename_images(self):
        d = self.img_dir.get()
        self.rename_log.delete("1.0", "end")
        if not os.path.isdir(d):
            self.rename_log.insert("end", "❌ Thư mục không hợp lệ\n")
            return
        
        threading.Thread(target=self._rename_thread, args=(d,)).start()

    def _rename_thread(self, d):
        for f in os.listdir(d):
            # Cập nhật regex để khớp với định dạng số_tênfile.đuôi
            m = re.match(r"^(\d+)_.*(\.\w+)$", f) 
            if m:
                new = f"slide-{m.group(1)}{m.group(2)}"
                try:
                    os.rename(os.path.join(d, f), os.path.join(d, new))
                    self.rename_log.insert("end", f"✅ {f} → {new}\n")
                except Exception as e:
                    self.rename_log.insert("end", f"❌ Lỗi đổi tên {f}: {e}\n")
            else:
                 self.rename_log.insert("end", f"⏩ Bỏ qua: {f} (Không khớp quy tắc)\n")
        self.rename_log.insert("end", "🎉 Hoàn tất!\n")


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
        info.grid_columnconfigure(1, weight=1)
        info.grid_columnconfigure(3, weight=1)

        ctk.CTkLabel(info, text="Audio Files:").grid(row=0, column=0, padx=10, pady=5, sticky="w")
        ctk.CTkLabel(info, textvariable=self.audio_count).grid(row=0, column=1, padx=10, pady=5, sticky="w")
        ctk.CTkLabel(info, text="Image Files:").grid(row=0, column=2, padx=10, pady=5, sticky="w")
        ctk.CTkLabel(info, textvariable=self.img_count).grid(row=0, column=3, padx=10, pady=5, sticky="w")
        
        ctk.CTkLabel(info, text="TỔNG SLIDE DỰ KIẾN:").grid(row=1, column=0, padx=10, pady=5, sticky="w")
        ctk.CTkEntry(info, textvariable=self.slide_count, width=80).grid(row=1, column=1, padx=10, pady=5, sticky="w")

        btns = ctk.CTkFrame(tab)
        btns.grid(row=2, column=0, padx=20, pady=10)
        ctk.CTkButton(btns, text="🔍 Kiểm tra", command=self.check_folders).pack(side="left", padx=10)
        ctk.CTkButton(btns, text="▶️ Tạo index.html", command=self.create_index).pack(side="left", padx=10)
        
        self.video_log.grid(row=3, column=0, padx=20, pady=10, sticky="nsew")
        tab.grid_rowconfigure(3, weight=1)

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
                self.video_log.insert("end", f"📁 Tạo thư mục con: {os.path.basename(p)}\n")
        
        mp3 = [f for f in os.listdir(ad) if re.match(r"^slide_\d+\.mp3$", f)]
        png = [f for f in os.listdir(im) if re.match(r"^slide-\d+\.(png|jpg|jpeg)$", f, re.IGNORECASE)]
        
        self.audio_count.set(len(mp3))
        self.img_count.set(len(png))
        self.slide_count.set(max(len(mp3), len(png)))
        self.video_log.insert("end", f"🔢 Tổng số Slides (max): {self.slide_count.get()}\n")

    def create_index(self):
        project_dir = self.video_dir.get()
        self.video_log.delete("1.0", "end")

        if not os.path.isdir(project_dir):
            self.video_log.insert("end", "❌ Vui lòng chọn thư mục Project hợp lệ\n")
            return

        # --- SỬA LỖI ĐƯỜNG DẪN TÀI NGUYÊN PYINSTALLER ---
        # Kiểm tra nếu đang chạy từ gói PyInstaller
        if getattr(sys, 'frozen', False):
            # PyInstaller đặt các file 'datas' vào thư mục gốc của gói temp (sys._MEIPASS)
            # Hoặc trong thư mục gốc của dist nếu dùng --onedir
            base_path = sys._MEIPASS
        else:
            # Khi chạy trong môi trường phát triển (VS Code)
            base_path = os.path.dirname(__file__)

        template_path = os.path.join(base_path, "index_backup.html")
        output_path = os.path.join(project_dir, "index.html")
        # ----------------------------------------------------

        if not os.path.exists(template_path):
            self.video_log.insert("end", f"❌ Không tìm thấy file template: {template_path}\n")
            # In ra đường dẫn gốc để debug
            self.video_log.insert("end", f"(Debug Path: {base_path})\n")
            return

        try:
            with open(template_path, "r", encoding="utf-8") as f:
                html = f.read()

            # Cập nhật 3 dòng JS chính (logic cũ giữ nguyên)
            total = self.slide_count.get()
            html = re.sub(r"const TOTAL_SLIDES\s*=\s*\d+;", f"const TOTAL_SLIDES = {total};", html)
            html = re.sub(r"const IMAGE_PREFIX\s*=\s*['\"].*?['\"];", "const IMAGE_PREFIX = 'images/slide-';", html)
            html = re.sub(r"const AUDIO_PREFIX\s*=\s*['\"].*?['\"];", "const AUDIO_PREFIX = 'audio/slide_';", html)

            # Ghi file mới vào thư mục project
            with open(output_path, "w", encoding="utf-8") as f:
                f.write(html)

            self.video_log.insert("end", f"✅ Đã tạo file index.html tại: {output_path}\n")

        except Exception as e:
            self.video_log.insert("end", f"❌ Lỗi khi tạo file index.html: {e}\n")
            
    # ============================================================
    # --- TAB 5: PDF -> IMAGES ---
    # ============================================================
    def setup_pdf_tab(self):
        tab = self.tabview.tab("5. PDF -> Images (Slide Extractor)")
        self.pdf_file = ctk.StringVar()
        self.pdf_log = ctk.CTkTextbox(tab, height=250)
        
        if POPPLER_BIN_DIR is None:
             ctk.CTkLabel(tab, text="⚠️ Vui lòng cài đặt Poppler hoặc chạy từ file EXE đã đóng gói Poppler.", text_color="orange").grid(row=0, column=0, columnspan=3, padx=20, pady=5, sticky="w")
             row_offset = 1
        else:
             row_offset = 0

        ctk.CTkLabel(tab, text="File PDF:").grid(row=0 + row_offset, column=0, padx=20, pady=5)
        ctk.CTkEntry(tab, textvariable=self.pdf_file, width=400).grid(row=0 + row_offset, column=1, padx=5, pady=5)
        ctk.CTkButton(tab, text="Browse", command=lambda: self.select_file(self.pdf_file, "pdf")).grid(row=0 + row_offset, column=2, padx=10)
        
        ctk.CTkButton(tab, text="▶️ Chuyển thành Ảnh", command=self.convert_pdf).grid(row=1 + row_offset, column=0, padx=20, pady=10)
        self.pdf_log.grid(row=2 + row_offset, column=0, columnspan=3, padx=20, pady=10, sticky="nsew")
        tab.grid_rowconfigure(2 + row_offset, weight=1)


    def convert_pdf(self):
        pdf_path = self.pdf_file.get()
        if not os.path.exists(pdf_path):
            self.pdf_log.insert("end", "❌ File không tồn tại\n")
            return
            
        # Kiểm tra Poppler (chỉ cần kiểm tra nếu đang chạy dev, nếu chạy EXE thì POPPLER_BIN_DIR đã được set)
        if POPPLER_BIN_DIR is None and not os.environ.get('PATH'):
             self.pdf_log.insert("end", "❌ Lỗi: Không tìm thấy Poppler. Vui lòng cài đặt Poppler hoặc chạy từ EXE.\n")
             return

        out_dir = os.path.join(os.path.dirname(pdf_path), "images")
        os.makedirs(out_dir, exist_ok=True)
        self.pdf_log.insert("end", f"📁 Lưu ảnh tại: {out_dir}\n")

        # Chạy trong luồng để không bị treo GUI
        threading.Thread(target=self._pdf_worker, args=(pdf_path, out_dir)).start()

    def _pdf_worker(self, pdf_path, out_dir):
            try:
                # Sử dụng POPPLER_BIN_DIR đã được xác định ở đầu file. 
                # Nếu nó là None (môi trường dev), pdf2image sẽ tự tìm trong PATH.
                pages = convert_from_path(pdf_path, dpi=200, poppler_path=POPPLER_BIN_DIR)
                
                self.pdf_log.insert("end", f"🔍 Tổng số trang: {len(pages)}\n")
                for i, page in enumerate(pages, 1):
                    # Lưu file theo định dạng slide-1.png, slide-2.png
                    out_file = os.path.join(out_dir, f"slide-{i}.png")
                    page.save(out_file, "PNG")
                    self.pdf_log.insert("end", f"✅ Trang {i} -> {out_file}\n")
                
                self.pdf_log.insert("end", "🎉 Hoàn tất!\n")
            
            except Exception as e:
                # Thông báo lỗi chi tiết cho người dùng
                error_msg = f"❌ Lỗi: Không thể chuyển đổi PDF.\nChi tiết: {e}"
                if 'No such file or directory' in str(e) and POPPLER_BIN_DIR is not None:
                    error_msg += f"\n👉 Lỗi Poppler: Hãy đảm bảo bạn đã đóng gói các file EXE/DLL của Poppler chính xác vào thư mục 'poppler' trong EXE."
                elif 'No such file or directory' in str(e):
                     error_msg += "\n👉 Lỗi Poppler: Hãy đảm bảo Poppler đã được cài đặt và thêm vào PATH."

                self.pdf_log.insert("end", error_msg + "\n")


if __name__ == "__main__":
    if sys.platform == "win32":
        try:
            asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())
        except Exception:
            pass
    app = App()
    app.mainloop()