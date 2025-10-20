import customtkinter as ctk
from customtkinter import filedialog, CTkScrollableFrame
import os
import re
import sys
import pypandoc
import asyncio
import edge_tts
import threading
from concurrent.futures import ThreadPoolExecutor # Unused but kept for context if needed

# --- CẤU HÌNH GIAO DIỆN ---
ctk.set_appearance_mode("System")  # Modes: "System" (default), "Dark", "Light"
ctk.set_default_color_theme("blue")

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        self.title("Công Cụ Tự Động Hóa Giảng Dạy (v1.0)")
        self.geometry("800x650") # Tăng kích thước cửa sổ một chút
        
        # Grid layout for the main frame
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)
        
        # Create a tabview
        self.tabview = ctk.CTkTabview(self, width=780, height=630)
        self.tabview.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
        
        # Add tabs
        self.tabview.add("1. Markdown -> DOCX")
        self.tabview.add("2. Text -> Audio (TTS)")
        self.tabview.add("3. Rename Slide Images")
        
        # Configure tabs
        self.tabview.tab("1. Markdown -> DOCX").grid_columnconfigure(0, weight=1)
        self.tabview.tab("2. Text -> Audio (TTS)").grid_columnconfigure(0, weight=1)
        self.tabview.tab("3. Rename Slide Images").grid_columnconfigure(0, weight=1)
        
        # Initialize tabs content
        self.setup_md2docx_tab()
        self.setup_tts_tab()
        self.setup_rename_tab()
        
    # --- UTILITY FUNCTIONS ---

    def select_file(self, entry_var, extension="*"):
        """Mở cửa sổ chọn file và cập nhật Entry."""
        filepath = filedialog.askopenfilename(
            defaultextension=f".{extension}",
            filetypes=[(f"{extension.upper()} files", f"*.{extension}"), ("All files", "*.*")]
        )
        if filepath:
            entry_var.set(filepath)

    def select_directory(self, entry_var):
        """Mở cửa sổ chọn thư mục và cập nhật Entry."""
        dirpath = filedialog.askdirectory()
        if dirpath:
            entry_var.set(dirpath)

    # =========================================================================================
    # --- TAB 1: MARKDOWN TO DOCX ---
    # =========================================================================================

    def setup_md2docx_tab(self):
        tab = self.tabview.tab("1. Markdown -> DOCX")
        
        # Variables
        self.md_input_path = ctk.StringVar(value="")
        self.docx_output_path = ctk.StringVar(value="")
        self.highlight_style = ctk.StringVar(value="tango")
        
        # --- Input Frame ---
        input_frame = ctk.CTkFrame(tab)
        input_frame.grid(row=0, column=0, padx=20, pady=10, sticky="ew")
        input_frame.grid_columnconfigure(1, weight=1)
        
        ctk.CTkLabel(input_frame, text="File Markdown (.md):").grid(row=0, column=0, padx=10, pady=5, sticky="w")
        ctk.CTkEntry(input_frame, textvariable=self.md_input_path, width=400).grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        ctk.CTkButton(input_frame, text="Browse", command=lambda: self.select_file(self.md_input_path, "md")).grid(row=0, column=2, padx=10, pady=5)

        ctk.CTkLabel(input_frame, text="File DOCX Output:").grid(row=1, column=0, padx=10, pady=5, sticky="w")
        ctk.CTkEntry(input_frame, textvariable=self.docx_output_path, width=400).grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        ctk.CTkButton(input_frame, text="Browse", command=lambda: self.select_file(self.docx_output_path, "docx")).grid(row=1, column=2, padx=10, pady=5)
        
        # Bind for auto-generating output path
        self.md_input_path.trace_add("write", self.update_docx_output_path)
        
        # --- Config Frame ---
        config_frame = ctk.CTkFrame(tab)
        config_frame.grid(row=1, column=0, padx=20, pady=10, sticky="ew")
        config_frame.grid_columnconfigure(1, weight=1)
        
        ctk.CTkLabel(config_frame, text="Highlight Style:").grid(row=0, column=0, padx=10, pady=10, sticky="w")
        styles = ["tango", "pygments", "kate", "espresso", "zenburn", "default"]
        ctk.CTkComboBox(config_frame, variable=self.highlight_style, values=styles).grid(row=0, column=1, padx=10, pady=10, sticky="ew")
        
        # --- Action & Status ---
        self.md_status_label = ctk.CTkLabel(tab, text="...", fg_color="transparent")
        self.md_status_label.grid(row=3, column=0, padx=20, pady=(10, 5), sticky="w")
        
        ctk.CTkButton(tab, text="▶️ BẮT ĐẦU CHUYỂN ĐỔI", command=self.run_md_conversion).grid(row=2, column=0, padx=20, pady=20)


    def update_docx_output_path(self, *args):
        """Tự động tạo tên file DOCX khi chọn file MD."""
        input_file = self.md_input_path.get()
        if input_file:
            base_name = os.path.splitext(os.path.basename(input_file))[0]
            folder = os.path.dirname(input_file)
            output_file = os.path.join(folder, f"{base_name}_Output.docx")
            self.docx_output_path.set(output_file)

    def run_md_conversion(self):
        """Khởi chạy chuyển đổi MD -> DOCX trong một luồng riêng biệt."""
        md_file = self.md_input_path.get()
        docx_file = self.docx_output_path.get()
        style = self.highlight_style.get()
        
        if not os.path.exists(md_file):
            self.md_status_label.configure(text=f"❌ Lỗi: Không tìm thấy file đầu vào.", text_color="red")
            return
        
        self.md_status_label.configure(text="⚙️ Đang xử lý... Vui lòng đợi.", text_color="orange")
        self.update_idletasks() # Force update UI

        # Chạy trong luồng để UI không bị treo
        threading.Thread(target=self._md_converter_thread, args=(md_file, docx_file, style)).start()

    def _md_converter_thread(self, md_file_path, docx_file_path, highlight_style):
        try:
            with open(md_file_path, 'r', encoding='utf-8') as f:
                md_content = f.read()

            # --- PREPROCESSING (from md2docx_v3A_ok.py) ---
            log = "Lịch sử tiền xử lý:\n"
            
            # 1. Sửa lỗi ba dấu đô la ($$$) thành hai dấu ($$)
            if '$$$' in md_content:
                md_content = md_content.replace('$$$', '$$')
                log += "  - Đã chuẩn hóa '$$$' thành '$$'.\n"

            # 2. Sửa lỗi dư dấu gạch chéo ngược cho các lệnh LaTeX
            original_pattern_count = len(re.findall(r'\\\\([a-zA-Z]+)', md_content))
            if original_pattern_count > 0:
                md_content = re.sub(r'\\\\([a-zA-Z]+)', r'\\\1', md_content)
                log += f"  - Đã sửa {original_pattern_count} lỗi lệnh LaTeX dư dấu '\\'.\n"
                
            # 3. Sửa lỗi dư dấu gạch chéo ngược cho ký tự gạch dưới
            original_underscore_count = md_content.count(r'\_')
            if original_underscore_count > 0:
                md_content = md_content.replace(r'\_', '_')
                log += f"  - Đã sửa {original_underscore_count} lỗi ký tự gạch dưới '\\_'.\n"
            
            self.md_status_label.configure(text=f"⚙️ {log}\nĐang gọi Pandoc...")

            # --- CONVERSION (from md2docx_v3A_ok.py) ---
            input_format = 'markdown+tex_math_dollars'
            extra_args = ['--standalone', '--mathml', f'--highlight-style={highlight_style}'] 

            pypandoc.convert_text(
                source=md_content,
                to='docx',
                format=input_format,
                outputfile=docx_file_path,
                extra_args=extra_args
            )
            
            self.md_status_label.configure(text=f"✅🏆 Thành công! File DOCX đã lưu tại:\n{docx_file_path}", text_color="green")

        except Exception as e:
            error_msg = f"❌ Lỗi Pandoc/Chuyển đổi: {e}"
            if "pypandoc.pandoc_download.NotInstalledError" in str(e):
                error_msg += "\n\n⚠️ Gợi ý: Hãy cài đặt Pandoc (https://pandoc.org/installing.html)"
            self.md_status_label.configure(text=error_msg, text_color="red")
            
    # =========================================================================================
    # --- TAB 2: TEXT TO AUDIO (TTS) ---
    # =========================================================================================

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
        
        # Variables
        self.tts_input_file = ctk.StringVar(value="")
        self.tts_voice = ctk.StringVar(value="EN - Guy (Nam, US)") # Mặc định tiếng Anh
        self.tts_rate = ctk.IntVar(value=0) # +0%
        self.tts_pitch = ctk.IntVar(value=10) # +10Hz
        self.tts_volume = ctk.IntVar(value=100) # +100%
        self.tts_max_retries = ctk.IntVar(value=3)
        # NEW Variables for Slide Range
        self.tts_start_slide = ctk.StringVar(value="1") 
        self.tts_end_slide = ctk.StringVar(value="")    # Mặc định trống = đến hết
        
        # --- Input Frame ---
        input_frame = ctk.CTkFrame(tab)
        input_frame.grid(row=0, column=0, padx=20, pady=10, sticky="ew")
        input_frame.grid_columnconfigure(1, weight=1)
        
        # Row 0: Input File
        ctk.CTkLabel(input_frame, text="File Script (.txt):").grid(row=0, column=0, padx=10, pady=5, sticky="w")
        ctk.CTkEntry(input_frame, textvariable=self.tts_input_file, width=400).grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        ctk.CTkButton(input_frame, text="Browse", command=lambda: self.select_file(self.tts_input_file, "txt")).grid(row=0, column=2, padx=10, pady=5)
        
        # Row 1: Voice Selection
        ctk.CTkLabel(input_frame, text="Giọng Đọc (Voice):").grid(row=1, column=0, padx=10, pady=5, sticky="w")
        ctk.CTkComboBox(input_frame, variable=self.tts_voice, values=list(self.VOICES.keys())).grid(row=1, column=1, columnspan=2, padx=5, pady=5, sticky="ew")

        # Row 2: Slide Range Selection (NEW)
        ctk.CTkLabel(input_frame, text="Phạm vi Slide:").grid(row=2, column=0, padx=10, pady=5, sticky="w")
        
        range_input_frame = ctk.CTkFrame(input_frame, fg_color="transparent")
        range_input_frame.grid(row=2, column=1, columnspan=2, padx=5, pady=5, sticky="w")
        
        ctk.CTkLabel(range_input_frame, text="Từ:").pack(side="left", padx=(0, 5))
        ctk.CTkEntry(range_input_frame, textvariable=self.tts_start_slide, width=50).pack(side="left", padx=(0, 15))

        ctk.CTkLabel(range_input_frame, text="Đến:").pack(side="left", padx=(0, 5))
        ctk.CTkEntry(range_input_frame, textvariable=self.tts_end_slide, width=50).pack(side="left", padx=(0, 10))

        # --- Config Frame (Scrollable) ---
        config_frame = CTkScrollableFrame(tab, label_text="Cấu Hình Âm Thanh")
        config_frame.grid(row=1, column=0, padx=20, pady=10, sticky="nsew")
        tab.grid_rowconfigure(1, weight=1) # Allow config frame to expand
        config_frame.grid_columnconfigure(1, weight=1)

        # Rate
        ctk.CTkLabel(config_frame, text="Tốc độ (Rate):").grid(row=0, column=0, padx=10, pady=5, sticky="w")
        self.rate_label = ctk.CTkLabel(config_frame, text=f"{self.tts_rate.get():+d}%")
        self.rate_label.grid(row=0, column=2, padx=10, pady=5)
        ctk.CTkSlider(config_frame, from_=-20, to=20, variable=self.tts_rate, command=lambda v: self.rate_label.configure(text=f"{int(v):+d}%")).grid(row=0, column=1, padx=10, pady=5, sticky="ew")

        # Pitch
        ctk.CTkLabel(config_frame, text="Cao độ (Pitch):").grid(row=1, column=0, padx=10, pady=5, sticky="w")
        self.pitch_label = ctk.CTkLabel(config_frame, text=f"+{self.tts_pitch.get()}Hz")
        self.pitch_label.grid(row=1, column=2, padx=10, pady=5)
        ctk.CTkSlider(config_frame, from_=-20, to=20, variable=self.tts_pitch, command=lambda v: self.pitch_label.configure(text=f"{int(v):+d}Hz")).grid(row=1, column=1, padx=10, pady=5, sticky="ew")

        # Volume
        ctk.CTkLabel(config_frame, text="Âm lượng (Volume):").grid(row=2, column=0, padx=10, pady=5, sticky="w")
        self.volume_label = ctk.CTkLabel(config_frame, text=f"+{self.tts_volume.get()}%")
        self.volume_label.grid(row=2, column=2, padx=10, pady=5)
        ctk.CTkSlider(config_frame, from_=0, to=100, variable=self.tts_volume, command=lambda v: self.volume_label.configure(text=f"+{int(v)}%")).grid(row=2, column=1, padx=10, pady=5, sticky="ew")

        # Retries
        ctk.CTkLabel(config_frame, text="Max Retries:").grid(row=3, column=0, padx=10, pady=5, sticky="w")
        ctk.CTkEntry(config_frame, textvariable=self.tts_max_retries, width=50).grid(row=3, column=1, padx=10, pady=5, sticky="w")

        # --- Action & Status ---
        self.tts_log_text = ctk.CTkTextbox(tab, height=150, activate_scrollbars=True, wrap="word")
        self.tts_log_text.grid(row=3, column=0, padx=20, pady=(10, 5), sticky="ew")
        
        ctk.CTkButton(tab, text="▶️ TẠO AUDIO", command=self.run_tts_generation).grid(row=2, column=0, padx=20, pady=20)


    def tts_log(self, message):
        """Hàm ghi log an toàn cho luồng."""
        self.tts_log_text.insert(ctk.END, message + "\n")
        self.tts_log_text.see(ctk.END)
        self.update_idletasks()
        
    def run_tts_generation(self):
        """Khởi chạy tạo audio trong luồng riêng biệt và xử lý tham số slide."""
        input_file = self.tts_input_file.get()
        if not os.path.exists(input_file):
            self.tts_log_text.delete("1.0", ctk.END)
            self.tts_log("❌ Lỗi: Không tìm thấy file script.")
            return

        # Lấy và xử lý phạm vi slide (NEW LOGIC)
        try:
            start_slide_str = self.tts_start_slide.get().strip()
            end_slide_str = self.tts_end_slide.get().strip()
            
            # Mặc định slide bắt đầu là 1, slide kết thúc là số rất lớn (hết)
            start_slide = int(start_slide_str) if start_slide_str.isdigit() else 1
            end_slide = int(end_slide_str) if end_slide_str.isdigit() else sys.maxsize 

            if start_slide <= 0:
                 raise ValueError("Slide bắt đầu phải lớn hơn 0.")
            if end_slide <= 0 and end_slide != sys.maxsize: # Chỉ kiểm tra nếu không phải giá trị mặc định
                 raise ValueError("Slide kết thúc phải lớn hơn 0.")
            if start_slide > end_slide:
                raise ValueError("Slide bắt đầu không thể lớn hơn Slide kết thúc.")

        except ValueError as e:
            self.tts_log_text.delete("1.0", ctk.END)
            self.tts_log(f"❌ Lỗi tham số Slide: {e}")
            return

        # Lấy cấu hình từ GUI
        voice_key = self.tts_voice.get()
        config = {
            "INPUT_FILE": input_file,
            "VOICE": self.VOICES.get(voice_key, 'en-US-GuyNeural'),
            "RATE": f"{self.tts_rate.get():+d}%",
            "PITCH": f"{self.tts_pitch.get():+d}Hz",
            "VOLUME": f"+{self.tts_volume.get()}%",
            "MAX_RETRIES": self.tts_max_retries.get(),
            "BASE_RETRY_SLEEP": 2.0,
            "BETWEEN_SLIDES_SLEEP": 0.8,
            # NEW CONFIGS
            "START_SLIDE": start_slide,
            "END_SLIDE": end_slide,
        }
        
        self.tts_log_text.delete("1.0", ctk.END)
        self.tts_log("▶️ Bắt đầu quá trình tạo audio...")
        slide_end_text = "Hết" if end_slide == sys.maxsize else end_slide
        self.tts_log(f"⚙️ Phạm vi Slide: Từ {start_slide} đến {slide_end_text}")
        self.tts_log(f"🗣️ Giọng đọc: {config['VOICE']} | Rate: {config['RATE']} | Pitch: {config['PITCH']}")
        
        # Chạy trong luồng để UI không bị treo
        threading.Thread(target=self._tts_generator_thread, args=(config,)).start()

    def _tts_generator_thread(self, config):
        """Hàm chính chạy bất đồng bộ cho TTS."""
        
        # Hàm parse_slides từ script gốc
        def parse_slides(text: str):
            text = text.lstrip("\ufeff").replace("\r\n", "\n").replace("\r", "\n")
            pattern = r"(Slide\s+\d+:\s*[^\n]*)(?:\n)(.*?)(?=\nSlide\s+\d+:|\Z)"
            matches = re.findall(pattern, text, flags=re.S)
            slides = []
            for idx, (title, body) in enumerate(matches, start=1):
                body = body.strip()
                if not body:
                    continue
                slides.append({"index": idx, "title": title.strip(), "body": body})
            return slides

        # Thiết lập output folder
        parent_directory = os.path.dirname(config['INPUT_FILE'])
        OUTPUT_FOLDER = os.path.join(parent_directory, "audio")
        
        os.makedirs(OUTPUT_FOLDER, exist_ok=True)
        self.tts_log(f"📁 Audio sẽ lưu tại: {OUTPUT_FOLDER}")
        
        async def tts_one(text: str, outfile: str):
            comm = edge_tts.Communicate(
                text,
                config['VOICE'],
                rate=config['RATE'],
                pitch=config['PITCH'],
                volume=config['VOLUME']
            )
            await comm.save(outfile)

        async def run_async_main():
            try:
                with open(config['INPUT_FILE'], "r", encoding="utf-8") as f:
                    raw = f.read()

                slides = parse_slides(raw)

                # --- APPLY SLIDE FILTERING (NEW LOGIC) ---
                start_slide = config['START_SLIDE']
                end_slide = config['END_SLIDE']
                
                filtered_slides = []
                for s in slides:
                    slide_no = s["index"]
                    if start_slide <= slide_no <= end_slide:
                        filtered_slides.append(s)
                
                if not filtered_slides:
                    self.tts_log(f"⚠️ Không tìm thấy slide nào trong phạm vi: {start_slide} đến {'Hết' if end_slide == sys.maxsize else end_slide}.")
                    return
                
                slides_to_process = filtered_slides
                # -----------------------------------

                self.tts_log(f"🔍 Đã tìm thấy {len(slides)} slide trong file script.")
                self.tts_log(f"✅ Sẽ xử lý {len(slides_to_process)} slide trong phạm vi đã chọn.")

                processed = 0
                failed = []

                for s in slides_to_process:
                    slide_no = s["index"]
                    text_chunk = s["body"]

                    self.tts_log(f"--- Đang xử lý Slide {slide_no}...")

                    outpath = os.path.join(OUTPUT_FOLDER, f"slide_{slide_no}.mp3")

                    ok = False
                    for attempt in range(1, config['MAX_RETRIES'] + 1):
                        try:
                            await tts_one(text_chunk, outpath)
                            self.tts_log(f"  ✅ Thành công sau {attempt} lần thử.")
                            ok = True
                            processed += 1
                            break
                        except Exception as e:
                            self.tts_log(f"  ⚠️ Lần thử {attempt}/{config['MAX_RETRIES']} thất bại: {repr(e)}")
                            if attempt < config['MAX_RETRIES']:
                                await asyncio.sleep(config['BASE_RETRY_SLEEP'] * (2 ** (attempt - 1)))

                    if not ok:
                        failed.append(slide_no)

                    await asyncio.sleep(config['BETWEEN_SLIDES_SLEEP'])

                self.tts_log("=" * 40)
                self.tts_log(f"🎉 Hoàn tất! Đã xử lý thành công {processed} file audio.")
                if failed:
                    self.tts_log(f"⚠️ Các slide sau bị lỗi: {failed}")
                
            except FileNotFoundError:
                self.tts_log(f"❌ Lỗi: Không tìm thấy file '{config['INPUT_FILE']}'.")
            except Exception as e:
                self.tts_log(f"❌ Đã xảy ra lỗi không mong muốn: {repr(e)}")

        # Chạy vòng lặp sự kiện asyncio
        try:
            loop = asyncio.new_event_loop()
            asyncio.set_event_loop(loop)
            loop.run_until_complete(run_async_main())
        except Exception as e:
            self.tts_log(f"❌ Lỗi Luồng Async: {repr(e)}")

    # =========================================================================================
    # --- TAB 3: RENAME SLIDE IMAGES ---
    # =========================================================================================

    def setup_rename_tab(self):
        tab = self.tabview.tab("3. Rename Slide Images")
        
        # Variables
        self.target_directory = ctk.StringVar(value="") 
        
        # --- Input Frame ---
        input_frame = ctk.CTkFrame(tab)
        input_frame.grid(row=0, column=0, padx=20, pady=10, sticky="ew")
        input_frame.grid_columnconfigure(1, weight=1)
        
        ctk.CTkLabel(input_frame, text="Thư mục chứa ảnh:").grid(row=0, column=0, padx=10, pady=5, sticky="w")
        ctk.CTkEntry(input_frame, textvariable=self.target_directory, width=400).grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        ctk.CTkButton(input_frame, text="Browse Folder", command=lambda: self.select_directory(self.target_directory)).grid(row=0, column=2, padx=10, pady=5)
        
        # --- Description & Action ---
        ctk.CTkLabel(tab, text="Quy tắc: Đổi tên file dạng 'SỐ_TÊN.ĐUÔI' thành 'slide-SỐ.ĐUÔI'", justify="left").grid(row=1, column=0, padx=20, pady=(10, 0), sticky="w")
        ctk.CTkLabel(tab, text="Ví dụ: '1_HinhAnh.png' -> 'slide-1.png'", justify="left").grid(row=2, column=0, padx=20, pady=(0, 10), sticky="w")
        
        self.rename_status_text = ctk.CTkTextbox(tab, height=200, activate_scrollbars=True, wrap="word")
        self.rename_status_text.grid(row=4, column=0, padx=20, pady=(10, 5), sticky="ew")
        
        ctk.CTkButton(tab, text="▶️ BẮT ĐẦU ĐỔI TÊN", command=self.run_rename).grid(row=3, column=0, padx=20, pady=20)


    def rename_log(self, message):
        """Hàm ghi log an toàn cho luồng."""
        self.rename_status_text.insert(ctk.END, message + "\n")
        self.rename_status_text.see(ctk.END)
        self.update_idletasks()

    def run_rename(self):
        """Khởi chạy đổi tên trong luồng riêng biệt."""
        target_dir = self.target_directory.get()
        self.rename_status_text.delete("1.0", ctk.END)

        if not os.path.isdir(target_dir):
            self.rename_log("❌ Lỗi: Thư mục không tồn tại. Vui lòng kiểm tra lại đường dẫn.")
            return

        self.rename_log(f"🔍 Bắt đầu quét thư mục: {target_dir}")
        self.update_idletasks()
        
        # Chạy trong luồng để UI không bị treo
        threading.Thread(target=self._rename_thread, args=(target_dir,)).start()

    def _rename_thread(self, target_dir):
        """Hàm xử lý đổi tên từ rename_files.py, điều chỉnh để log ra Textbox."""
        
        renamed_count = 0
        skipped_count = 0
        
        try:
            filenames = os.listdir(target_dir)
        except OSError as e:
            self.rename_log(f"❌ Lỗi: Không thể truy cập thư mục. Chi tiết: {e}")
            return

        for filename in filenames:
            # Tìm các file có dạng "số_tênfile.đuôi"
            match = re.match(r'^(\d+)_.*(\.\w+)$', filename)
            
            if match:
                number = match.group(1)
                extension = match.group(2)
                new_filename = f"slide-{number}{extension}"
                
                old_path = os.path.join(target_dir, filename)
                new_path = os.path.join(target_dir, new_filename)
                
                # Thực hiện đổi tên
                try:
                    os.rename(old_path, new_path)
                    self.rename_log(f"✅ Đã đổi tên: '{filename}'  ->  '{new_filename}'")
                    renamed_count += 1
                except OSError as e:
                    self.rename_log(f"❌ Lỗi khi đổi tên file '{filename}': {e}")
                    skipped_count += 1
            else:
                skipped_count += 1

        self.rename_log("-" * 40)
        self.rename_log("🎉 Hoàn tất!")
        self.rename_log(f"👍 Đã đổi tên thành công: {renamed_count} file.")
        self.rename_log(f"⏩ Đã bỏ qua: {skipped_count} file (không khớp định dạng).")


if __name__ == "__main__":
    # Fix cho PyInstaller/asyncio trên Windows
    if sys.platform == "win32":
        try:
            asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())
        except AttributeError:
            pass # Vòng lặp mặc định có thể đã ổn

    app = App()
    app.mainloop()