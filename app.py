import customtkinter as ctk
import os
from mp4_maker import VideoMakerApp
from markdown_to_docx import MarkdownToDocxConverter
from tts_google import TextToSpeechGoogle
from tts_pyx3 import TextToSpeechPyx3


class MultiToolApp(ctk.CTk):
    """Ứng dụng chính có thể chứa nhiều tool"""

    def __init__(self):
        super().__init__()
        self.title("Auto Teaching Tools - MultiTool")
        self.geometry("1000x700")

        # Tạo tabview để chứa nhiều tool
        self.tabview = ctk.CTkTabview(self)
        self.tabview.pack(
            fill="both",
            expand=True,
            padx=20,
            pady=20,
        )

        # Thêm các tab
        self.tab_md = self.tabview.add("Markdown → DOCX")
        self.tab_video = self.tabview.add("Image + Audio → MP4")
        self.tab_text_to_speech_gg = self.tabview.add("Text → Speech Google TTS")
        self.tab_text_to_speech_pyx3 = self.tabview.add("Text → Speech Pyx3 TTS")

        # Gắn các component vào từng tab
        MarkdownToDocxConverter(master=self.tab_md)
        VideoMakerApp(master=self.tab_video)
        TextToSpeechGoogle(master=self.tab_text_to_speech_gg)
        TextToSpeechPyx3(master=self.tab_text_to_speech_pyx3)


if __name__ == "__main__":
    app = MultiToolApp()
    app.mainloop()
