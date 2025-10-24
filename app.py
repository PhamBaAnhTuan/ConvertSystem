import customtkinter as ctk
import os
from video_builder import VideoBuilderTool
from markdown_to_docx import MarkdownToDocxConverter
from tts_google import TextToSpeechGoogle
from tts_google_v2 import TextToSpeechToolV2
from tts_pyx3 import TextToSpeechPyx3


class MultiToolApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Auto Teaching Tools - MultiTool")
        self.geometry("1000x650")

        self.tabview = ctk.CTkTabview(self)
        self.tabview.pack(
            fill="both",
            expand=True,
            padx=20,
            pady=20,
        )

        self.tab_md = self.tabview.add("Markdown → DOCX")
        self.tab_text_to_speech_v2 = self.tabview.add("Text → Speech Google TTS v2")
        self.tab_text_to_speech_gg = self.tabview.add("Text → Speech Google TTS")
        self.tab_text_to_speech_pyx3 = self.tabview.add("Text → Speech Pyx3 TTS")
        self.tab_video = self.tabview.add("Image + Audio → MP4")

        MarkdownToDocxConverter(master=self.tab_md)
        TextToSpeechToolV2(master=self.tab_text_to_speech_v2)
        TextToSpeechGoogle(master=self.tab_text_to_speech_gg)
        TextToSpeechPyx3(master=self.tab_text_to_speech_pyx3)
        VideoBuilderTool(master=self.tab_video)


if __name__ == "__main__":
    app = MultiToolApp()
    app.mainloop()
