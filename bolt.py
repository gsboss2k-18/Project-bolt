# bolt.py — SRMJV GUI + voice (win32com) + safe image handling

import customtkinter as ctk
import pandas as pd
from PIL import Image
import datetime
import os
import speech_recognition as sr  # Google Speech Recognition
import win32com.client

CSV_FILE = "bolt_brain.csv"   # columns: question, answer, image (or images)
BOT_NAME = "⚡ Bolt"
MAX_IMG_W, MAX_IMG_H = 420, 220   # auto-resize bounds for answer image

# -------------------- Voice (Windows SAPI) --------------------
speaker = win32com.client.Dispatch("SAPI.SpVoice")

try:
    voices = speaker.GetVoices()
    if voices.Count > 0:
        selected_voice = None
        for i in range(voices.Count):
            voice = voices.Item(i)
            if "Zira" in voice.GetDescription():
                selected_voice = voice
                break
        if selected_voice:
            speaker.Voice = selected_voice
        else:
            speaker.Voice = voices.Item(0)

    speaker.Rate = 0  # You can adjust speaking rate
    speaker.Volume = 100  # Volume level 0-100
except Exception as e:
    print("Voice initialization error:", e)

def speak(text: str):
    if text and text.strip():
        try:
            speaker.Speak(text)
        except Exception as e:
            print("Speak error:", e)

# -------------------- Google Voice Input --------------------
recognizer = sr.Recognizer()
mic = sr.Microphone()

def listen_google():
    try:
        with mic as source:
            print("Listening...")
            recognizer.adjust_for_ambient_noise(source)
            audio = recognizer.listen(source, timeout=5, phrase_time_limit=8)
        text = recognizer.recognize_google(audio)  # send to Google API
        print("You said:", text)
        return text
    except sr.UnknownValueError:
        return ""
    except sr.RequestError as e:
        print("Google API error:", e)
        return ""
    except Exception as e:
        print("Voice input error:", e)
        return ""

# -------------------- Load CSV into memory --------------------
qa_map = {}  # question (lower) -> (answer, image_path or "")
if os.path.exists(CSV_FILE):
    try:
        df = pd.read_csv(CSV_FILE)
        # case-insensitive column handling
        col = {c.lower(): c for c in df.columns}
        qcol = col.get("question")
        acol = col.get("answer")
        icol = col.get("image") if "image" in col else col.get("images")
        if qcol and acol:
            for _, row in df.iterrows():
                q = str(row[qcol]).strip().lower()
                a = str(row[acol]).strip()
                img = ""
                if icol and pd.notna(row[icol]):
                    img = str(row[icol]).strip()
                qa_map[q] = (a, img)
    except Exception as e:
        print("CSV load error:", e)

def find_answer(user_text: str):
    t = user_text.strip().lower()
    # special: tell time
    if "time" in t:
        now = datetime.datetime.now().strftime("%I:%M %p")
        return f"The current time is {now}", ""
    # exact match in CSV
    if t in qa_map:
        return qa_map[t][0], qa_map[t][1]
    return "Sorry, I don’t know that yet.", ""

# -------------------- GUI --------------------
class BoltApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("green")
        self.title("Bolt")
        self.geometry("780x520")
        self.minsize(720, 480)

        # Header
        self.header = ctk.CTkLabel(self, text="Sri R.M Jain Vidhyapeeth matric Hr sec. school", font=("Bahnschrift SemiBold", 28))
        self.header.pack(pady=(20, 10))

        # Message entry
        self.entry = ctk.CTkEntry(
            self,
            placeholder_text="Type your question here...",
            width=560,
            height=44,
            corner_radius=25,
            font=("Consolas", 16),
        )
        self.entry.pack(pady=(6, 12))
        self.entry.bind("<Return>", self._on_ask)

        # Answer card
        self.answer_card = ctk.CTkFrame(self, corner_radius=18, fg_color="#1a1a1a")
        self.answer_card.pack(pady=6, ipadx=12, ipady=12)

        self.answer_text = ctk.CTkLabel(
            self.answer_card,
            text="Answer will appear here",
            font=("Consolas", 16),
            text_color="lime",
            justify="center",
        )
        self.answer_text.pack(padx=18, pady=(16, 10))

        self.answer_img = ctk.CTkLabel(self.answer_card, text="")
        self.answer_img.pack(padx=12, pady=(0, 12))

        # --- blank fallback image (1x1 black pixel) ---
        self._blank_img = ctk.CTkImage(Image.new("RGB", (1, 1), color=(0, 0, 0)), size=(1, 1))
        self.answer_img.configure(image=self._blank_img, text="")
        self.answer_img.image = self._blank_img
        self._img_ref = self._blank_img

        # Buttons row
        btn_row = ctk.CTkFrame(self, fg_color="transparent")
        btn_row.pack(pady=20)

        self.listen_btn = ctk.CTkButton(
            btn_row,
            text="🎤 Ask by Voice(online)",
            width=180,
            height=56,
            corner_radius=30,
            font=("Consolas", 16),
            command=self._on_listen,
        )
        self.listen_btn.grid(row=0, column=0, padx=12)

        self.speak_btn = ctk.CTkButton(
            btn_row,
            text="🔊 Repeat Answer",
            width=180,
            height=56,
            corner_radius=30,
            font=("Consolas", 16),
            command=self._on_speak,
        )
        self.speak_btn.grid(row=0, column=1, padx=12)

        self.ask_btn = ctk.CTkButton(
            btn_row,
            text="Ask",
            width=120,
            height=44,
            corner_radius=22,
            font=("Consolas", 14),
            command=self._on_ask,
        )
        self.ask_btn.grid(row=0, column=2, padx=12)

        self.last_answer = ""

    # --------- Actions ---------
    def _on_ask(self, event=None):
        q = self.entry.get().strip()
        if not q:
            return
        self._process(q)

    def _on_listen(self):
        q = listen_google()
        if q:
            self.entry.delete(0, ctk.END)
            self.entry.insert(0, q)
            self._process(q)
        else:
            self.answer_text.configure(text="I couldn't hear you clearly.")

    def _on_speak(self):
        if self.last_answer:
            speak(self.last_answer)

    def _process(self, q: str):
        ans, img_path = find_answer(q)
        self.last_answer = ans
        self.answer_text.configure(text=ans)
        self._show_image(img_path)
        speak(ans)

    # --------- Image helper (supports JPG + PNG) ---------
    def _show_image(self, path: str):
        if not path:
            # use fallback instead of None
            self.answer_img.configure(image=self._blank_img, text="")
            self.answer_img.image = self._blank_img
            self._img_ref = self._blank_img
            return

        full = path
        if not os.path.exists(full):
            alt = os.path.join("image", path)
            full = alt if os.path.exists(alt) else ""

        if not full or not os.path.exists(full):
            # fallback
            self.answer_img.configure(image=self._blank_img, text="")
            self.answer_img.image = self._blank_img
            self._img_ref = self._blank_img
            return

        try:
            im = Image.open(full).convert("RGB")  # jpg/jpeg/png safe
            w, h = im.size
            scale = min(MAX_IMG_W / w, MAX_IMG_H / h, 1.0)
            new_size = (int(w * scale), int(h * scale))
            im_resized = im.resize(new_size, Image.LANCZOS)

            ctk_img = ctk.CTkImage(
                light_image=im_resized, dark_image=im_resized, size=new_size
            )

            self.answer_img.configure(image=ctk_img, text="")
            self.answer_img.image = ctk_img
            self._img_ref = ctk_img
        except Exception as e:
            print("Image error:", e)
            self.answer_img.configure(image=self._blank_img, text="")
            self.answer_img.image = self._blank_img
            self._img_ref = self._blank_img

# -------------------- Run --------------------
if __name__ == "__main__":
    app = BoltApp()
    app.mainloop()
