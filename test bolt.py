import customtkinter as ctk
import pandas as pd
from PIL import Image
import datetime
import os
import speech_recognition as sr 
import win32com.client         
import numpy as np 
from sklearn.feature_extraction.text import TfidfVectorizer 
from sklearn.naive_bayes import MultinomialNB 
from sklearn.pipeline import make_pipeline 
import joblib                   
import matplotlib.pyplot as plt 
from matplotlib.backends.backend_agg import FigureCanvasAgg as FigureCanvas
import io   

# --- Constants ---
CSV_FILE = "bolt_brain.csv"   
BOT_NAME = "⚡ Bolt"
MODEL_FILE = "bolt_intent_model.joblib"
MAX_IMG_W, MAX_IMG_H = 420, 220  


# -------------------- Voice (Windows SAPI) --------------------
speaker = win32com.client.Dispatch("SAPI.SpVoice")

try:
    voices = speaker.GetVoices()
    if voices.Count > 0:
        selected_voice = None
        for i in range(voices.Count):
            voice = voices.Item(i)
            # Prioritize Zira voice as requested
            if "Zira" in voice.GetDescription() or "Microsoft Zira" in voice.GetDescription():
                selected_voice = voice
                break
        
        if selected_voice:
            speaker.Voice = selected_voice
            print("Voice set to Zira.")
        else:
            speaker.Voice = voices.Item(0)
            print("Zira voice not found. Using default voice.")

    speaker.Rate = 0  
    speaker.Volume = 100 
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
        text = recognizer.recognize_google(audio)
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

# -------------------- Data Management & Self-Learning --------------------
qa_map = {} 
def load_csv_data():
    """Loads or reloads the Q&A data from the CSV file."""
    global qa_map
    qa_map = {}
    if os.path.exists(CSV_FILE):
        try:
            df = pd.read_csv(CSV_FILE)
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

load_csv_data()


def cache_answer(question: str, answer: str):
    """Writes the new Q&A pair to the bolt_brain.csv file (Bolt is learning!)."""
    try:
        new_row = pd.DataFrame([{"question": question, "answer": answer, "image": ""}], 
                               columns=["question", "answer", "image"])
        
        if os.path.exists(CSV_FILE):
            existing_df = pd.read_csv(CSV_FILE)
            updated_df = pd.concat([existing_df, new_row], ignore_index=True)
        else:
            updated_df = new_row
            
        updated_df.to_csv(CSV_FILE, index=False)
        load_csv_data()
        print(f"[{BOT_NAME}]: Successfully cached new knowledge.")
        
    except Exception as e:
        print(f"[{BOT_NAME}]: Error saving to CSV:", e)


# --- NEW: Google Search Function (API Call) ---
def search_internet_for_answer(query: str):
    """
    Calls the Google Search tool to find a relevant answer snippet.
    Returns a simplified list of results (snippet, source_title) or None.
    """
    
    # This is the tool invocation that gets real-time data
    search_results = google.search(queries=[query])
    
    formatted_results = []
    if search_results:
        for result in search_results:
            formatted_results.append({
                "source_title": result.get('source_title', 'Internet Source'),
                "snippet": result.get('snippet', 'No detailed snippet available.'),
                "url": result.get('url', '#')
            })
    return formatted_results


# -------------------- AI Intent Classification Setup --------------------
TRAINING_DATA = [
    ("hi there, bolt", "greeting"), ("hello", "greeting"), ("good morning", "greeting"), ("hey bolt", "greeting"), ("hlo", "greeting"),
    ("thank you so much", "thanks"), ("thanks a lot", "thanks"), ("tq", "thanks"),
    ("bye now", "farewell"), ("goodbye", "farewell"), ("see you later", "farewell"),
    ("what is the current time", "time"), ("time now", "time"), ("tell me the time", "time"),
    ("what is the name of our school", "school_info"), ("tell me about srmjv", "school_info"), 
    ("who is the correspondent", "correspondent"), 
    ("who are you", "bot_info"), ("what is your name", "bot_info"), 
    ("what is the capital of india", "general_fact"), ("what is photosynthesis", "science_fact"), 
]
X_train = [q.lower() for q, intent in TRAINING_DATA]
y_train = [intent for q, intent in TRAINING_DATA]


def get_intent_model():
    """Loads the model if it exists, otherwise trains and saves a new one."""
    if os.path.exists(MODEL_FILE):
        return joblib.load(MODEL_FILE)
    
    model = make_pipeline(TfidfVectorizer(), MultinomialNB())
    model.fit(X_train, y_train)
    
    joblib.dump(model, MODEL_FILE)
    return model

intent_classifier = get_intent_model()

def predict_intent(user_text: str):
    """Predicts the most likely intent tag and returns its confidence."""
    if not user_text.strip():
        return "fallback", 0.0
    
    text_to_predict = [user_text.lower()]
    intent = intent_classifier.predict(text_to_predict)[0]
    probabilities = intent_classifier.predict_proba(text_to_predict)[0]
    max_prob = np.max(probabilities)
    
    CONFIDENCE_THRESHOLD = 0.60 
    
    if max_prob < CONFIDENCE_THRESHOLD:
        return "fallback", 0.0
    
    return intent, max_prob

# -------------------- Matplotlib Visualization --------------------
def create_confidence_chart(intent, confidence):
    """Generates a Matplotlib bar chart showing the model's confidence."""
    plt.style.use('dark_background') 
    
    fig, ax = plt.subplots(figsize=(4.5, 0.6), facecolor='#1a1a1a') 
    
    intents = [f"Intent: {intent}"]
    scores = [confidence]
    
    ax.barh(intents, scores, color=['#2ECC71'], height=0.5) 
    
    ax.set_xlim(0, 1.0)
    ax.set_xticks(np.arange(0, 1.1, 0.25))
    ax.tick_params(axis='x', colors='gray', labelsize=8)
    ax.tick_params(axis='y', colors='white', labelsize=10)
    
    ax.text(confidence + 0.01, 0, f'{confidence*100:.1f}%', va='center', color='white', fontsize=9, fontweight='bold')
    
    ax.spines['bottom'].set_color('gray')
    ax.spines['left'].set_color('gray')
    ax.spines['right'].set_visible(False)
    ax.spines['top'].set_visible(False)
    
    plt.tight_layout()
    buf = io.BytesIO()
    FigureCanvas(fig).print_png(buf)
    plt.close(fig) 
    
    buf.seek(0)
    return buf


# -------------------- Core Logic: The Brain --------------------

def handle_intent(user_text: str):
    """Main function: Intent -> CSV Lookup -> Self-Learn (API) -> Fallback"""
    
    intent, confidence = predict_intent(user_text)
    chart_buffer = None
    default_fallback = "Sorry, I am still learning! I don't have a specific answer for that. Please ask a factual query I've been taught."
    ans = default_fallback
    img_path = None
    
    # 1. Intent Classification Response (if high confidence)
    if intent != "fallback":
        chart_buffer = create_confidence_chart(intent, confidence)
        
        if intent == "greeting":
            ans = "Hello! I am Bolt, your AI knowledge assistant. How can I assist you today?"
        elif intent == "thanks":
            ans = "You're most welcome! Do you have any other questions?"
        elif intent == "farewell":
            ans = "Goodbye! Remember to stay fit and focused on your studies, Captain America fan!"
        elif intent == "time":
            now = datetime.datetime.now().strftime("%I:%M %p")
            ans = f"The current time is {now}"
        elif intent == "school_info":
            ans = "Our school is the Sri R.M Jain Vidhyapeeth Matriculation Higher Secondary School, dedicated to quality education."
        elif intent == "correspondent":
            ans = "Our Correspondent is Thiru. N. M. K. K. Ponnusamy, a visionary leader for our school."
        elif intent == "bot_info":
            ans = f"I am BOLT , an AI assistant created by a student sarvin to help the school. I learn from the 'bolt_brain.csv' file and use machine learning to understand you."

    # 2. CSV Lookup (Primary Knowledge Source)
    t = user_text.strip().lower()
    if t in qa_map:
        ans, img_path = qa_map[t]
        chart_buffer = None 

    # --- 3. Self-Learning (Live Internet Search) ---
    if ans == default_fallback: 
        print(f"[{BOT_NAME}]: Initiating Internet Search for: {user_text}")
        
        search_results = search_internet_for_answer(user_text)
        
        if search_results:
            best_result = search_results[0]
            
            # Format the answer for the user
            found_snippet = best_result['snippet']
            new_answer = f"[Learned from the Web] '{found_snippet}' (Source: {best_result['source_title']}). I have saved this knowledge permanently."
            
            # Cache the answer for future use (Bolt is learning!)
            cache_answer(user_text, new_answer) 
            ans = new_answer
            chart_buffer = None 
        else:
            # Internet search failed too, keep the original default fallback
            pass 

    # 4. Final Return
    if chart_buffer is not None:
        return ans, chart_buffer
    
    return ans, img_path


# -------------------- GUI (customtkinter) --------------------
class BoltApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("green")
        self.title(f"{BOT_NAME} - AI Assistant")
        self.geometry("780x680") 
        self.minsize(720, 580)

        # Header
        self.header = ctk.CTkLabel(self, text=f"BOLT: AI Assistant", 
                                   font=("Bahnschrift SemiBold", 30, "bold"))
        self.header.pack(pady=(20, 10))

        # Chat Log Frame
        self.chat_frame = ctk.CTkScrollableFrame(self, label_text="Chat Log", 
                                                 width=700, height=450, 
                                                 fg_color="#1a1a1a", label_font=("Consolas", 14, "bold"))
        self.chat_frame.pack(pady=(6, 12), padx=20, fill="x", expand=True)
        self.chat_frame.grid_columnconfigure(0, weight=1)
        
        # Message entry
        self.entry = ctk.CTkEntry(
            self, placeholder_text="Ask me a question or challenge my knowledge (online!)",
            width=560, height=44, corner_radius=25, font=("Consolas", 16),
        )
        self.entry.pack(pady=(6, 12))
        self.entry.bind("<Return>", self._on_ask)

        # Buttons row
        btn_row = ctk.CTkFrame(self, fg_color="transparent")
        btn_row.pack(pady=(0, 20))

        self.listen_btn = ctk.CTkButton(
            btn_row, text="🎤 Ask by Voice (online)", width=180, height=56, 
            font=("Consolas", 16), command=self._on_listen,
        )
        self.listen_btn.grid(row=0, column=0, padx=12)

        self.speak_btn = ctk.CTkButton(
            btn_row, text="🔊 Repeat Answer", width=180, height=56, 
            font=("Consolas", 16), command=self._on_speak,
        )
        self.speak_btn.grid(row=0, column=1, padx=12)

        self.ask_btn = ctk.CTkButton(
            btn_row, text="Ask", width=120, height=56, 
            font=("Consolas", 18, "bold"), command=self._on_ask,
        )
        self.ask_btn.grid(row=0, column=2, padx=12)

        self.last_answer = ""
        self._blank_img = ctk.CTkImage(Image.new("RGB", (1, 1), color=(0, 0, 0)), size=(1, 1))

        self._add_message(f"I'm {BOT_NAME}. I now learn new facts dynamically using Google Search and Machine Learning!", is_bot=True)


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
            self._add_message("I couldn't hear you clearly.", is_bot=True)

    def _on_speak(self):
        if self.last_answer:
            speak(self.last_answer)

    def _process(self, q: str):
        self._add_message(q, is_bot=False) 
        
        ans, img_source = handle_intent(q) 
        self.last_answer = ans
        
        self._add_message(ans, is_bot=True, img_source=img_source)
        
        speak(ans)
        self.entry.delete(0, ctk.END)

    def _add_message(self, text, is_bot=False, img_source=None):
        if is_bot:
            color = "lime"
            bg_color = "#0e4d00" 
            anchor = "w"
        else:
            color = "#00F0FF" 
            bg_color = "#004d4d" 
            anchor = "e"
            
        msg_frame = ctk.CTkFrame(self.chat_frame, fg_color=bg_color, corner_radius=12)
        
        msg_label = ctk.CTkLabel(msg_frame, text=text, text_color=color, 
                                 font=("Consolas", 14), wraplength=550, justify="left")
        msg_label.pack(padx=10, pady=8)
        
        # Handle Matplotlib Chart Buffer
        if isinstance(img_source, io.BytesIO): 
            img_plt = Image.open(img_source)
            ctk_img = ctk.CTkImage(light_image=img_plt, dark_image=img_plt, size=img_plt.size)
            
            img_label = ctk.CTkLabel(msg_frame, text="AI Confidence:", font=("Consolas", 12, "bold"), text_color="yellow")
            img_label.pack(padx=10, pady=(0, 2))
            
            chart_label = ctk.CTkLabel(msg_frame, text="", image=ctk_img)
            chart_label.image = ctk_img 
            chart_label.pack(padx=10, pady=(0, 8))
            
        # Handle Image File Path from CSV
        elif img_source and isinstance(img_source, str): 
            ctk_img, self._img_ref = self._load_image_ctk(img_source) 
            img_label = ctk.CTkLabel(msg_frame, text="", image=ctk_img)
            img_label.image = ctk_img 
            img_label.pack(padx=10, pady=(0, 8))
        
        row_num = self.chat_frame.grid_size()[1]
        self.chat_frame.grid_rowconfigure(row_num, weight=0) 
        msg_frame.grid(row=row_num, column=0, sticky=anchor, padx=10, pady=5)
        
        self.chat_frame.after(100, lambda: self.chat_frame._parent_canvas.yview_moveto(1.0))
        
    def _load_image_ctk(self, path: str):
        full = path
        if not os.path.exists(full):
            alt = os.path.join("image", path)
            full = alt if os.path.exists(alt) else ""

        if not full or not os.path.exists(full):
            return self._blank_img, self._blank_img

        try:
            im = Image.open(full).convert("RGB") 
            w, h = im.size
            scale = min(MAX_IMG_W / w, MAX_IMG_H / h, 1.0)
            new_size = (int(w * scale), int(h * scale))
            im_resized = im.resize(new_size, Image.LANCZOS)

            ctk_img = ctk.CTkImage(
                light_image=im_resized, dark_image=im_resized, size=new_size
            )
            return ctk_img, ctk_img
        except Exception as e:
            print("Image load error:", e)
            return self._blank_img, self._blank_img


# -------------------- Run --------------------
if __name__ == "__main__":
    plt.style.use('dark_background')
    app = BoltApp()
    app.mainloop()
