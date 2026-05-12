import customtkinter as ctk
import subprocess

class Dashboard(ctk.CTk):
    def __init__(self):
        super().__init__()
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("green")
        self.geometry("700x400")
        self.title("DASH BOARD")
        self.minsize(680, 370)

        # Header
        self.header = ctk.CTkLabel(
            self,
            text="DASH BOARD",
            font=("Rockwell Extra Bold", 38),
            anchor="center",
        )
        self.header.pack(pady=(30, 40))

        # Buttons container frame centered horizontally
        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.pack(pady=10, padx=20)
        btn_frame.grid_columnconfigure((0,1), weight=1)

        # Left column buttons
        self.bolt_btn = ctk.CTkButton(
            btn_frame,
            text="BOLT",
            font=("Rockwell Extra Bold", 20, "bold"),
            width=280,
            height=52,
            corner_radius=22,
            command=lambda: self.open_script("bolt.py"),
        )
        self.bolt_btn.grid(row=0, column=0, padx=12, pady=12)

        self.social_btn = ctk.CTkButton(
            btn_frame,
            text="SOCIAL",
            font=("Rockwell Extra Bold", 20, "bold"),
            width=280,
            height=52,
            corner_radius=22,
            command=lambda: self.open_script("social.py"),
        )
        self.social_btn.grid(row=1, column=0, padx=12, pady=12)

        # Right column buttons
        self.grammar_btn = ctk.CTkButton(
            btn_frame,
            text="ENGLISH GRAMMAR",
            font=("Rockwell Extra Bold", 20, "bold"),
            width=280,
            height=52,
            corner_radius=22,
            command=lambda: self.open_script("eng gra.py"),
        )
        self.grammar_btn.grid(row=0, column=1, padx=12, pady=12)

        self.poets_btn = ctk.CTkButton(
            btn_frame,
            text="ENGLISH POETS AND AUTHORS",
            font=("Rockwell Extra Bold", 20, "bold"),
            width=280,
            height=52,
            corner_radius=22,
            command=lambda: self.open_script("eng pa.py"),
        )
        self.poets_btn.grid(row=1, column=1, padx=12, pady=12)

    def open_script(self, script):
        try:
            subprocess.Popen(["python", script])
        except Exception as e:
            ctk.CTkMessagebox.show_error(title="Error", message=f"Failed to open {script}\n{e}")

if __name__ == "__main__":
    app = Dashboard()
    app.mainloop()
