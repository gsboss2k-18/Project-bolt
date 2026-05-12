# Project-bolt
Project Bolt is a specialized, condition-based AI chatbot designed for high accuracy and offline utility. Unlike generative AI models that can hallucinate, Bolt provides precise information based on a curated knowledge base, making it ideal for educational environments like schools and colleges.

🚀 Overview
Project Bolt features a custom graphical interface and voice integration. It is built to be lightweight, fast, and completely free-to-use, aligning with the "1% Better Every Day" philosophy of continuous improvement.

🛠️ Technical Architecture
Core Logic: Condition-based response system utilizing pandas for data management.

GUI: Developed with CustomTkinter for a modern, dark-themed dashboard.

Voice: Integrated text-to-speech using win32com.client (SAPI) with a preference for the "Zira" voice profile.

Intelligence: Includes intent classification experiments using scikit-learn (Tfidf & Naive Bayes).

Data: Powered by bolt_brain.csv, a customizable knowledge base.

HOW TO INSTALL DEPENDENCIES:
FOR WINDOWS

             pip install customtkinter pandas pillow SpeechRecognition pywin32 scikit-learn joblib matplotlib

FOR LINUX

              . Update your System
First, make sure your package list is up to date:

Bash
sudo apt update
2. Install System Dependencies
Some Python libraries (like SpeechRecognition and Pillow) need Linux system tools to work properly:

Bash
sudo apt install python3-pip python3-tk python3-pil.imagetk portaudio19-dev python3-pyaudio
3. Install the Python Libraries
Now, use pip to install the core libraries you used in your code:

Bash
pip install customtkinter pandas pillow SpeechRecognition scikit-learn joblib matplotlib
4. The Linux "Voice" Fix
Your current code uses win32com.client for the "Zira" voice, which cannot be installed on Linux. To make Bolt speak on your Linux laptop, you should install pyttsx3:

Bash
pip install pyttsx3
Why? pyttsx3 is cross-platform. It will use the Windows voice on Windows and the Linux "espeak" voice on your Xubuntu setup automatically.

             
🎯 Key Features
Zero Hallucination: Responds only with verified data from the knowledge base.

Offline Capable: Works without an internet connection.

Multimodal: Supports text, voice recognition, and image display.

Low Cost: Designed to run on accessible hardware (like the Lenovo E41-25).

📈 Future Goals
[ ] Expand the bolt_brain.csv database for 10th Standard TN State Board syllabus.

[ ] Improve voice input accuracy.

[ ] Implement self-learning capabilities.

[ ] Add multi-language support.

🧠 About the Developer
Developed by G S Boss, an aspiring GPU Infrastructure Manager and AI Research Engineer.

"Focused on creating AGI and mastering Linux terminal one step at a time."
