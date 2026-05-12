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

📁 File Structure
main_gui.py: The central hub/dashboard for navigating between Bolt and other academic tools.

bolt.py: The primary engine handling the chat interface, voice output, and image rendering.

bolt_brain.csv: The "Brain" of the project containing verified questions, answers, and image paths.

test bolt.py: Experimental script for testing intent recognition and advanced features.

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
