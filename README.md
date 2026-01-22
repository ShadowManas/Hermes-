# HERMYS – Smart Email System

[cite_start]Hermys is an intelligent, user-centric email assistant designed to manage emails more efficiently by combining functional design with AI-powered automation[cite: 26, 42]. [cite_start]Built using Python and PyQt5, Hermys is a smart, voice-enabled assistant that enhances how users interact with their inbox through Artificial Intelligence and Natural Language Processing (NLP)[cite: 27].

## 🚀 Executive Summary

[cite_start]Hermys aims to overcome challenges like information overload and tedious navigation by offering a smart desktop assistant for email handling[cite: 61, 72]. [cite_start]It automates repetitive actions, interprets email context, and provides voice-activated controls[cite: 74, 75, 76]. [cite_start]The system does not require a database, handling all operations live using IMAP, SMTP, and Gemini’s APIs[cite: 41].

## ✨ Key Features

### 📨 Intelligent Email Dashboard
* [cite_start]**Categorized Views:** View lists of Inbox, Sent, and Spam emails with real-time fetching via IMAP[cite: 31].
* [cite_start]**Smart Viewer:** Read selected emails with options to immediately forward or delete them[cite: 32].
* [cite_start]**Rich Composer:** Auto-fills sender/recipient fields, supports CC/BCC, attachments, rich formatting, and message discarding[cite: 33].

### 🧠 NLP Analysis Panel
[cite_start]Real-time analysis of email content using tools like TextBlob, spaCy, VADER, and Sumy[cite: 34].
* [cite_start]**Intent Recognition:** Automatically detects what the email is about[cite: 171].
* [cite_start]**Summarization:** Provides a short summary of the email content[cite: 172].
* [cite_start]**Sentiment Analysis:** Detects positive, negative, or neutral tones[cite: 173].
* [cite_start]**Keyword Extraction:** Highlights important keywords used in the message[cite: 174].

### 🤖 Assistant & Insights
* [cite_start]**Voice Control:** Issue commands to open, forward, delete, or read emails aloud using speech-to-text[cite: 35, 38].
* [cite_start]**AI Reply Generation:** Generates and schedules intelligent replies using the Gemini API[cite: 36, 126].
* [cite_start]**Reminders:** Set reminders for email follow-ups[cite: 36].
* [cite_start]**Visual Analytics:** Toggle dynamic graphs showing email volume by day, sentiment trends, keyword usage, and top senders[cite: 39].

## 🛠️ Technology Stack

* [cite_start]**Programming Language:** Python[cite: 88].
* [cite_start]**GUI Framework:** PyQt5 (Supports resizable panels and Dark/Light mode)[cite: 89, 106].
* [cite_start]**AI & ML:** Google Gemini via GenAI API[cite: 91].
* [cite_start]**NLP Libraries:** NLTK, spaCy, TextBlob, Sumy, VADER[cite: 90].
* [cite_start]**Protocols:** SMTP (Sending), IMAP (Fetching)[cite: 88].

## 📸 Screenshots

### Login Screen
[cite_start]*The system requires secure authentication via Email ID, App Password, and Gemini API Key[cite: 28].*
![Login Screen](https://via.placeholder.com/800x600?text=Login+Screen+Snapshot)  
*(Reference: Source 186)*

### AI Email Dashboard & NLP Analysis
[cite_start]*A five-panel interactive dashboard displaying the Inbox, Viewer, Composer, NLP Analysis, and Assistant[cite: 29].*
![Dashboard](https://via.placeholder.com/800x600?text=Dashboard+and+NLP+Analysis)  
*(Reference: Source 186)*

### Analytics & Dark Mode
[cite_start]*Visual data representation including Email Volume and Sentiment Trends in Dark Mode[cite: 180, 186].*
![Analytics](https://via.placeholder.com/800x600?text=Analytics+and+Dark+Mode)  
*(Reference: Source 186)*

## ⚙️ Setup & Installation

1.  **Clone the Repository:**
    ```bash
    git clone [https://github.com/yourusername/hermys.git](https://github.com/yourusername/hermys.git)
    ```
2.  **Install Dependencies:**
    [cite_start]Ensure you have the required NLP libraries installed (NLTK, spaCy, etc.)[cite: 90].
    ```bash
    pip install -r requirements.txt
    ```
3.  **Credentials Required:**
    To log in, you will need:
    * **Email Address**
    * [cite_start]**App Password:** (Not your standard email password. Enable 2FA on your email account and generate an App Password)[cite: 28].
    * [cite_start]**Gemini API Key:** Required for AI generation and NLP features[cite: 28].

## 🗣️ Voice Commands
[cite_start]Hermys supports voice commands for hands-free operation[cite: 129]. Examples include:
* [cite_start]"Read emails from 1 to 3"[cite: 188].
* [cite_start]"Forward this email to John"[cite: 188].
* [cite_start]"Generate reply to this email and send after 1 min"[cite: 188].

## 🔮 Future Enhancements
* [cite_start]**Database Integration:** Adding SQLite to store login history and user preferences[cite: 206].
* [cite_start]**Full Threading:** Implementing email threading for better conversation organization[cite: 207].
* [cite_start]**Mobile Support:** Developing a mobile app version with cloud sync[cite: 218].
* [cite_start]**Advanced LLMs:** Integrating larger models for deeper contextual understanding[cite: 210].

## 👥 Author
[cite_start]**Manas M Surve** [cite: 7]  
[cite_start]*Navinchandra Mehta Institute of Technology & Development (NMITD)* [cite: 9]  
[cite_start]*Deccan Education Society* [cite: 8]

---
[cite_start]*This project was submitted as a Mini Project for Semester I of the MCA program (2024-2025).* [cite: 11, 12]
