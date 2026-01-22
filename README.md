# 🦅 HERMYS – Smart Email System

**Hermys** is a smart, voice-enabled desktop email assistant built using **Python** and **PyQt5**. It is designed to enhance productivity by integrating Artificial Intelligence and Natural Language Processing (NLP) into your daily email workflow. Hermys automates repetitive tasks, provides context-aware insights, and allows for hands-free voice control.

---

## 📋 Table of Contents
- [Executive Summary](#-executive-summary)
- [Key Features](#-key-features)
- [Tech Stack](#-tech-stack)
- [Screenshots](#-screenshots)
- [Installation & Setup](#-installation--setup)
- [Voice Commands](#-voice-commands)
- [Future Enhancements](#-future-enhancements)
- [Author](#-author)

---

## 🚀 Executive Summary
Hermys overcomes the challenges of traditional email clients—such as information overload and tedious navigation—by offering an intelligent assistant. It operates without a local database, handling all operations live via **IMAP**, **SMTP**, and **Google's Gemini API**.

The system is designed for students, professionals, and tech-savvy users who want to:
* Automate repetitive email actions.
* Understand email context instantly via NLP.
* Respond intelligently with minimal effort.

---

## ✨ Key Features

### 📨 Smart Dashboard & Viewer
* **Live Fetching:** View categorized lists (Inbox, Sent, Spam) fetched in real-time via IMAP.
* **Quick Actions:** Read, forward, or delete emails directly from the viewer.
* **Search:** Quickly find specific emails by keywords.

### ✍️ Intelligent Composer
* **Rich Formatting:** Supports attachments, CC/BCC, and rich text.
* **AI Replies:** Generates professional replies automatically using the Gemini API.
* **Scheduling:** Schedule replies to be sent at a later time.

### 🧠 NLP Analysis Panel
Instantly analyzes the content of opened emails using **TextBlob**, **spaCy**, **VADER**, and **Sumy**:
* **Intent Recognition:** Detects if an email is a warning, inquiry, complaint, etc.
* **Summarization:** Provides a concise summary of long emails.
* **Sentiment Analysis:** Identifies the tone (Positive, Negative, Neutral) and emotion.
* **Keyword Extraction:** Highlights key entities and topics.

### 🤖 Voice Assistant
* **Hands-Free Control:** Issue voice commands to open, read, or manage emails.
* **Text-to-Speech:** Hermys can read your emails out loud to you.

### 📊 Visual Insights
* **Interactive Graphs:** View email volume by day, sentiment trends over time, and top senders.
* **Dark/Light Mode:** Toggle between themes for visual comfort.

---

## 🛠 Tech Stack

* **Language:** Python 3.x
* **GUI Framework:** PyQt5
* **AI Model:** Google Gemini (GenAI API)
* **NLP Libraries:**
    * `NLTK` & `spaCy` (Processing)
    * `TextBlob` & `VADER` (Sentiment)
    * `Sumy` (Summarization)
* **Protocols:** `imaplib` (IMAP), `smtplib` (SMTP)

---

## 📸 Screenshots

### 1. Login Screen
*Secure login using Email, App Password, and Gemini API Key.*
<img width="665" height="446" alt="Image" src="https://github.com/user-attachments/assets/872ee386-5ab0-46a5-b7b5-1b70ef556f3a" />

### 2. Main Dashboard & NLP Analysis
*Inbox view with real-time Sentiment and Intent analysis side-by-side.*
![Image](https://github.com/user-attachments/assets/34b72dc7-9866-42cc-99ff-2760fa3463c6)

### 3. Analytics & Dark Mode
*Visual data representation of email habits.*
![Image](https://github.com/user-attachments/assets/cb29b828-864a-43fb-bb6f-f64810420a27)

---

## ⚙ Installation & Setup

### Prerequisites
1.  **Python 3.8+** installed.
2.  **Gmail Account** with 2-Step Verification enabled.
3.  **Google App Password** (required for secure login).
4.  **Google Gemini API Key** (for AI features).

### Steps
1.  **Clone the Repository**
    ```bash
    git clone [https://github.com/yourusername/hermys.git](https://github.com/yourusername/hermys.git)
    cd hermys
    ```

2.  **Install Dependencies**
    ```bash
    pip install -r requirements.txt
    ```
    *Note: You may need to download specific spaCy/NLTK models:*
    ```bash
    python -m spacy download en_core_web_sm
    ```

3.  **Run the Application**
    ```bash
    python email_sender1.py
    ```

---

## 🗣 Voice Commands
Hermys listens for specific keywords to execute actions. Examples include:

* **"Read emails from 1 to 3"** -> Reads the first three emails in the list.
* **"Forward this email to [Name]"** -> Opens the composer with the contact's email pre-filled.
* **"Generate reply to this email and send after 1 minute"** -> Uses AI to draft and schedule a response.
* **"Delete this email"** -> Moves the current email to trash.

---

## 🔮 Future Enhancements
* **Database Integration:** Adding SQLite to store user preferences and login history.
* **Mobile Support:** Developing a companion mobile application.
* **Advanced Threading:** improved handling of long email conversation threads.
* **Offline Mode:** Caching emails for offline access.

---

## 👨‍💻 Author

**Manas M Surve**
* **Project:** Mini Project (Semester I)
* **Institute:** Navinchandra Mehta Institute of Technology & Development (NMITD)
* **Year:** 2024-2025

---
*Disclaimer: This project uses live IMAP/SMTP connections. Ensure you keep your App Password and API Keys secure.*
