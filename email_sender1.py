import sys
import os
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QPushButton,
    QTextEdit, QTreeWidget, QTreeWidgetItem, QFrame, QSizePolicy, QSplitter, QMenu, 
    QCompleter, QButtonGroup, QStackedWidget, QDialog, QFormLayout, QDateEdit, QComboBox,
    QDialogButtonBox, QMessageBox, QFileDialog, QToolButton, QInputDialog, QTextBrowser
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QTimer, QDate, QUrl, QObject
from PyQt5.QtGui import QFont, QPixmap
from PyQt5.QtWebEngineWidgets import QWebEngineView
import imaplib
import smtplib
import email
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import parseaddr, make_msgid
from email.mime.image import MIMEImage
from email.mime.base import MIMEBase
from email import encoders
from PyQt5.QtMultimedia import QSound
from textblob import TextBlob
import spacy
from nltk import word_tokenize
from nltk.tokenize import sent_tokenize
from nltk.corpus import stopwords
from vaderSentiment.vaderSentiment import SentimentIntensityAnalyzer
import re
from email.utils import parseaddr
import json
import plotly.graph_objs as go
import plotly.express as px
import tempfile
import datetime
from email.utils import parsedate_to_datetime
import pandas as pd
from collections import Counter, defaultdict
import time
from sumy.parsers.plaintext import PlaintextParser
from sumy.nlp.tokenizers import Tokenizer as SumyTokenizer
from sumy.summarizers.lex_rank import LexRankSummarizer
from email.header import decode_header
import speech_recognition as sr
import pyttsx3
import re
import google.generativeai as genai
from threading import Timer
from word2number import w2n
from datetime import datetime, timedelta
import base64
from docx import Document
from pptx import Presentation
from PIL import Image
import subprocess
import platform
from transformers import BartTokenizer, BartForConditionalGeneration
import torch

class EmailDashboard(QWidget):
    def __init__(self, email_address, app_password, api_key):
        super().__init__()
        self.email_address = email_address
        self.app_password = app_password
        self.api_key = api_key
        self.dark_mode = False
        self.setWindowTitle("AI Email Dashboard")
        self.setGeometry(100, 100, 1400, 800)
        self.setStyleSheet(self.style_sheet())
        self.recent_recipients = set()
        self.imap_session = None
        self.is_forwarding = False
        self.refresh_timer = QTimer(self)
        self.refresh_timer.timeout.connect(self.load_emails)
        self.summarizer = SmartSummarizer()
        self.refresh_backoff = 60000  # Initial backoff (60 sec)
        self.refresh_timer.start(self.refresh_backoff)
        self.refresh_in_progress = False
        self.loaded_email_ids = set()
        self.notification_manager = NotificationManager(self)
        self.seen_email_ids = set()
        self.initial_load_done = False
        self.current_analysis = {}
        self.load_contact_dict()
        self.nlp = spacy.load("en_core_web_sm")
        self.intent_triggers = {"forward", "delete", "read", "open", "remind", "reply", "send", "cancel", "stop"}
        self.graph_views = []
        self.all_emails = []  # Stores full parsed email objects for graphing
        self.notification_sound = QSound("assets\notification.wav") 
        self.message_sound = QSound("assets\message.wav")
        self.recognizer = sr.Recognizer()
        self.microphone = sr.Microphone()
        self.tts_engine = pyttsx3.init()
        self.pending_reply = None
        self.pending_timer = None
        self.listening = False
        self.idle = False
        self.wake_word = "hermes" 
        self.index_word_map = {
            "first": 0, "second": 1, "third": 2, "fourth": 3,
            "fifth": 4, "sixth": 5, "seventh": 6, "eighth": 7,
            "ninth": 8, "tenth": 9, "eleventh": 10, "twelfth": 11,
            "thirteenth": 12, "fourteenth": 13, "fifteenth": 14,
            "sixteenth": 15, "seventeenth": 16, "eighteenth": 17,
            "nineteenth": 18, "twentieth": 19
        }
        self.folder_email_map = {}
        self.email_item_list = []
        self.user_action = None
        self.last_voice_command = "" 
        try:
            self.nlp = spacy.load("en_core_web_sm")
        except:
            import os
            os.system("python -m spacy download en_core_web_sm")
            self.nlp = spacy.load("en_core_web_sm")
        self.build_ui()

    def build_ui(self):

        sidebar_frame = QFrame()
        sidebar_frame.setObjectName("sidebar")
        sidebar_frame.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Expanding)

        sidebar_layout = QVBoxLayout(sidebar_frame)
        sidebar_layout.setSpacing(12)
        sidebar_layout.setContentsMargins(10, 10, 10, 10)

        inbox_label = QLabel("📥 Inbox")
        inbox_label.setFont(QFont("Segoe UI", 18, QFont.Bold))
        inbox_label.setProperty("class", "section-title")
        sidebar_layout.addWidget(inbox_label)

        self.search_bar = QLineEdit()
        self.search_bar.setPlaceholderText("Search emails...")
        sidebar_layout.addWidget(self.search_bar)
        self.search_bar.textChanged.connect(self.filter_email_tree)

        self.email_tree = QTreeWidget()
        self.email_tree.setHeaderLabels(["Emails"])
        self.email_tree.setIndentation(10)
        self.email_tree.setRootIsDecorated(True)
        self.email_tree.setExpandsOnDoubleClick(True)
        self.email_tree.itemClicked.connect(self.display_email)
        sidebar_layout.addWidget(self.email_tree)

        self.email_tree.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.email_tree.setMinimumHeight(400)
        sidebar_layout.addSpacing(10) 

        profile_layout = QHBoxLayout()
        email_label = QLabel(self.email_address)

        menu_button = QPushButton("☰")
        menu_button.setFixedSize(30, 30)
        menu_button.setStyleSheet("QPushButton { border: none; font-size: 18px; }")
        menu_button.clicked.connect(self.show_menu)

        profile_layout.addWidget(email_label)
        profile_layout.addStretch()
        profile_layout.addWidget(menu_button)

        sidebar_layout.addLayout(profile_layout)

        middle_splitter = QSplitter(Qt.Vertical)
        middle_splitter.setHandleWidth(5) 
        middle_splitter.addWidget(self.build_panel("📧 Email Viewer", is_viewer=True))
        middle_splitter.addWidget(self.build_panel("✍️ Email Composer", is_composer=True))
        middle_splitter.setSizes([300, 300])
        middle_splitter.setChildrenCollapsible(True)

        right_splitter = QSplitter(Qt.Vertical)
        right_splitter.addWidget(self.build_panel("🧠 NLP Analysis", is_analysis=True))
        right_splitter.addWidget(self.build_panel("🤖 Assistant & Insights", is_assistant=True))
        right_splitter.setSizes([300, 300])
        right_splitter.setChildrenCollapsible(True)

        sidebar_wrapper = QWidget()
        sidebar_wrapper_layout = QVBoxLayout(sidebar_wrapper)
        sidebar_wrapper_layout.setContentsMargins(10, 10, 10, 10)
        sidebar_wrapper_layout.addWidget(sidebar_frame)
        sidebar_wrapper.setMinimumWidth(300)

        middle_widget = QWidget()
        middle_layout = QVBoxLayout(middle_widget)
        middle_layout.setContentsMargins(10, 10, 10, 10)
        middle_layout.addWidget(middle_splitter)
        
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_layout.setContentsMargins(10, 10, 10, 10)
        right_layout.addWidget(right_splitter)

        middle_splitter.setStyleSheet("""
            QSplitter::handle {
                height: 10px;
            }
        """)
        middle_splitter.setHandleWidth(10)
        
        right_splitter.setStyleSheet("""
            QSplitter::handle {
                height: 10px;
            }
        """)
        right_splitter.setHandleWidth(10)

        splitter = QSplitter(Qt.Horizontal)
        splitter.addWidget(sidebar_wrapper)
        splitter.addWidget(middle_widget)
        splitter.addWidget(right_widget)
        splitter.setSizes([350, 700, 350])
        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 2)
        splitter.setStretchFactor(2, 1)


        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(5, 5, 5, 5)
        main_layout.addWidget(splitter)
        self.setLayout(main_layout)

        self.load_emails()

    def show_menu(self):
        menu = QMenu()
        theme_action = menu.addAction("Toggle Dark/Light Mode")
        theme_action.triggered.connect(self.toggle_theme)
        menu.exec_(self.mapToGlobal(self.sender().pos()))

    def build_button_row(self, labels):
        layout = QHBoxLayout()
        layout.setSpacing(8)
        for label in labels:
            btn = QPushButton(label)
            layout.addWidget(btn)
        layout.addStretch()
        return layout

    def build_panel(self, title, is_viewer=False, is_composer=False, is_analysis=False, is_assistant=False):
        frame = QFrame()
        frame.setStyleSheet("QFrame { border: 1px solid #ccc; border-radius: 10px; background-color: transparent; }")
        frame.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

        box = QVBoxLayout(frame)
        box.setSpacing(10)
        box.setContentsMargins(12, 12, 12, 12)

        title_label = QLabel(title)
        title_label.setObjectName("section-title")
        title_label.setFont(QFont("Segoe UI", 18, QFont.Bold))
        title_label.setStyleSheet("QLabel { border: none; background: none; }")
        box.addWidget(title_label)

        if is_viewer:
            self.email_viewer = QTextBrowser()
            self.email_viewer.setPlaceholderText("Click an email to view...")
            self.email_viewer.setOpenExternalLinks(True)
            box.addWidget(self.email_viewer)
            self.preview_area = QVBoxLayout()
            box.addLayout(self.preview_area)
            box.addLayout(self.build_button_row(["Forward", "Delete"]))

        elif is_composer:
            self.attached_files = []
            self.to_input = QLineEdit(); self.to_input.setPlaceholderText("To")
            self.cc_input = QLineEdit(); self.cc_input.setPlaceholderText("CC"); self.cc_input.setVisible(False)
            self.bcc_input = QLineEdit(); self.bcc_input.setPlaceholderText("BCC"); self.bcc_input.setVisible(False)
            self.subject_input = QLineEdit(); self.subject_input.setPlaceholderText("Subject")
            self.body_input = QTextEdit(); self.body_input.setPlaceholderText("Write your message...")
            # CC/BCC clickables
            cc_label = QLabel("<a href='#'>Cc</a>")
            cc_label.linkActivated.connect(lambda: self.cc_input.setVisible(True))
            bcc_label = QLabel("<a href='#'>Bcc</a>")
            bcc_label.linkActivated.connect(lambda: self.bcc_input.setVisible(True))
            cc_bcc_row = QHBoxLayout(); cc_bcc_row.addWidget(cc_label); cc_bcc_row.addWidget(bcc_label); cc_bcc_row.addStretch()
            # Icon tools (Attach, Photo, Link, Discard)
            send_row = QHBoxLayout()
            send_btn = QPushButton("Send")
            send_btn.clicked.connect(self.send_email)
            send_row.addWidget(send_btn)
            spacer = QWidget(); spacer.setFixedWidth(20)
            send_row.addWidget(spacer)
            attach_btn = QToolButton(); attach_btn.setText("📎"); attach_btn.setToolTip("Attach File")
            attach_btn.clicked.connect(self.attach_file)
            photo_btn = QToolButton(); photo_btn.setText("🖼️"); photo_btn.setToolTip("Insert Image")
            photo_btn.clicked.connect(self.insert_photo)
            link_btn = QToolButton(); link_btn.setText("🔗"); link_btn.setToolTip("Insert Link")
            link_btn.clicked.connect(self.insert_link)
            discard_btn = QToolButton(); discard_btn.setText("🗑️"); discard_btn.setToolTip("Discard Draft")
            discard_btn.clicked.connect(self.discard_draft)
            for btn in [attach_btn, photo_btn, link_btn]:
                send_row.addWidget(btn)
            send_row.addStretch()
            send_row.addWidget(discard_btn)
            # Layout assembly  
            box.addWidget(self.to_input)
            box.addLayout(cc_bcc_row)
            box.addWidget(self.cc_input)
            box.addWidget(self.bcc_input)
            box.addWidget(self.subject_input)
            box.addWidget(self.body_input)
            box.addLayout(send_row)

        elif is_analysis:
            self.analysis_display = QTextEdit("Sentiment | Keywords | Summary")
            self.analysis_display.setReadOnly(True)
            self.analysis_display.setFont(QFont("Segoe UI", 16))  # Font Size
            self.analysis_display.setAcceptRichText(True)
            box.addWidget(self.analysis_display)
            # Create toggle buttons
            self.summary_btn = QPushButton("📝 Summary")
            self.sentiment_btn = QPushButton("😃 Sentiment")
            self.keywords_btn = QPushButton("🔑 Keywords")

            for btn in [self.summary_btn, self.sentiment_btn, self.keywords_btn]:
                btn.setCheckable(True)
                btn.setFont(QFont("Segoe UI", 12))
                
            # Create button group
            self.toggle_group = QButtonGroup()
            self.toggle_group.setExclusive(True)
            self.toggle_group.addButton(self.summary_btn)
            self.toggle_group.addButton(self.sentiment_btn)
            self.toggle_group.addButton(self.keywords_btn)
            
            # Toggle button layout
            toggle_layout = QHBoxLayout()
            toggle_layout.addWidget(self.summary_btn)
            toggle_layout.addWidget(self.sentiment_btn)
            toggle_layout.addWidget(self.keywords_btn)
            box.addLayout(toggle_layout)
            
            # Connect buttons to update display
            self.summary_btn.clicked.connect(lambda: self.update_analysis_display("summary"))
            self.sentiment_btn.clicked.connect(lambda: self.update_analysis_display("sentiment"))
            self.keywords_btn.clicked.connect(lambda: self.update_analysis_display("keywords"))

        elif is_assistant:
            self.assist_display = QTextEdit("I'm Hermes, your assistant.")
            self.assist_display.setReadOnly(True)
            self.assist_display.setFont(QFont("Segoe UI", 16))  # Font Size
            self.assist_display.setAcceptRichText(True)
            box.addWidget(self.assist_display)
            # Chatbot input (only shown in chatbot mode)
            self.chat_input = QLineEdit()
            self.chat_input.setPlaceholderText("Type or speak a command...")
            self.chat_input.setFont(QFont("Segoe UI", 12))
            self.mic_btn = QPushButton("🎙️")
            self.mic_btn.setToolTip("Start voice input")
            self.mic_btn.clicked.connect(self.start_voice_input)  # Hook your STT logic
            self.mic_btn.setVisible(False)
            self.chat_input.setVisible(False)  # hidden by default
            self.chat_send_btn = QPushButton("↑")
            self.chat_send_btn.setVisible(False)
            self.chat_send_btn.clicked.connect(self.chatbot_chat)
            chat_row = QHBoxLayout()
            chat_row.addWidget(self.chat_input)
            chat_row.addWidget(self.mic_btn)
            chat_row.addWidget(self.chat_send_btn)
            box.addLayout(chat_row)

            #Time Pattern
            self.graph_views = []
            self.graph_nav_container = QWidget()
            self.graph_nav_layout = QHBoxLayout(self.graph_nav_container)
            self.prev_graph_btn = QPushButton("⬅")
            self.next_graph_btn = QPushButton("➡")
            self.graph_title = QLabel("📊 Email Volume by Day")
            self.graph_title.setStyleSheet("font-weight: bold; font-size: 16px;")
            self.graph_nav_layout.addWidget(self.prev_graph_btn)
            self.graph_nav_layout.addWidget(self.graph_title)
            self.graph_nav_layout.addWidget(self.next_graph_btn)
            self.graph_nav_layout.setAlignment(Qt.AlignCenter)
            self.filter_btn = QPushButton("⏱️ Filter")
            self.graph_nav_layout.addWidget(self.filter_btn)
            # Later show a small dialog when clicked
            self.filter_btn.clicked.connect(self.open_time_filter_dialog)
            self.export_btn = QPushButton("📤 Export")
            self.graph_nav_layout.addWidget(self.export_btn)
            self.export_btn.clicked.connect(self.export_current_graph)
            self.graph_nav_container.setVisible(False)
            self.graph_stack = QStackedWidget()
            box.addWidget(self.graph_stack)
            self.graph_stack.setVisible(False)
            self.graph_index = 0
            self.graph_names = [
                "📊 Email Volume by Day",
                "📈 Top Senders",
                "📅 Sentiment Over Time",
                "🧠 Keywords Over Time"
            ]
            for name in self.graph_names:
                placeholder = QWidget()
                layout = QVBoxLayout(placeholder)
                web_view = QWebEngineView()
                layout.addWidget(web_view)
                self.graph_stack.addWidget(placeholder)
                self.graph_views.append(web_view)
            self.prev_graph_btn.clicked.connect(self.show_prev_graph)
            self.next_graph_btn.clicked.connect(self.show_next_graph)
            box.addWidget(self.graph_nav_container)
            # Time Pattern Filter Panel (Docked)
            self.filter_panel = QWidget()
            self.filter_layout = QHBoxLayout(self.filter_panel)
            self.filter_layout.setContentsMargins(0, 0, 0, 0)
            # Date Range Filter
            self.start_date_edit = QDateEdit()
            self.start_date_edit.setCalendarPopup(True)
            self.start_date_edit.setDisplayFormat("yyyy-MM-dd")
            self.start_date_edit.setDate(QDate.currentDate().addMonths(-1))  # default: 1 month back
            self.end_date_edit = QDateEdit()
            self.end_date_edit.setCalendarPopup(True)
            self.end_date_edit.setDisplayFormat("yyyy-MM-dd")
            self.end_date_edit.setDate(QDate.currentDate())  # default: today
            self.sender_filter = QLineEdit()
            self.sender_filter.setPlaceholderText("Filter by sender...")
            self.apply_filter_btn = QPushButton("✅ Apply")
            self.apply_filter_btn.clicked.connect(self.apply_filters_and_refresh_graphs)
            # Add to layout
            self.filter_layout.addWidget(QLabel("From:"))
            self.filter_layout.addWidget(self.start_date_edit)
            self.filter_layout.addWidget(QLabel("To:"))
            self.filter_layout.addWidget(self.end_date_edit)
            self.filter_layout.addWidget(self.sender_filter)
            self.filter_layout.addWidget(self.apply_filter_btn)
            self.filter_panel.setVisible(False)
            box.addWidget(self.filter_panel)
            # Create toggle button
            self.assistant_btn = QPushButton("🤖 Assistant")
            self.pattern_btn = QPushButton("📊 Time Pattern")

            for btn in [self.assistant_btn, self.pattern_btn]:
                btn.setCheckable(True)
                btn.setFont(QFont("Segoe UI", 12))
                
            # Create button group
            self.toggle_group = QButtonGroup()
            self.toggle_group.setExclusive(True)
            self.toggle_group.addButton(self.assistant_btn)
            self.toggle_group.addButton(self.pattern_btn)
            
            # Toggle button layout
            toggle_layout = QHBoxLayout()
            toggle_layout.addWidget(self.assistant_btn)
            toggle_layout.addWidget(self.pattern_btn)
            box.addLayout(toggle_layout)
            
            # Connect buttons to update display
            self.assistant_btn.clicked.connect(lambda: self.update_assist_display("assistant"))
            self.pattern_btn.clicked.connect(lambda: self.update_assist_display("pattern"))

        return frame
    



# Start of Email System [Sidebar, Email Viewer and Email Composer]
    def load_emails(self):
        if self.refresh_in_progress:
            return
        self.refresh_in_progress = True
        self.refresh_timer.stop()
        self.loader_thread = EmailLoaderThread(self.email_address, self.app_password)
        self.loader_thread.result.connect(self.handle_loader_result)
        self.loader_thread.error.connect(self.handle_loader_error)
        self.loader_thread.start()

    def handle_loader_result(self, results):
        self.populate_email_tree(results)
        self.refresh_in_progress = False
        self.refresh_timer.start(self.refresh_backoff)  # restart the timer

    def handle_loader_error(self, error_msg):
        self.email_viewer.setText(f"❌ {error_msg}")
        self.refresh_in_progress = False
        if "rate" in error_msg.lower() or "limit" in error_msg.lower():
            self.refresh_backoff = min(self.refresh_backoff * 2, 10 * 60000)  # cap at 10 mins
        else:
            self.refresh_backoff = 60000
        self.refresh_timer.start(self.refresh_backoff)

    def display_email(self, item):
        try:
            data = item.data(0, Qt.UserRole)
            if not data:
                return  # Ignore clicks on headers
            if not self.is_forwarding:
                self.to_input.clear()
                self.subject_input.clear()
                self.body_input.clear()
            folder, mail_id = data
            imap = self.get_imap_session()
            try:
                imap.noop()  
                imap.select(folder)
                if not mail_id:
                    raise ValueError("Invalid mail ID")
                imap.noop()  # Optional health check
                res, msg_data = imap.fetch(mail_id, "(RFC822)")
            except imaplib.IMAP4.abort:
                self.email_viewer.setText("❌ Connection aborted unexpectedly during fetch.")
                return
            except imaplib.IMAP4.error as e:
                self.email_viewer.setText(f"❌ IMAP error: {e}")
                return
            except Exception as e:
                self.email_viewer.setText(f"❌ General error during email display: {e}")
                return
            
            attachments = [] 
            self.current_attachments = attachments

            for response_part in msg_data:
                if isinstance(response_part, tuple):
                    msg = email.message_from_bytes(response_part[1])
                    subject = self.decode_mime_header(msg["subject"])
                    sender = self.decode_mime_header(msg["from"])
                    to = self.decode_mime_header(msg.get("to", ""))
                    body = ""
                    html_body = ""
                    inline_images = {}

                    if msg.is_multipart():
                        for part in msg.walk():
                            content_type = part.get_content_type()
                            content_disposition = part.get("Content-Disposition", "").lower()
                            if content_type == "text/plain" and "attachment" not in content_disposition:
                                body = part.get_payload(decode=True).decode(errors="ignore")
                            elif content_type == "text/html" and "attachment" not in content_disposition:
                                html_body = part.get_payload(decode=True).decode(errors="ignore")
                            elif "attachment" in content_disposition or part.get_filename():
                                filename = part.get_filename()
                                if filename:
                                    attachments.append((filename, part))
                            elif "image/" in content_type:  # Inline images
                                content_id = part.get("Content-ID", "").strip("<>")
                                if content_id:
                                    inline_images[content_id] = part
                    else:
                        body = msg.get_payload(decode=True).decode(errors="ignore")

                    if html_body:
                        for content_id, part in inline_images.items():
                            img_data = part.get_payload(decode=True)
                            img_base64 = base64.b64encode(img_data).decode("utf-8")
                            html_body = html_body.replace(f'cid:{content_id}', f'data:image/png;base64,{img_base64}')
                    
                    display_html = html_body or f"<pre>{body}</pre>"

                    full_html = (
                        f"<p><b>From:</b> {sender}<br>"
                        f"<b>To:</b> {to}<br>"
                        f"<b>Subject:</b> {subject}</p><hr>{display_html}"
                    )
                    self.email_viewer.setHtml(full_html)
                    self.show_attachment_thumbnails(attachments)
                    # Autofill "To" field
                    if not self.is_forwarding:
                        name, address = parseaddr(to if sender == self.email_address else sender)
                        self.to_input.setText(address)

                    # 🧠 Run NLP analysis on email body
                    analysis = self.analyze_email(body)
                    self.current_analysis = analysis
                    self.summary_btn.setChecked(True)
                    self.update_analysis_display("summary")  # Automatically trigger summary view

                    email_obj = { 
                         "from": sender,
                         "to": to,
                         "subject": subject,
                         "body": body,
                         "date": parsedate_to_datetime(msg["Date"]) if msg["Date"] else datetime.now(),
                         "folder": folder,
                         "mail_id": mail_id,
                         "sentiment": analysis.get("sentiment_score", 0),  # From your `analyze_email()`
                         "resolved": analysis.get("resolved", False),       # You can decide logic
                         "reply_time": analysis.get("reply_time", None),    # Optional to comput
                    }
                    email_key = f"{folder}:{mail_id}"
                    if email_key not in self.loaded_email_ids:
                        self.all_emails.append(email_obj)
                        self.loaded_email_ids.add(email_key)

            imap.logout()
            self.is_forwarding = False
        except Exception as e:
            self.email_viewer.setText(f"Error displaying email: {e}")

    def send_email(self):
        to = self.to_input.text()
        cc = self.cc_input.text()
        bcc = self.bcc_input.text()
        subject = self.subject_input.text()
        html_body = self.body_input.toHtml()
        try:
            msg = MIMEMultipart("mixed")
            msg["From"] = self.email_address
            msg["To"] = to
            if cc: msg["Cc"] = cc
            if bcc: msg["Bcc"] = bcc
            msg["Subject"] = subject
            # Multipart/alternative and related
            alternative = MIMEMultipart("alternative")
            related = MIMEMultipart("related")
            # Plain fallback text
            plain_text = "You're viewing the plain version of this email. Switch to HTML view for full content."
            alternative.attach(MIMEText(plain_text, "plain"))
            # Handle inline images from QTextEdit (src="file:///...")
            image_pattern = r'<img[^>]+src=["\']file:///([^"\']+)["\']'
            matches = re.findall(image_pattern, html_body)
            for abs_path in matches:
                abs_path = abs_path.replace("\\", "/")
                if not os.path.exists(abs_path):
                    continue
                cid = make_msgid()[1:-1]
                html_body = html_body.replace(f"file:///{abs_path}", f"cid:{cid}")
                with open(abs_path, "rb") as f:
                    img_data = f.read()
                img = MIMEImage(img_data)
                img.add_header("Content-ID", f"<{cid}>")
                img.add_header("Content-Disposition", "inline", filename=os.path.basename(abs_path))
                related.attach(img)
            # Attach final HTML (after cid replacement)
            related.attach(MIMEText(html_body, "html"))
            alternative.attach(related)
            msg.attach(alternative)
            # Attach non-image files normally
            for file_path in getattr(self, "attached_files", []):
                if not file_path.lower().endswith((".png", ".jpg", ".jpeg", ".bmp")):
                    if os.path.exists(file_path):
                        with open(file_path, "rb") as f:
                            content = f.read()
                        attachment = MIMEBase("application", "octet-stream")
                        attachment.set_payload(content)
                        encoders.encode_base64(attachment)
                        filename = os.path.basename(file_path)
                        attachment.add_header("Content-Disposition", f"attachment; filename={filename}")
                        msg.attach(attachment)
            # Send via SMTP
            with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
                smtp.login(self.email_address, self.app_password)
                smtp.send_message(msg)
            # Cleanup UI
            self.to_input.clear()
            self.cc_input.clear(); self.cc_input.setVisible(False)
            self.bcc_input.clear(); self.bcc_input.setVisible(False)
            self.subject_input.clear()
            self.body_input.clear()
            self.attached_files = []
            self.email_viewer.setText("✅ Email sent successfully.")
            self.recent_recipients.add(to)
            completer = QCompleter(sorted(self.recent_recipients))
            completer.setCaseSensitivity(Qt.CaseInsensitive)
            self.to_input.setCompleter(completer)
        except smtplib.SMTPException as smtp_error:
            self.email_viewer.setText(f"❌ Error sending email (SMTP): {smtp_error}")
        except Exception as e:
            self.email_viewer.setText(f"❌ Error sending email: {e}")
    
    def populate_email_tree(self, results):
        self.email_tree.clear()
        self.email_item_list = []  # 🔥 NEW: flat list for number-based access
        if not hasattr(self, 'contact_dict'):
            self.contact_dict = {}
        for category, mails in results:
            parent = QTreeWidgetItem(self.email_tree, [category])
            parent.setExpanded(True)
            parent.setFlags(parent.flags() & ~Qt.ItemIsSelectable)
            self.folder_email_map[category.lower()] = []
            for folder, mail_id, subject, sender, to in mails:
                subject = self.decode_mime_header(subject)
                sender = self.decode_mime_header(sender)
                sender_name, sender_addr = parseaddr(sender)
                sender = f"{sender_name} <{sender_addr}>" if sender_name else sender_addr
                to_name, to_addr = parseaddr(to)
                to = f"{to_name} <{to_addr}>" if to_name else to_addr
                to = self.decode_mime_header(to)
                display = f"{subject} - {sender}" if category != "Sent" else f"{subject} - to: {to}"
                child = QTreeWidgetItem([display])
                child.setData(0, Qt.UserRole, (folder, mail_id))
                parent.addChild(child)
                self.email_item_list.append(child)
                self.folder_email_map[category.lower()].append(child)  # 💥 Add to index map
                # Avoid adding your own email
                if category != "Sent":
                    self.store_contact_from_line(sender)  # Extract sender from inbox emails
                else:
                    for recipient in to.split(","):
                        if self.email_address.lower() not in recipient.lower():
                            self.store_contact_from_line(recipient) # Extract recipients from sent emails
        if self.initial_load_done: 
            for category, mails in results:
                for folder, mail_id, subject, sender, to in mails:
                    self.notification_manager.maybe_notify(sender, folder, mail_id, category)
                    return
        self.initial_load_done = True

    def forward_email(self, to_name=None):
        self.is_forwarding = True
        viewer_text = self.email_viewer.toPlainText()
        if not viewer_text:
            return
        # Extract subject/body as before
        lines = viewer_text.split("\n")
        subject_line = next((line for line in lines if line.startswith("Subject:")), "")
        subject = subject_line.replace("Subject:", "").strip()
        body = "\n".join(lines[3:])
        self.subject_input.setText("FWD: " + subject)
        self.body_input.setPlainText(body)
        if to_name:
            email = self.get_email_by_name(to_name)
            if email:
                self.to_input.setText(email)
                self.assist_display.setText("📝 Forwarding via assistant. Sending...")
                self.send_email()
            else:
                self.email_viewer.setText(f"⚠️ Could not find contact '{to_name}'")
                return
        else:
            self.to_input.clear()
        self.assist_display.setText(f"📝 Forwarding the current Email")

    def delete_email(self):
        item = self.email_tree.currentItem()
        if not item:   
            index = getattr(self, "current_selected_index", None)
            if index is not None and index < self.email_tree.topLevelItemCount():
                item = self.email_tree.topLevelItem(index)
        if not item:
            self.email_viewer.setText("⚠️ No email selected to delete.")
            return
        data = item.data(0, Qt.UserRole)
        if not data:
            return
        folder, mail_id = data
        try:
            imap = imaplib.IMAP4_SSL("imap.gmail.com")
            imap.login(self.email_address, self.app_password) 
            imap.select(folder)
            imap.store(mail_id, '+FLAGS', '\\Deleted')
            imap.expunge()
            imap.logout()
            self.email_viewer.setText("Email deleted successfully.")
            parent = item.parent()
            if parent:
                parent.removeChild(item)
            else:
                index = self.email_tree.indexOfTopLevelItem(item)
                self.email_tree.takeTopLevelItem(index)
        except Exception as e:
            self.email_viewer.setText(f"❌ Error deleting email: {e}")

    def build_button_row(self, labels):
        layout = QHBoxLayout()
        layout.setSpacing(8)
        for label in labels:
            btn = QPushButton(label)
            if label == "Forward":
                btn.clicked.connect(self.forward_email)
            elif label == "Delete":
                btn.clicked.connect(self.delete_email)  
            layout.addWidget(btn)
        layout.addStretch()
        return layout
    
    def filter_email_tree(self, text):
        text = text.lower().strip()
        for i in range(self.email_tree.topLevelItemCount()):
            parent = self.email_tree.topLevelItem(i)
            parent_visible = False
            for j in range(parent.childCount()):
                child = parent.child(j)
                label = child.text(0).lower()
                match = text in label
                child.setHidden(not match)
                if match:
                    parent_visible = True
            parent.setHidden(not parent_visible)

    def attach_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "Attach File")
        if path:
            self.attached_files.append(path)
            self.email_viewer.setText(f"📎 Attached: {path.split('/')[-1]}")
    
    def insert_photo(self):
        path, _ = QFileDialog.getOpenFileName(self, "Insert Image", filter="Images (*.png *.jpg *.jpeg *.bmp)")
        if path:
            path = path.replace("\\", "/")  # Ensure forward slashes (important for regex and HTML)
            if not path.startswith("file:///"):
                path = "file:///" + path  # Make it a proper file URI
            self.body_input.insertHtml(f'<img src="{path}" width="400">')

    def insert_link(self):
        link, ok = QInputDialog.getText(self, "Insert Link", "Enter URL:")
        if ok and link:
            self.body_input.insertHtml(f'<a href="{link}">{link}</a>')
            
    def discard_draft(self):
        self.to_input.clear()
        self.subject_input.clear()
        self.body_input.clear()
        self.cc_input.clear(); self.cc_input.setVisible(False)
        self.bcc_input.clear(); self.bcc_input.setVisible(False)
        self.attached_files = []
        self.email_viewer.setText("🗑️ Draft discarded.")

    def decode_mime_header(self, value):
        if not value:
            return ""
        parts = decode_header(value)
        decoded = ""
        for part, encoding in parts:
            if isinstance(part, bytes):
                decoded += part.decode(encoding or "utf-8", errors="ignore")
            else:
                decoded += part
        return decoded

    def get_imap_session(self):
        if self.imap_session is None:
            self.imap_session = imaplib.IMAP4_SSL("imap.gmail.com")
            self.imap_session.login(self.email_address, self.app_password)
        try:
            self.imap_session.noop()
        except:
            # Reset on failure
            self.imap_session = imaplib.IMAP4_SSL("imap.gmail.com")
            self.imap_session.login(self.email_address, self.app_password)
        return self.imap_session

    def open_email_from_notification(self, folder, mail_id):
        # Find and select the email in the tree
        root = self.email_tree.invisibleRootItem()
        for i in range(root.childCount()):
            parent = root.child(i)
            for j in range(parent.childCount()):
                child = parent.child(j)
                data = child.data(0, Qt.UserRole)
                if data == (folder, mail_id):
                    self.email_tree.setCurrentItem(child)
                    self.display_email(child)
                    self.notification_manager.hide_notification()
                    return



    def open_with_default_app(self, file_path):
        try:
            system_platform = platform.system()
            if system_platform == "Windows":
                os.startfile(file_path)
            elif system_platform == "Darwin":  # macOS
                subprocess.Popen(["open", file_path])
            else:  
                subprocess.Popen(["xdg-open", file_path])
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Could not open file: {e}")
            
    def save_attachment(self, filename, part):
        save_path, _ = QFileDialog.getSaveFileName(
            self,
            f"Save Attachment - {filename}",
            filename,
            "All Files (*)"
        )
        if save_path:
            try:
                with open(save_path, "wb") as file:
                    file.write(part.get_payload(decode=True))
                QMessageBox.information(self, "Success", f"Attachment saved to: {save_path}")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to save attachment: {e}")

    def show_attachment_thumbnails(self, attachments):
        layout = self.preview_area.layout()
        while layout.count():
            child = layout.takeAt(0)
            if child.widget():
                child.widget().deleteLater()
        if not attachments:
            return
        for filename, part in attachments:
            hbox = QHBoxLayout()
            icon_label = QLabel()
            ext = filename.split('.')[-1].lower()
            if ext in ["jpg", "jpeg", "png", "gif"]:
                img_data = part.get_payload(decode=True)
                pixmap = QPixmap()
                pixmap.loadFromData(img_data)
                icon_label.setPixmap(pixmap.scaled(60, 60, Qt.KeepAspectRatio))
            else:
                icon_label.setText("📎")  
            link = QLabel(f'<a href="attachment://{filename}">{filename}</a>')
            link.setOpenExternalLinks(False)
            link.linkActivated.connect(lambda url, p=part: self.open_attachment_popup(p))
            hbox.addWidget(icon_label)
            hbox.addWidget(link)
            hbox.addStretch()
            container = QWidget()
            container.setLayout(hbox)
            layout.addWidget(container)

    def on_attachment_click(self, url, part):
        filename = url.split("attachment://")[-1]
        ext = filename.split('.')[-1].lower()
        temp_dir = tempfile.gettempdir()
        temp_path = os.path.join(temp_dir, filename)
        with open(temp_path, "wb") as f:
            f.write(part.get_payload(decode=True))
        if ext in ['pdf', 'png', 'jpg', 'jpeg', 'gif']:
            self.open_attachment_popup(part)
        else:
            self.open_with_default_app(temp_path)

    def open_attachment_popup(self, part):
        img_data = part.get_payload(decode=True)
        pixmap = QPixmap()
        pixmap.loadFromData(img_data)
        if not pixmap.isNull():  # It's an image
            dialog = AttachmentPreviewDialog(part, self)
            dialog.exec_()
        else:
            filename = part.get_filename() or "attachment"
            self.ext = filename.split('.')[-1].lower()
            temp_path = os.path.join(tempfile.gettempdir(), filename)
            with open(temp_path, "wb") as f:
                f.write(img_data)
            self.open_with_default_app(temp_path)


# Start of NLP Analysis Panel
    def analyze_email(self, text):
        blob = TextBlob(text)
        doc = self.nlp(text)
        sentences = sent_tokenize(text)

        # ✍️ SMART SUMMARY (Named entities + verbs + people/orgs + context)
        summary = self.generate_summary(text, doc, sentences)
        
        # 😊 EMOJI-SENTIMENT + TONE
        tone, emoji, polarity, subjectivity, compound = self.analyze_sentiment(text)

        # 📌 INTENT CLASSIFICATION
        intent = self.classify_intent(text)

        # 🔑 KEYWORD EXTRACTION (color labels later)
        keywords = list({ent.text for ent in doc.ents if ent.label_ in (
        "ORG", "PERSON", "DATE", "TIME", "GPE", "MONEY", "CARDINAL", "EVENT", "PRODUCT")})
    
        # 📦 FORMAT OUTPUT
        return {
            "summary": f"✍️ {summary}",
            "sentiment": {
                "emoji": emoji,
                "tone": tone,
                "polarity": polarity,
                "subjectivity": subjectivity,
                "compound": compound
            },
            "intent": f"📌 Intent: {intent}",
            "keywords": keywords or ["No significant keywords found."]
        } 
        
    def generate_summary(self, text, doc, sentences):
        return self.summarizer.generate_summary(text, doc, sentences)
    
    def classify_intent(self, text):
        lowered = text.lower()
        # --- REGEX FIRST: Bulletproof intent detection --
        intent_map = {
            "Warning / Escalation": [
                r"final (?:notice|warning)", r"\bescalat(?:e|ion)\b", r"\bunacceptable\b",
                r"this (?:cannot|won't) continue", r"serious (?:issue|problem)", r"\bmust (?:resolve|fix|respond)\b",
                r"\burgent\b", r"\bwe expect\b", r"no longer tolerat", r"\btake this seriously\b"
            ],
            "Complaint / Negative Feedback": [
                r"\bdissatisf(?:ied|action)\b", r"\bdisappointed\b", r"\bfrustrated\b",
                r"bad (?:experience|product|service)", r"\bpoor\b", r"not (?:happy|satisfied)",
                r"\bissue with\b", r"problem with", r"unresolved", r"\bnot what I expected\b"
            ],
            "Apology / Customer Support Follow-up": [
                r"\b(?:i|we)? ?apolog(?:ize|y)\b", r"\bsorry\b", r"inconvenience", r"delay(?:ed)?",
                r"\bour mistake\b", r"\bwe regret\b", r"technical difficulties", r"please bear with us"
            ],
            "Inquiry / Request": [
                r"can you", r"could you", r"please (?:help|advise|explain|let me know)",
                r"i would like to (?:know|inquire|ask)", r"\bquestion about\b", r"would you mind",
                r"\bwhat is the\b", r"\bhow do i\b", r"i need help", r"i have a question"
            ],
            "Delivery Status / Confirmation": [
                r"\bshipped\b", r"\bdelivered\b", r"\bhas arrived\b", r"received (?:my|the) order",
                r"tracking number", r"order (?:status|update)", r"\bETA\b", r"dispatch(?:ed)?"
            ],
            "Return / Refund": [
                r"\breturn\b", r"\brefund\b", r"send (?:it )?back", r"not satisfied", r"replacement",
                r"under warranty", r"request a refund", r"\bexchange\b", r"\bproduct defect\b", r"\bdefective\b"
            ],
            "Appreciation / Feedback": [
                r"\bthank you\b", r"\bthanks\b", r"\bi appreciate\b", r"\bgrateful\b", r"\bpleased\b",
                r"great (?:job|support|service)", r"\bexcellent\b", r"\bawesome\b", r"well done"
            ],
            "Celebration / Positive News": [
                r"\bgot the job\b", r"\bpromotion\b", r"\bwe (?:won|secured)\b", r"\bexcited\b",
                r"\bgreat news\b", r"\bcelebrate\b", r"\bso happy\b", r"\bthrilled\b", r"achieved"
            ]
        }

        # Try regex patterns first
        for intent, patterns in intent_map.items():
            for pattern in patterns:
                if re.search(pattern, lowered):
                    return intent
            
        # --- NLTK FALLBACK: Fuzzy intent detection ---
        tokens = word_tokenize(lowered)
        token_set = set(tokens)
        stop_words = set(stopwords.words("english"))
        if {"escalate", "unacceptable", "violation", "terminate", "final", "warning"} & token_set:
            return "Warning / Escalation"
        if {"dissatisfied", "disappointed", "frustrated", "terrible", "angry", "bad", "issue", "problem", "not", "happy"} & token_set:
            return "Complaint / Negative Feedback"
        if {"apologize", "apology", "sorry", "inconvenience", "regret", "delay", "mistake"} & token_set:
            return "Apology / Customer Support Follow-up"
        if {"can", "could", "would", "please", "help", "advise", "how", "may", "let", "know"} & token_set:
            return "Inquiry / Request"
        if {"delivered", "shipped", "received", "arrival", "tracking", "update", "status"} & token_set:
            return "Delivery Status / Confirmation"
        if {"refund", "return", "replacement", "warranty", "not", "satisfied", "send", "back"} & token_set:
            return "Return / Refund"
        if {"thank", "thanks", "appreciate", "grateful", "pleased", "kudos", "great", "job"} & token_set:
            return "Appreciation / Feedback"
        if {"promotion", "won", "congrats", "happy", "excited", "great", "news", "celebrate"} & token_set:
            return "Celebration / Positive News"
        return "General / Uncategorized"
    
    def detect_tone(self, text):
        casual_markers = ["i'm", "you're", "cheers", "thanks", "btw", "hey", "hi"]
        formal_markers = ["sincerely", "regards", "apologize", "please", "unfortunately"]
        lower = text.lower()
        casual = sum(1 for word in casual_markers if word in lower)
        formal = sum(1 for word in formal_markers if word in lower)
        if formal > casual:
            return "Professional"
        if casual > formal:
            return "Casual"
        return "Neutral"
    
    def format_sentiment_for_ui(self, sentiment_data):
        if not sentiment_data:
            return "<i>No sentiment detected.</i>"
        
        # fallback map if no emoji provided
        emoji_map = {
            "excited": "🥳",
            "happy": "😊",
            "positive": "😄",
            "neutral": "😐",
            "negative": "😞",
            "angry": "😡",
            "sad": "😢",
            "confused": "🤔",
            "apologetic": "🙏",
            "complaint": "😞",
            "grateful": "🤝"
        }
        
        tone = sentiment_data.get("tone", "Neutral")
        emoji = sentiment_data.get("emoji", emoji_map.get(tone.lower(), "😐"))
        polarity = sentiment_data.get("polarity", "N/A")
        subjectivity = sentiment_data.get("subjectivity", "N/A")
        vader = sentiment_data.get("vader", "N/A")
        
        return f"""
        <div style="text-align:center; font-size:64px;">{emoji}</div>
        <div style="text-align:center; font-size:18px; font-weight:bold;">Tone: {tone}</div>
        <div style="text-align:center; font-size:14px;">Polarity: {polarity} | Subjectivity: {subjectivity} | VADER: {vader}</div>
        """

    def analyze_sentiment(self, text):
        blob = TextBlob(text)
        vader_analyzer = SentimentIntensityAnalyzer()
        vader_scores = vader_analyzer.polarity_scores(text)
        tb_polarity = blob.sentiment.polarity
        tb_subjectivity = blob.sentiment.subjectivity
        vader_compound = vader_scores['compound']
        lowered = text.lower()
        tone_rules = {
            "Apologetic": [
                r"\b(we apologize|i apologize|sorry for|regret to inform|our apologies)\b"
            ],
            "Complaint": [
                r"\b(disappointed|unacceptable|concerning|untenable|dissatisfaction|neglected|ignored|impacting our team|lack of communication|no response|frustrated |frustration |undermines? |not acceptable |not satisfied| not okay| missed deadlines?| lack of clarity| lack of response| lack of respect)\b"
            ],
            "Grateful": [
                r"\b(thank you|thanks a lot|truly appreciate|grateful for|many thanks|really appreciate|sincerely appreciate)\b"
            ],
            "Angry": [
                r"\b(fed up|angry|furious|final warning|completely unacceptable|tired of this|we demand| out of patience| enough is enough| absolutely not)\b"
            ],
            "Confused": [
                r"\b(confused|unclear|mixed messages|what’s going on|not sure|uncertain|inconsistent)\b"   
            ],
            "Excited": [
                r"\b(so excited|thrilled|can’t wait|awesome|over the moon|beyond happy|fantastic news)\b"
            ],
            "Happy": [
                r"\b(glad to|pleased to|happy to report|it’s working great|satisfying results|positive outcome|great to see)\b"
            ],
            "Sad": [
                r"\b(disheartened|saddened|let down|upset by|we regret|feeling down|not what we hoped)\b"
            ]
        }
        sarcasm_patterns = [
            r"\b(yeah right|as if|sure that’s working|oh great|what a surprise|just perfect|awesome\.\.\. not|bold reinterpretation|a masterclass in|rare privilege|mystery of it all|we remain inspired|return to this plane of existence|meditative pace|stillness)\b",
            r"(!{2,}|\?{2,})",  # excessive punctuation
            r"\b(nice job|great work)\b.*\b(not|really)\b",  # sarcastic praise
            r"\b(so inspiring|truly remarkable|we’re in awe)\b"
        ]
        for pattern in sarcasm_patterns:
            if re.search(pattern, lowered):
                return "Sarcastic", "🙃", tb_polarity, tb_subjectivity, vader_compound
        for tone, patterns in tone_rules.items():
            for pattern in patterns:
                if re.search(pattern, lowered):
                    emoji = {
                        "Apologetic": "🙏", "Complaint": "😞", "Grateful": "🤝", "Angry": "😡",
                        "Confused": "🤔", "Excited": "🥳", "Happy": "😊", "Sad": "😢"
                    }.get(tone, "😐")
                    return tone, emoji, tb_polarity, tb_subjectivity, vader_compound
        final_score = (tb_polarity + vader_compound) / 2
        if final_score >= 0.7:
            return "Excited", "🥳", tb_polarity, tb_subjectivity, vader_compound
        elif final_score >= 0.4:
            return "Happy", "😊", tb_polarity, tb_subjectivity, vader_compound
        elif final_score >= 0.2:
            return "Positive", "😄", tb_polarity, tb_subjectivity, vader_compound
        elif -0.1 <= final_score <= 0.1:
            if tb_subjectivity > 0.5:
                return "Confused", "🤔", tb_polarity, tb_subjectivity, vader_compound
            return "Neutral", "😐", tb_polarity, tb_subjectivity, vader_compound
        elif final_score <= -0.6:
            return "Angry", "😡", tb_polarity, tb_subjectivity, vader_compound
        elif final_score <= -0.3:
            return "Sad", "😢", tb_polarity, tb_subjectivity, vader_compound
        elif final_score <= -0.15:
            return "Negative", "😞", tb_polarity, tb_subjectivity, vader_compound
        else:
            return "Neutral", "😐", tb_polarity, tb_subjectivity, vader_compound

    def format_keywords_for_ui(self, keywords):
        if not keywords:
            return "<i>No keywords detected.</i>"
        
        colors = ["#F48FB1", "#90CAF9", "#A5D6A7", "#FFF176", "#CE93D8"]
        labels = []
        for i, kw in enumerate(keywords):
            color = colors[i % len(colors)]
            labels.append(
                f"<div style='margin:6px 0;'>"
                f"<span style='background-color:{color}; border-radius:8px; padding:6px 12px; font-size:18px; display:inline-block;'>"
                f"{kw}</span></div>" 
            )
        return "<div><b style='font-size:20px;'>Keywords:</b><br>" + "".join(labels) + "</div>"

    def update_analysis_display(self, mode):
        if not self.current_analysis:
            self.analysis_display.setHtml("<i>No analysis available.</i>")
            return
        
        intent = self.current_analysis.get("intent", "No intent detected.")
        intent_html = f"<div style='font-size:18px; font-weight:regular;'> {intent}</div><hr>"
        
        if mode == "summary":
            summary = self.current_analysis.get("summary", "<i>No summary available.</i>")
            self.analysis_display.setHtml(intent_html + f"<div style='font-size:16px;'>{summary}</div>")
        elif mode == "sentiment":
            sentiment = self.current_analysis.get("sentiment", {})
            sentiment_html = self.format_sentiment_for_ui(sentiment)
            self.analysis_display.setHtml(sentiment_html)
        elif mode == "keywords":
            keywords = self.current_analysis.get("keywords", [])
            keywords_html = self.format_keywords_for_ui(keywords)
            self.analysis_display.setHtml(keywords_html)



# UI Design
    def style_sheet(self):
        if self.dark_mode:
            return """
            QWidget {
                background-color: #0f172a;
                color: #f1f5f9;
                font-family: 'Segoe UI', sans-serif;
                font-size: 14px;
            }
            QLineEdit, QTextEdit, QTreeWidget {
                background-color: #1e1e1e;
                border: 1px solid #444;
                border-radius: 6px;
                padding: 6px;
                color: white;
            }
            QPushButton, QToolButton {
                background-color: #334155;
                border: 1px solid #475569;
                color: #f1f5f9;
                border-radius: 6px;
                padding: 6px 12px;
            }
            QPushButton:hover, QToolButton:hover {
                background-color: #475569;
            }
            QPushButton:checked, QToolButton:checked {
            background-color: #3b82f6;
            color: white;
            }
            QLabel.section-title {
                font-weight: bold;
                font-size: 18px;
            }
            QFrame#sidebar {
                border: 1px solid #444;
                border-radius: 10px;
                padding: 10px;
            }
            QTreeWidget {
            background-color: #1e1e1e;
            border: 1px solid #444;
            border-radius: 6px;
            padding: 6px;
            color: white;
            }
            QTreeWidget::item {
            padding: 8px;
            border-bottom: 1px solid #333;
            color: #f0f0f0;
            }
            QTreeWidget::item:selected {
            background-color: #3a72d8;
            color: white;
            }
            QHeaderView::section {
            background-color: #1e1e1e;
            color: white;
            border: none;
            }
            QScrollBar:vertical, QScrollBar:horizontal {
            background: #1e1e1e;
            }
            QScrollBar:vertical, QScrollBar:horizontal {
            width: 8px;
            background: transparent;
            }
            QScrollBar::handle:vertical, QScrollBar::handle:horizontal {
            background: #555;
            border-radius: 4px;
            }
            QScrollBar::add-line, QScrollBar::sub-line {
            background: none;
            }
            QSplitter::handle {
            background-color: #334155;
            }
            """
        else:
            return """
            QWidget {
                background-color: #f8f9fa;
                color: #212529;
                font-family: 'Segoe UI', sans-serif;
                font-size: 14px;
            }
            QLineEdit, QTextEdit, QTreeWidget {
                background-color: #ffffff;
                border: 1px solid #d1d5db;
                border-radius: 6px;
                padding: 6px;
            }
            QPushButton, QToolButton{
                background-color: #e9ecef;
                border: 1px solid #ced4da;
                border-radius: 6px;
                padding: 6px 12px;
            }
            QPushButton:hover, QToolButton:hover {
                background-color: #dee2e6;
            }
            QPushButton:checked, QToolButton:checked {
            background-color: #0d6efd;
            color: white;
            }
            QLabel.section-title {
                font-weight: bold;
                font-size: 18px;
            }
            QFrame#sidebar {
                border: 1px solid #ccc;
                border-radius: 10px;
                padding: 10px;
                background-color: #ffffff;
            }
            QTreeWidget {
            background-color: #ffffff;
            border: 1px solid #d0d0d0;
            border-radius: 6px;
            padding: 6px;
            color: #2c3e50;
            }
            QTreeWidget::item {
            padding: 8px;
            border-bottom: 1px solid #ccc;
            color: #2c3e50;
            }
            QTreeWidget::item:selected {
            background-color: #3498db;
            color: white;
            }
              QScrollBar:vertical, QScrollBar:horizontal {
            width: 8px;
            background: transparent;
            }
            QScrollBar::handle:vertical, QScrollBar::handle:horizontal {
            background: #555;
            border-radius: 4px;
            }
            QScrollBar::add-line, QScrollBar::sub-line {
            background: none;
            }
            QSplitter::handle {
            background-color: #e9ecef;
            }
            """

    def toggle_theme(self):
        self.dark_mode = not self.dark_mode
        self.setStyleSheet("")  # Clear
        self.setStyleSheet(self.style_sheet())  # Apply updated style



# Start of Assistant & Insights Panel
    def update_assist_display(self, mode):
        self.chat_input.setVisible(mode == "assistant")
        self.chat_send_btn.setVisible(mode == "assistant")
        if mode == "assistant":
            self.assist_display.setHtml("<b>Hermes is ready. Type or speak your command.</b>")
            self.assist_display.setVisible(True)
            self.chat_input.setVisible(True)
            self.chat_send_btn.setVisible(True)
            self.mic_btn.setVisible(True)
            self.graph_stack.setVisible(False)
            self.graph_nav_container.setVisible(False)
            self.filter_panel.setVisible(False)
        elif mode == "pattern":
            self.assist_display.setVisible(False)
            self.chat_input.setVisible(False)
            self.chat_send_btn.setVisible(False)
            self.graph_stack.setVisible(True)
            self.graph_nav_container.setVisible(True)
            self.filter_panel.setVisible(True)
            self.graph_index = 0  # Reset to first graph
            self.graph_stack.setCurrentIndex(self.graph_index)
            self.graph_title.setText(self.graph_names[self.graph_index])
            self.load_graph_view(self.graph_index)
            self.render_interactive_graph(self.graph_index) 

# Assistant
    def start_voice_input(self):
        if self.listening:
            self.listening = False
            self.voice_worker.stop()
            if self.voice_thread and self.voice_thread.isRunning():
                self.voice_thread.quit()
                self.voice_thread.wait()  # Ensure the thread fully exits before deleting
            self.tts_engine.say("Stopping voice input.")
            self.tts_engine.runAndWait()
            return
        self.listening = True
        self.assist_display.append("<b>Hermes:</b> 🎤 Listening...")
        self.voice_thread = QThread()
        self.voice_worker = VoiceWorker(self.recognizer, self.microphone)
        self.voice_worker.moveToThread(self.voice_thread)
        self.voice_thread.started.connect(self.voice_worker.run)
        self.voice_worker.result.connect(self.handle_voice_command)
        self.voice_worker.error.connect(self.handle_voice_error)
        self.voice_thread.finished.connect(self.cleanup_voice_thread)
        self.voice_thread.start()
     
    def voice_loop(self):
        idle_timeout = 20  # seconds of silence before idle
        while self.listening:
            try:    
                with self.microphone as source:
                    self.recognizer.adjust_for_ambient_noise(source)
                    audio = self.recognizer.listen(source, timeout=10, phrase_time_limit=15)
                    command = self.recognizer.recognize_google(audio)
                    self.handle_voice_command(command)
            except sr.WaitTimeoutError:
                if not self.idle:
                    self.idle = True
                    self.assist_display.append("<b>Hermes:</b> 💤 Idle. Say 'Hermes' to wake me.")
            except sr.UnknownValueError:
                continue  # skip on bad input
            except sr.RequestError:
                self.assist_display.append("<b>Hermes:</b> 🔌 Speech service unavailable.")
                self.listening = False

    def handle_voice_command(self, text):
        self.listening = False
        text = text.strip().lower()
        if not text or len(text) < 3:  
            return
        first_word = text.split()[0]
        if first_word not in self.intent_triggers:
            self.assist_display.append("<b>Hermes:</b> No valid command detected. Ignoring.")
            return
        if hasattr(self, "last_voice_command") and self.last_voice_command == text:
            return  # Prevent repeat
        self.last_voice_command = text
        QTimer.singleShot(3000, self.reset_last_voice_command)
        self.chat_input.setText(text)
        self.chatbot_chat()

    def handle_voice_error(self, msg):
        self.assist_display.append(f"<b>Hermes:</b> 🔌 {msg}")
        self.tts_engine.say(msg)
        self.tts_engine.runAndWait()
        self.listening = False

    def reset_last_voice_command(self):
        self.last_voice_command = None

    def _resume_listening(self):
        if not self.idle:
            self.listening = True
        
    def cleanup_voice_thread(self):
        self.voice_worker.deleteLater()
        self.voice_thread.deleteLater()
        self.voice_worker = None
        self.voice_thread = None

    def stop_voice_input(self):
        self.listening = False
        if self.voice_worker:
            self.voice_worker.stop()
        if self.voice_thread and self.voice_thread.isRunning():
            self.voice_thread.quit()
            self.voice_thread.wait()
        self.cleanup_voice_thread()

    def parse_intent(self, command):
        command = self.normalize_ranges(command)
        handled = False
        command = command.strip().lower()
        if self.idle and self.wake_word in command:
            self.idle = False
            return "I'm awake."
        match = re.search(r'\b(stop|quit|disable|end|shut (?:up)?)\b.*\b(listening|voice|assistant)?\b', command)
        if match:
            self.listening = False
            if self.voice_worker:
                self.voice_worker.stop()
            if self.voice_thread and self.voice_thread.isRunning():
                self.voice_thread.quit()
                self.voice_thread.wait()
                return " Voice input stopped."        
        # ----- Command: Forward Email -----
        match = re.search(r'forward (?:this email )?(?:to )?([\w\s]+)', command)
        if match:
            name = match.group(1).strip()
            email = self.get_email_by_name(name)
            if email:
                self.forward_email(to_name=name)
                return f"Forwarding this to {name.title()}."
            else:
                return f"⚠️ Couldn’t find contact '{name}'."
        # ----- Command: Delete this Email -----
        if re.search(r'(delete|remove) (this )?(email|message|image)', command):
            self.delete_email()
            return "Deleted the currently opened email."
        # ----- Command: Open Email by Index -----
        match = re.search(r'\bopen (?:the )?(?:(first|second|third|fourth|fifth|sixth|seventh|eighth|ninth|tenth)|(\d{1,2})(?:st|nd|rd|th)?)(?:.*?(email|one|message|number)?)?', command)
        if match:
            ordinal_word = match.group(1)
            number_str = match.group(2)
            if ordinal_word:
                index = self.index_word_map.get(ordinal_word.lower(), -1)
            elif number_str:
                index = int(number_str) - 1  
            else:
                index = -1
            if hasattr(self, 'email_item_list') and 0 <= index < len(self.email_item_list):  # Ensure the list exists
                self.open_email_by_index(index)  
                return f"Opening email {index + 1}."
            else:
                return f"Sorry, I couldn’t find email number {index + 1}."
        # ----- Command: Read Email Range -----
        match = re.search(r'\bread (?:emails|messages|email)?(?: from)? (\d+) (?:to|-|through|until|till) (\d+)', command)
        if match:
            start = int(match.group(1))
            end = int(match.group(2))
            self.read_emails_in_range(start, end)
            handled = True
        # ----- Command: Reminder -----    
        if re.search(r'\bremind\b.*\bsend\b', command):
            match = re.search(r'remind(?: me)? to send(?: a)?(?: reply|email)?(?: to)? ([\w\s]+?)(?:\s+(?:in|after|at|next|tomorrow)\b|$)',  command)
            name = match.group(1).strip().title() if match else "someone"
            delay_seconds = self.parse_reminder_time(command)
            if delay_seconds <= 0:
                msg = "⚠️ Couldn't parse a valid reminder time."
                self.assist_display.append(f"<b>Hermes:</b> {msg}")
                self.tts_engine.say("Couldn't understand the reminder time.")
                self.tts_engine.runAndWait()
                return None
            if self.pending_timer and self.pending_timer.isActive():
                self.pending_timer.stop()
            trigger_time = datetime.now() + timedelta(seconds=delay_seconds)
            time_str = trigger_time.strftime('%I:%M %p').lstrip('0')
            confirmation = f"⏰ Reminder set to reply to {name} at around {time_str}."
            self.tts_engine.say(confirmation)
            self.tts_engine.runAndWait()
            self.assist_display.append(f"<b>Hermes:</b> {confirmation}")
            def reminder_trigger():
                reminder_text = f"⏰ Reminder: You wanted to reply to <b>{name}</b>."
                self.message_sound.play()
                self.assist_display.append(reminder_text)
                self.tts_engine.say(f"Reminder: You wanted to reply to {name}.")
                self.tts_engine.runAndWait()
            self.pending_timer = QTimer()
            self.pending_timer.setSingleShot(True)
            self.pending_timer.timeout.connect(reminder_trigger)
            self.pending_timer.start(int(delay_seconds * 1000)) 
            return None
        # ----- Command: Reply to Email (Generate but Don't Send) -----
        if re.search(r'\b(reply|respond|answer)\b.*\b(this|that)? email\b', command) and "send" not in command:
            self.generate_email_reply(schedule_time="now", auto_send=False)
            return None
        # ----- Command: Send Email (Now or After Some Time) ----
        match = re.search(r"(?:send|schedule).*reply(?:\s*(?:in|after|at)\s+(.+))?", command)
        if match:
            time_str = match.group(1) or "now"
            time_str = time_str.strip().lower()
            # Handle bad captures like "n", "a", etc.
            if time_str in ("n", "a", "the", ""):
                time_str = "now"
            auto_send = True
            self.generate_email_reply(schedule_time=time_str, auto_send=auto_send)
            return f"📨 Preparing reply to be sent {'at ' + time_str }."
        # ----- Command: Send ----
        if re.search(r'\bsend (it )?now\b', command, re.IGNORECASE) and self.pending_reply:
            reply = self.pending_reply
            self.subject_input.setText(reply["subject"])
            self.body_input.setPlainText(reply["body"])
            self.send_email()
            self.pending_reply = None
            return "📨 Sent the drafted reply."
        # -------- Cancel Scheduled Send ----------------
        if re.search(r'\bcancel\b.*(scheduled|delayed|pending).*(reply|send)', command) and self.pending_timer:
            if self.pending_timer.isActive():
                self.pending_timer.stop()
            self.pending_timer = None
            self.pending_reply = None
            return "⛔ Scheduled send has been cancelled."
        if "didn’t catch that" in command or "did not catch" in command:
            return None
        if not handled:
            return "Sorry, I didn’t catch that. Try again?"
        
    def normalize_ranges(self, text):
        match = re.search(r'\bread (?:email|emails)?(?: from)? (\d{2,3})$', text)
        if match:
            digits = match.group(1)
            if len(digits) == 2:
                return text.replace(digits, f"{digits[0]} to {digits[1]}")
            elif len(digits) == 3:
                return text.replace(digits, f"{digits[0]} to {digits[2]}")
        return text
        
    def chatbot_chat(self):
        text = self.chat_input.text().strip()
        if not text:
            return
        self.assist_display.append(f"<b>You:</b> {text}")
        if "start voice" in text.lower():
            self.start_voice_input()
            return
        response = self.parse_intent(text)
        if response:  # avoid speaking if parse_intent handled TTS internally
            self.assist_display.append(f"<b>Hermes:</b> {response}")
            self.tts_engine.say(response)
            self.tts_engine.runAndWait()

    def open_email_by_index(self, index, folder=None):
        if isinstance(index, str):
            index = self.index_word_map.get(index.lower(), -1)
        if folder:
            folder = folder.lower()
            items = self.folder_email_map.get(folder, [])
        else:
            items = self.email_item_list
        if index < 0 or index >= len(items):
            self.email_viewer.setText(f"[❌] No email at position {index + 1} in {folder or 'all folders'}")
            return
        item = items[index]
        self.email_tree.setCurrentItem(item)
        self.current_selected_index = index
        folder_name, mail_id = item.data(0, Qt.UserRole)
        print(f"[📂] Opening #{index + 1}: Folder={folder_name}, ID={mail_id}")
        self.current_subject = item.text(0)
        sender_raw = item.text(1)
        self.store_contact_from_line(sender_raw)
        name, email = parseaddr(sender_raw)
        self.selected_contact_email = email
        self.display_email(item)

    def read_emails_in_range(self, start, end):
        if start < 1 or end < 1 or start > len(self.email_item_list) or end > len(self.email_item_list):
            self.tts_engine.say("Invalid range of emails.")
            self.tts_engine.runAndWait()
            return
        self.tts_engine.say(f"Reading emails from {start} to {end}.")
        self.tts_engine.runAndWait()
        self.current_email_index = start - 1
        self.is_reading = True
        while self.current_email_index < end and self.is_reading:
            self.open_email_by_index(self.current_email_index)  # Opens and highlights the email
            item = self.email_item_list[self.current_email_index]
            subject = item.text(0)
            sender_raw = item.text(1)
            # Extract and store contact name from raw sender text
            self.store_contact_from_line(sender_raw)
            name, email = parseaddr(sender_raw)
            display_name = name.strip() or email
            self.tts_engine.say(f"Email {self.current_email_index + 1}: From {display_name}({email}), Subject: {subject}")
            self.tts_engine.runAndWait()
            self.assist_display.append(f"<b>Hermes:</b> Email {self.current_email_index + 1}: From {display_name}({email}), Subject: {subject}")
            # Now read full email body
            self.read_email_body()
            self.current_email_index += 1
        self.tts_engine.say("Finished reading the emails.")
        self.tts_engine.runAndWait()
        self.is_reading = False

    def read_email_body(self):
        body_text = self.email_viewer.toPlainText()
        if body_text:
            parts = body_text.split("\n\n", 1)  # Split headers from body
            body = parts[1] if len(parts) == 2 else body_text
            self.tts_engine.say(body)
            self.tts_engine.runAndWait()
        else:
            self.tts_engine.say("This email has no content.")
            self.tts_engine.runAndWait()

    def generate_email_reply(self, schedule_time="now", auto_send=False):
        email_body = self.email_viewer.toPlainText()
        if not email_body.strip():
            self.assist_display.append("<b>Hermes:</b> ❌ No email is selected.")
            self.tts_engine.say("Please open an email first.")
            self.tts_engine.runAndWait()
            return
        # Extract sender's email to reply to
        email_parts = email_body.split("\n")
        from_line = ""
        for line in email_parts:
            if line.startswith("From:"):
                from_line = line
                break
        # Parse the email address from the From line
        name, recipient_email = parseaddr(from_line.replace("From:", "").strip())
        if not recipient_email:
            self.assist_display.append("<b>Hermes:</b> ❌ Couldn’t find a valid recipient email.")
            return
        # Get the subject from current_subjec
        subject = "Re: " + getattr(self, "current_subject", "")
        prompt = (
            f"You are an AI assistant helping a user reply to emails. "
            f"Generate **one** formal, polite, and context-aware reply **on behalf of the user**, as if they are responding personally."
            f"Do not present multiple options or ask questions. The reply should be complete, professional, "
            f"with appropriate tone, and should not include any placeholders like 'your name'. "
            f"Ensure the message is clear, and professional, and has a proper closing. "
            f"This reply should be suitable to send after {schedule_time}.\n\n"
            f"--- EMAIL TO REPLY TO ---\n{email_body}\n--- EMAIL END ---"
        )
        self.assist_display.append("<b>Hermes:</b> ✍️ Writing your reply...")
        try:
            reply = self.send_to_gemini(prompt)
            self.generated_reply = reply
            if auto_send and schedule_time == "now":
                self.subject_input.setText(subject)
                self.body_input.setPlainText(reply)
                self.send_email()
                self.assist_display.append("<b>Hermes:</b> ✅ Reply generated and sent.")
                self.tts_engine.say("Reply has been generated and sent.")
            elif auto_send:
                delay_seconds = self.parse_schedule_time(schedule_time)
                if delay_seconds <= 0:
                    # fallback
                    self.assist_display.append("<b>Hermes:</b> ⚠️ Invalid or zero delay time. Sending immediately.")
                    self.tts_engine.say("Invalid time detected. Sending now.")
                    self.tts_engine.runAndWait()
                    self.subject_input.setText(subject)
                    self.body_input.setPlainText(reply)
                    self.send_email()
                    self.assist_display.append("<b>Hermes:</b> ✅ Reply generated and sent.")
                    return  
                else:
                    self.schedule_delayed_send(subject, reply, delay_seconds)
                    self.pending_reply = None
            else: 
                self.pending_reply = {
                    "subject": subject,
                    "to": recipient_email,
                    "body": reply
                }
                self.subject_input.setText(subject)
                self.body_input.setPlainText(reply)
                self.assist_display.append("<b>Hermes:</b> Reply drafted. Say 'send it now' to send.")
                self.tts_engine.say("Reply drafted. Say 'send it now' to send.")
                self.tts_engine.runAndWait()     
        except Exception as e:
            self.assist_display.append(f"<b>Hermes:</b> ❌ Failed to generate reply: {e}")
            self.tts_engine.say("Something went wrong while generating the reply.")
            self.tts_engine.runAndWait()

    def send_to_gemini(self, prompt):
        genai.configure(api_key=self.api_key) 
        model = genai.GenerativeModel(
            "gemini-2.0-flash",  
            system_instruction=(
            "You generate **one** professional, ready-to-send email replies *on behalf of the user*. "
            "Do not offer options, ask questions, or provide multiple drafts. "
            "The reply should be polite, formal, and contextually aware, with appropriate salutation and closing. "
            "Do not mention yourself (the assistant) or use placeholders like 'your name'."
            )
        )
        response = model.generate_content(prompt)
        return response.text.strip()

    def parse_schedule_time(self, time_str):
        try:
            time_str = time_str.lower().strip()
            def normalize_numbers(text):
                try:
                    words = text.split()
                    for i in range(len(words)):
                        try:
                            number = w2n.word_to_num(' '.join(words[i:i+2]))
                            words[i:i+2] = [str(number)]
                            return ' '.join(words)
                        except:
                            continue
                    return text
                except:
                    return text
            time_str = normalize_numbers(time_str)
            # Try relative patterns like: "in 10 minutes", "after 30 mins"
            match = re.search(r'(in|after)?\s*(\d+)\s*(minute|minutes|min|mins)?', time_str)
            if match:
                num = int(match.group(2))
                if 1 <= num <= 60:
                    print(f"[DEBUG] Parsed delay: {num * 60} seconds.")
                    return num * 60  # seconds
                else:
                    raise ValueError("Only delays between 1 to 60 minutes are allowed.")
            raise ValueError("Invalid format. Only minute-based delays are supported.")
        except Exception as e:
            print(f"[!] Time parsing failed: {e}")
            return 0  # fallback to immediate send
        
    def schedule_delayed_send(self, subject, body, delay_seconds):
        def send_delayed_email():
            self.subject_input.setText(subject)
            self.body_input.setPlainText(body)
            self.send_email()
            self.assist_display.append("<b>Hermes:</b> ✅ Reply sent as scheduled.")
            self.tts_engine.say("Reply sent as scheduled.")
            self.tts_engine.runAndWait()
            self.pending_reply = None
            self.pending_timer = None
        if delay_seconds <= 0:
            self.assist_display.append("<b>Hermes:</b> ⚠️ Invalid or zero delay time. Sending immediately.")
            self.tts_engine.say("Invalid time detected. Sending now.")
            self.tts_engine.runAndWait()
            send_delayed_email()
            return
        self.assist_display.append(f"<b>Hermes:</b> 🕒 Will send in {delay_seconds // 60} min(s). You can cancel or edit before that.")
        self.tts_engine.say(f"Okay, sending this in {delay_seconds // 60} minutes.")
        self.tts_engine.runAndWait()
        self.pending_timer = QTimer()
        self.pending_timer.setSingleShot(True)
        self.pending_timer.timeout.connect(send_delayed_email)
        self.pending_timer.start(int(delay_seconds * 1000))

    def cancel_scheduled_reply(self):
        if self.pending_timer and self.pending_timer.isActive():
            self.pending_timer.stop()
            self.pending_reply = None
            self.pending_timer = None
            self.assist_display.append("<b>Hermes:</b> ❌ Scheduled reply canceled.")
            self.tts_engine.say("Scheduled reply canceled.")
            self.tts_engine.runAndWait()
        else:
            self.assist_display.append("<b>Hermes:</b> ⚠️ No scheduled reply to cancel.")
            self.tts_engine.say("There is no scheduled reply to cancel.")
            self.tts_engine.runAndWait()

    def parse_reminder_time(self, text: str) -> int:
        return self.parse_schedule_time(text)



# Time Pattern
    def change_graph(self, delta):
       self.graph_index = (self.graph_index + delta) % len(self.graph_names)
       self.load_graph_view(self.graph_index)

    def export_current_graph(self):
        index = self.graph_index
        if not hasattr(self, "graph_views") or not self.graph_views:
            return
        
        view = self.graph_views[index]
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Export Graph", f"{self.graph_names[index]}.html", "HTML Files (*.html);;All Files (*)"
        )
        if file_path:
            # Re-render and save the graph to the selected path
            name = self.graph_names[index]
            fig = self.generate_figure(self.graph_names[index])  # You'll define this to return the current plotly Figure
            fig.write_html(file_path)
            QMessageBox.information(self, "Export Successful", f"Graph exported to:\n{file_path}")

    def generate_figure(self, name):
        fig = go.Figure()
        emails = self.all_emails if hasattr(self, "all_emails") else []
        if not emails:
            fig.add_annotation(text="📭 No email data available", showarrow=False)
            return fig
        base_layout = dict(
            font=dict(family="Segoe UI", size=14),
            title_font=dict(size=22, family="Arial Black"),
            plot_bgcolor="rgba(240,240,240,0.95)",
            paper_bgcolor="white",
            margin=dict(l=40, r=40, t=60, b=40)
        )
        if "Top Senders" in name:
            sender_counts = Counter(email['from'] for email in emails)
            top = sender_counts.most_common(10)
            senders, counts = zip(*top)
            fig = px.bar(x=senders, y=counts, title="Top Email Senders")
        elif "Email Volume" in name:
            days = defaultdict(int)
            for email in emails:
                day = email["date"].strftime("%A")
                days[day] += 1     
            ordered_days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
            y = [days.get(day, 0) for day in ordered_days]
            fig = px.bar(
                x=ordered_days, y=y,
                title="📈 Email Volume by Day",
                color=ordered_days,
                color_discrete_sequence=px.colors.sequential.Blues,
                labels={'x': 'Day', 'y': 'Email Count'}
            )
            fig.update_traces(marker_line=dict(width=1, color='DarkSlateGrey'))
        elif "Sentiment" in name:
            dates = [email["date"].date() for email in emails if "sentiment" in email]
            sentiments = [email["sentiment"] for email in emails if "sentiment" in email]
            if dates and sentiments:
                df = pd.DataFrame({"date": dates, "sentiment": sentiments})
                df = df.groupby("date").mean().reset_index()
                def sentiment_to_emoji(score):
                    if score >= 0.7: return "🤩"
                    elif score >= 0.4: return "😄"
                    elif score >= 0.2: return "😊"
                    elif score > -0.2: return "😐"
                    elif score > -0.4: return "😕"
                    elif score > -0.7: return "😠"
                    else: return "💢"
                df["emoji"] = df["sentiment"].apply(sentiment_to_emoji)
                fig = px.line(df, x="date", y="sentiment", title="📊 Average Sentiment Over Time", markers=True)
                for i, row in df.iterrows():
                    fig.add_annotation(x=row["date"], y=row["sentiment"], text=row["emoji"], showarrow=False, yshift=10)
            else:
                fig.add_annotation(text="No sentiment data available", showarrow=False)
        elif "Keywords" in name:
            keyword_map = defaultdict(int)
            for email in emails:
                doc = self.nlp(email["body"])
                for ent in doc.ents:
                    if ent.label_ in ["ORG", "PRODUCT", "PERSON"]:
                        keyword_map[ent.text] += 1
            if keyword_map:
                top_keywords = dict(sorted(keyword_map.items(), key=lambda x: x[1], reverse=True)[:10])
                fig = px.bar(
                    x=list(top_keywords.keys()), y=list(top_keywords.values()),
                    title="🏷️ Top Keywords (NER)",
                    color=list(top_keywords.keys()),
                    color_discrete_sequence=px.colors.qualitative.Set3,
                    labels={'x': 'Keyword', 'y': 'Frequency'}
                )
                fig.update_traces(marker_line=dict(width=1, color='DarkSlateGrey'))
            else:
                fig.add_annotation(text="No keywords found", showarrow=False)
        else:
            fig.add_annotation(text="Graph not implemented.", showarrow=False)
        fig.update_layout(**base_layout)
        return fig
    
    def open_time_filter_dialog(self):
        senders = list({email["from"] for email in self.all_emails}) if hasattr(self, "all_emails") else []
        dialog = TimeFilterDialog(senders, self)
        if dialog.exec_() == QDialog.Accepted:
            filters = dialog.get_filters()
            print("Filters selected:", filters)
            # You can now apply these filters to self.all_emails and re-render graphs
            self.apply_filters_and_refresh_graphs(filters)

    def load_graph_view(self, index):
        self.graph_index = index % len(self.graph_names)
        self.graph_stack.setCurrentIndex(self.graph_index)
        self.graph_title.setText(self.graph_names[self.graph_index])
        self.render_interactive_graph(self.graph_index)

    def show_prev_graph(self):
        self.graph_index = (self.graph_index - 1) % len(self.graph_names)
        self.load_graph_view(self.graph_index)

    def show_next_graph(self):
        self.graph_index = (self.graph_index + 1) % len(self.graph_names)
        self.load_graph_view(self.graph_index)

    def render_interactive_graph(self, index, filters=None):
        name = self.graph_names[index]
        if not hasattr(self, "all_emails") or not self.all_emails:
            fig = go.Figure()
            fig.add_annotation(text="📭 No email data available", showarrow=False)
        else:
            emails = self.all_emails
            # ✅ Apply filters
            if filters:    
                if filters.get("start_date"):
                    emails = [e for e in emails if e["date"].date() >= filters["start_date"]]  
                if filters.get("end_date"):
                    emails = [e for e in emails if e["date"].date() <= filters["end_date"]]
                if filters.get("sender"):
                    emails = [e for e in emails if filters["sender"].lower() in e["from"].lower()]
            base_layout = dict(
                font=dict(family="Segoe UI", size=14),
                title_font=dict(size=22, family="Arial Black"),
                plot_bgcolor="rgba(240,240,240,0.95)",
                paper_bgcolor="white",
                margin=dict(l=40, r=40, t=60, b=40)
            )
            def sentiment_to_emoji(score):
                if score >= 0.7:
                    return "🤩"  # ecstatic
                elif score >= 0.4:
                    return "😄"  # very happy
                elif score >= 0.2:
                    return "😊"  # happy
                elif score > -0.2:
                    return "😐"  # neutral
                elif score > -0.4:
                    return "😕"  # slightly negative
                elif score > -0.7:
                    return "😠"  # angry
                else:
                    return "💢"  # furious

            if "Top Senders" in name:
                from collections import Counter
                sender_counts = Counter(email['from'] for email in emails)
                top = sender_counts.most_common(10)
                if top:
                    senders, counts = zip(*top)
                    fig = px.bar(
                        x=senders, y=counts, title="📬 Top Email Senders",
                        color=senders, color_discrete_sequence=px.colors.qualitative.Vivid,
                        labels={'x': 'Sender', 'y': 'Email Count'}
                    )
                    fig.update_traces(marker_line=dict(width=1, color='DarkSlateGrey'))
                else:
                    fig = go.Figure()
                    fig.add_annotation(text="No sender data in selected range", showarrow=False)
            elif "Email Volume" in name:
                days = defaultdict(int)
                for email in emails:
                    day = email["date"].strftime("%A")  # 'Monday', etc.
                    days[day] += 1
                ordered_days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
                y = [days.get(day, 0) for day in ordered_days]
                fig = px.bar(
                    x=ordered_days, y=y, title="📈 Email Volume by Day",
                    color=ordered_days, color_discrete_sequence=px.colors.sequential.Blues,
                    labels={'x': 'Day', 'y': 'Email Count'}
                )
                fig.update_traces(marker_line=dict(width=1, color='DarkSlateGrey'))
            elif "Sentiment" in name:   
                sentiments = [
                    (email["date"].date(), email["sentiment"])
                    for email in emails if "sentiment" in email
                ]
                if sentiments:
                    dates, scores = zip(*sentiments)
                    df = pd.DataFrame({"date": dates, "sentiment": scores})
                    df = df.groupby("date").mean().reset_index()
                    df["emoji"] = df["sentiment"].apply(sentiment_to_emoji)
                    fig = px.line(
                        df, x="date", y="sentiment", title="📊 Average Sentiment Over Time",
                        markers=True, line_shape="spline"
                    )
                    for i, row in df.iterrows():
                        fig.add_annotation(
                            x=row["date"], y=row["sentiment"],
                            text=row["emoji"],
                            showarrow=False,
                            yshift=10
                        )
                else:
                    fig = go.Figure()
                    fig.add_annotation(text="No sentiment data found", showarrow=False)
            elif "Keywords" in name:
                # Placeholder: Extract NER keywords over time
                keyword_map = defaultdict(lambda: defaultdict(int))
                for email in emails:
                    email_date = email["date"].date()
                    doc = self.nlp(email["body"])
                    for ent in doc.ents:
                        if ent.label_ in ["ORG", "PRODUCT", "PERSON"]:
                            keyword_map[email_date][ent.text] += 1
                # Flatten data into a long dataframe: date | keyword | count
                rows = []
                for date, keywords in keyword_map.items():
                    for kw, count in keywords.items():
                        rows.append({"date": date, "keyword": kw, "count": count})
                if rows:
                    df = pd.DataFrame(rows)
                    top_keywords = df.groupby("keyword")["count"].sum().nlargest(5).index  # Top 5 overall
                    df = df[df["keyword"].isin(top_keywords)]
                    fig = px.line(
                        df, x="date", y="count", color="keyword",
                        title="🏷️ Top Keywords (NER) Over Time", markers=True
                    )
                else:
                    fig = go.Figure()
                    fig.add_annotation(text="No keywords found", showarrow=False)
            else:
                fig = go.Figure()
                fig.add_annotation(text="Graph logic not implemented yet", showarrow=False)
            fig.update_layout(**base_layout)
        with tempfile.NamedTemporaryFile(delete=False, suffix=".html") as tmpfile:
            fig.write_html(tmpfile.name)
            self.graph_views[index].load(QUrl.fromLocalFile(tmpfile.name))

    def apply_filters_and_refresh_graphs(self, filters):
        start = self.start_date_edit.date().toPyDate()
        end = self.end_date_edit.date().toPyDate()
        sender = self.sender_filter.text().strip()
        # Validate date range
        if start > end:
            QMessageBox.warning(self, "Invalid Filter", "Start date must be before or equal to end date.")
            return
        # Normalize sender input
        sender = sender.lower() if sender else None
        # Build filters dictionary
        filters = {
            "start_date": start,
            "end_date": end,
            "sender": sender
        }
        print("📊 Applying filters:", filters)
        # Store current filters if needed elsewhere
        self.active_filters = filters
        # Refresh current graph with these filters
        self.render_interactive_graph(self.graph_index, filters=filters)




# Start of Email List Integration
    def update_email_contact_list(self, raw_email_text):
        if not hasattr(self, 'contact_dict'):
            self.load_contact_dict() # Initialize only once
        # Extract lines like "From: John Smith <john@example.com>"
        match = re.search(r"From:\s*(.+)", raw_email_text)
        if match:
            from_line = match.group(1).strip()
            name, email = parseaddr(from_line)
            if email:  # basic email check
                cleaned_name = name.strip() or email.split('@')[0].title()
                if cleaned_name not in self.contact_dict:
                    self.contact_dict[cleaned_name] = email
                    print(f"[📥] Stored contact: {cleaned_name} → {email}")
                    self.save_contact_dict()
                
    def get_email_by_name(self, name_query):
        name_query = name_query.lower()
        matches = []
        for name, email in self.contact_dict.items():
            if name_query in name.lower():
                matches.append((name, email))
        if len(matches) == 1:
            return matches[0][1]
        elif len(matches) > 1:
            print(f"[⚠️] Multiple matches for '{name_query}': {[n for n, e in matches]}")
            return matches[0][1]  # Return the first match anyway
        else:
            print(f"[❌] No contact found for '{name_query}'")
            return None
        
    def store_contact_from_line(self, raw_line):
        name, email = parseaddr(raw_line)
        if not email or self.email_address.lower() in email.lower():
            return
        name = name.strip()
        if not name:
            # Extract from email before @
            local_part = email.split('@')[0]
            # Split on common separators
            parts = re.split(r"[._\-+]", local_part)
            for part in parts:
                if part and part.lower() not in ("mail", "info", "contact", "team", "support", "noreply"):
                    name = part.title()
                    break
            # Still empty? Use the whole local part capitalized
            if not name:
                name = local_part.title()
        # Ensure uniqueness (avoid overwriting)
        if name not in self.contact_dict:
            self.contact_dict[name] = email
            print(f"[📥] Stored contact: {name} → {email}")
        elif self.contact_dict[name] != email:
            alt_name = f"{name} ({email.split('@')[0]})"
            self.contact_dict[alt_name] = email
        self.save_contact_dict()
        print(f"[📥] Stored contact: {name} → {email}")

    def load_contact_dict(self):
        if os.path.exists("contacts.json"):
            with open("contacts.json", "r") as f:
                self.contact_dict = json.load(f)
        else:
            self.contact_dict = {}
    
    def save_contact_dict(self):
        with open("contacts.json", "w") as f:
            json.dump(self.contact_dict, f, indent=2)




class SmartSummarizer:
    def __init__(self, model_name="facebook/bart-large-cnn"):
        self.device = "cuda" if torch.cuda.is_available() else "cpu"
        self.tokenizer = BartTokenizer.from_pretrained(model_name)
        self.model = BartForConditionalGeneration.from_pretrained(model_name).to(self.device)

    def generate_summary(self, text, doc, sentences, max_length=130, min_length=30):
        try:
            inputs = self.tokenizer.encode("summarize: " + text, return_tensors="pt", max_length=1024, truncation=True)
            inputs = inputs.to(self.device)
            summary_ids = self.model.generate(
                inputs,
                max_length=max_length,
                min_length=min_length,
                length_penalty=2.0,
                num_beams=4,
                early_stopping=True
            )
            transformer_summary = self.tokenizer.decode(summary_ids[0], skip_special_tokens=True)

            # Check if summary is bad or unrelated
            if (
                not transformer_summary.strip() or
                "CNN.com" in transformer_summary or
                "plain version of this email" in transformer_summary.lower()
            ):
                raise ValueError("Bad transformer summary fallback to Sumy")

            return transformer_summary.strip()

        except Exception as e:
            print(f"[Fallback] Transformer summarization failed: {e}")
            return self._fallback_sumy_summary(text, doc, sentences)

    def _fallback_sumy_summary(self, text, doc, sentences):
        key_entities = {
            ent.text for ent in doc.ents
            if ent.label_ in ("ORG", "PERSON", "DATE", "PRODUCT", "EVENT", "CARDINAL")
        }

        parser = PlaintextParser.from_string(text, SumyTokenizer("english"))
        summarizer = LexRankSummarizer()
        sumy_summary = summarizer(parser.document, sentences_count=2)

        important_sents = []
        for s in sumy_summary:
            cleaned = str(s).strip()
            if any(ent in cleaned for ent in key_entities) or "apolog" in cleaned.lower() or "confirm" in cleaned.lower():
                important_sents.append(cleaned)

        if len(important_sents) < 2:
            for sent in sentences:
                if any(e in sent for e in key_entities) or "apolog" in sent.lower() or "confirm" in sent.lower():
                    if sent.strip() not in important_sents:
                        important_sents.append(sent.strip())
                if len(important_sents) >= 2:
                    break

        if not important_sents:
            return (sentences[0] if sentences else text[:200] + "...").strip()

        return " ".join(important_sents).strip()

class AttachmentPreviewDialog(QDialog):
    def __init__(self, attachment_part, parent=None):
        super().__init__(parent)
        self.attachment_part = attachment_part
        self.setWindowTitle("Attachment Preview")
        self.setGeometry(200, 200, 800, 600)
        self.setModal(True)
        self.setLayout(QVBoxLayout())

        # Create a QLabel to show the image or file preview
        self.preview_label = QLabel(self)
        self.layout().addWidget(self.preview_label)

        # Load the attachment into the QLabel
        self.load_attachment()

    def load_attachment(self):
        try:
            img_data = self.attachment_part.get_payload(decode=True)
            pixmap = QPixmap()
            pixmap.loadFromData(img_data)

            # Check if the attachment is an image
            if not pixmap.isNull():
                self.preview_label.setPixmap(pixmap.scaled(800, 600, Qt.KeepAspectRatio))
            else:
                # For non-image attachments, provide a generic message
                self.preview_label.setText("Unable to display this file format.")
        except Exception as e:
            self.preview_label.setText(f"Error loading attachment: {e}")

class NotificationManager:
    def __init__(self, parent):
        self.parent = parent
        self.notified_emails = set()
        self.seen_emails = set() 
        self.frame = QFrame(parent)
        self.frame.setObjectName("notification_frame")
        self.frame.setStyleSheet("""
            QFrame#notification_frame {
                background-color: #3498db;
                border-radius: 8px;
            }
            QLabel {
                font-size: 14px;
                font-family: 'Segoe UI', Arial, sans-serif;
                color: white;
                background: transparent;
            }
            QPushButton {
                background: transparent;
                border: none;
                color: white;
                font-size: 16px;
                font-weight: bold;
            }
        """)
        self.frame.setVisible(False)
        self.frame.setFixedHeight(36)

        layout = QHBoxLayout(self.frame)
        layout.setContentsMargins(12, 4, 12, 4)
        layout.setSpacing(10)

        self.label = QLabel("📬 New email from someone@example.com")
        self.label.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        self.label.setAlignment(Qt.AlignVCenter | Qt.AlignLeft)

        self.close_btn = QPushButton("x")
        self.close_btn.setFixedSize(20, 20)
        self.close_btn.clicked.connect(self.hide_notification)

        layout.addWidget(self.label)
        layout.addStretch()
        layout.addWidget(self.close_btn)

        self.timer = QTimer(parent)
        self.timer.setSingleShot(True)
        self.timer.timeout.connect(self.hide_notification)
    
    def maybe_notify(self, sender_email, folder, mail_id, category):
        email_key = f"{folder}:{mail_id}"
        if email_key in self.seen_emails:
            return
        self.seen_emails.add(email_key)
        if category.lower() == "sent":
            return  # Don't notify for sent mails
        if email_key in self.notified_emails:
            return  # Already seen, don't notify
        self.notify(sender_email, folder, mail_id)


    def notify(self, sender_email, folder, mail_id):
        email_key = f"{folder}:{mail_id}"
        self.notified_emails.add(email_key)
        if len(sender_email) > 40:
            sender_email = sender_email[:37] + "..."

        self.label.setText(f"📬 New email from {sender_email}")
        self.frame.adjustSize()
        self.frame.setMinimumWidth(360)
        self.frame.move((self.parent.width() - self.frame.width()) // 2, 10)
        self.frame.raise_()
        self.frame.setVisible(True)
        self.timer.start(15000)

        if hasattr(self.parent, 'notification_sound'):
            self.parent.notification_sound.play()

        self.frame.mousePressEvent = lambda event: self.parent.open_email_from_notification(folder, mail_id)


    def hide_notification(self):
        self.frame.setVisible(False)

    def clear_state(self):
        self.notified_emails.clear()
        self.seen_emails.clear()

class VoiceWorker(QObject):
    result = pyqtSignal(str)
    error = pyqtSignal(str)

    def __init__(self, recognizer, microphone):
        super().__init__()
        self.recognizer = recognizer
        self.microphone = microphone
        self._running = True

    def stop(self):
        self._running = False

    def run(self):
        while self._running:
            try:
                with self.microphone as source:
                    self.recognizer.adjust_for_ambient_noise(source)
                    audio = self.recognizer.listen(source, timeout=5, phrase_time_limit=8)
                    command = self.recognizer.recognize_google(audio)
                    self.result.emit(command)
            except sr.WaitTimeoutError:
                continue
            except sr.UnknownValueError:
                continue
            except sr.RequestError:
                self.error.emit("Speech service unavailable.")
                break

class TimeFilterDialog(QDialog):
    def __init__(self, senders, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Time Filter")
        self.setMinimumWidth(300)
        layout = QFormLayout(self) 
        # Date range
        self.start_date = QDateEdit(calendarPopup=True)
        self.start_date.setDate(QDate.currentDate().addMonths(-1))
        layout.addRow("From:", self.start_date)
        self.end_date = QDateEdit(calendarPopup=True)
        self.end_date.setDate(QDate.currentDate())
        layout.addRow("To:", self.end_date)
        # Day filter
        self.day_filter = QComboBox()
        self.day_filter.addItems(["All Days", "Weekdays Only", "Weekends Only"])
        layout.addRow("Day Type:", self.day_filter)
        # Sender filter
        self.sender_filter = QComboBox()
        self.sender_filter.addItem("All Senders")
        self.sender_filter.addItems(sorted(senders))
        layout.addRow("Sender:", self.sender_filter)
        # Buttons
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        layout.addRow(button_box)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)

    def get_filters(self):
        return {
            "start_date": self.start_date.date().toPyDate(),
            "end_date": self.end_date.date().toPyDate(),
            "day_type": self.day_filter.currentText(),
            "sender": self.sender_filter.currentText()
        }

class EmailLoaderThread(QThread):
    result = pyqtSignal(list)
    error = pyqtSignal(str)

    def __init__(self, email_address, app_password):
        super().__init__()
        self.email_address = email_address
        self.app_password = app_password

    def run(self):
        try:
            imap = imaplib.IMAP4_SSL("imap.gmail.com")
            imap.login(self.email_address, self.app_password)

            results = []
            folders = {
                "Inbox": "inbox",
                "Sent": '"[Gmail]/Sent Mail"',
                "Spam": '"[Gmail]/Spam"',
            }

            for category, folder in folders.items():
                try:
                    imap.noop()
                    imap.select(folder)
                except imaplib.IMAP4.abort as e:
                    if "Too many simultaneous connections" in str(e) or "Rate Limit" in str(e):
                        self.error.emit("🚫 Gmail IMAP limit hit: Too many connections or rate-limited.")
                        return
                    imap.logout()
                    time.sleep(5)
                    imap = imaplib.IMAP4_SSL("imap.gmail.com")
                    imap.login(self.email_address, self.app_password)
                    imap.select(folder)
                status, messages = imap.search(None, "ALL")
                ids = messages[0].split()[-8:]
                mails = []
                for mail_id in reversed(ids):
                    if not mail_id:
                        continue
                    try:
                        imap.noop()
                        res, msg_data = imap.fetch(mail_id, "(BODY.PEEK[HEADER.FIELDS (FROM TO SUBJECT DATE)])")
                        time.sleep(0.1)
                    except imaplib.IMAP4.abort as e:
                        self.error.emit("❌ Connection aborted during FETCH.")
                        return
                    except imaplib.IMAP4.error as e:
                        if "Too many simultaneous connections" in str(e) or "Rate Limit" in str(e):
                            self.error.emit("🚫 Gmail IMAP limit hit during fetch.")
                        else:
                            self.error.emit(f"❌ IMAP fetch error: {e}")
                        return
                    for response_part in msg_data:
                        if isinstance(response_part, tuple):
                            msg = email.message_from_bytes(response_part[1])
                            subject = msg["subject"]
                            sender = msg["from"]
                            to = msg.get("to", "")
                            mails.append((folder, mail_id, subject, sender, to))
                results.append((category, mails))
            imap.noop()
            imap.logout()
            self.result.emit(results)
        except Exception as e:
            self.error.emit(str(e))

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = EmailDashboard()
    window.show()
    sys.exit(app.exec_())