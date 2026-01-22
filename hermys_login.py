import sys
from PyQt5.QtWidgets import (
    QApplication, QWidget, QLabel, QLineEdit, QPushButton, QVBoxLayout,
    QHBoxLayout, QTextEdit
)
from PyQt5.QtGui import QFont, QPixmap
from PyQt5.QtCore import Qt, QPropertyAnimation
import email_sender1  # Replace with your actual module


class LoginWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("HERMYS - Login")
        self.setGeometry(100, 100, 1200, 700)
        self.setStyleSheet(self.stylesheet())
        self.build_ui()

    def build_ui(self):
        main_layout = QHBoxLayout(self)

        # LEFT SIDE
        left_panel = QWidget()
        left_panel.setFixedWidth(550)
        left_panel.setStyleSheet("background-color: rgba(255, 255, 255, 0.02); border-radius: 30px;")
        left_layout = QVBoxLayout(left_panel)
        left_layout.setContentsMargins(40, 40, 40, 40)
        left_layout.setSpacing(30)

        branding = QLabel("HERMYS")
        branding.setStyleSheet("font-size: 36px; font-weight: bold; color: white;")
        branding.setAlignment(Qt.AlignLeft)

        ai_image = QLabel()
        pixmap = QPixmap("C:/Users/manas/Hermes/icons/ai_avatar.png")
        ai_image.setPixmap(pixmap.scaledToHeight(500, Qt.SmoothTransformation))
        ai_image.setAlignment(Qt.AlignCenter)

        tagline = QLabel("Your Smart Email Assistant\nThat Actually Understands You")
        tagline.setStyleSheet("font-size: 20px; color: #cccccc; font-weight: 500; ")
        tagline.setAlignment(Qt.AlignCenter)
        tagline.setWordWrap(True)

        left_layout.addWidget(branding)
        left_layout.addWidget(ai_image)
        left_layout.addWidget(tagline)
        left_layout.addStretch()

        # RIGHT SIDE
        right_panel = QWidget()
        right_layout = QVBoxLayout(right_panel)
        right_layout.setContentsMargins(80, 60, 80, 60)
        right_layout.setSpacing(20)

        login_title = QLabel("Log in to Hermys")
        login_title.setStyleSheet("font-size: 28px; font-weight: bold; color: white;")
        login_title.setAlignment(Qt.AlignLeft)

        self.email_input = QLineEdit()
        self.email_input.setPlaceholderText("Email ID")

        self.pass_input = QLineEdit()
        self.pass_input.setPlaceholderText("App Password")
        self.pass_input.setEchoMode(QLineEdit.Password)

        self.api_input = QLineEdit()
        self.api_input.setPlaceholderText("Gemini API Key")
        self.api_input.setEchoMode(QLineEdit.Password)

        self.info_icon_button = QPushButton("i")
        self.info_icon_button.setFixedSize(24, 24)
        self.info_icon_button.setCursor(Qt.PointingHandCursor)
        self.info_icon_button.setStyleSheet("""
            QPushButton {
                background-color: #7f5af0;
                color: white;
                border-radius: 12px;
                font-weight: bold;
                font-size: 12px;
            }
            QPushButton:hover {
                background-color: #6c4bd0;
            }
        """)
        self.info_icon_button.clicked.connect(self.toggle_info)
        self.info_label = QLabel("How to get App Password & API Key")
        self.info_label.setStyleSheet("""
            QLabel {
                color: white;
                font-size: 16px;
                padding-left: 6px;
            }
        """)
        self.info_toggle_widget = QWidget()
        info_layout = QHBoxLayout()
        info_layout.setContentsMargins(0, 0, 0, 0)
        info_layout.setSpacing(4)
        info_layout.addWidget(self.info_icon_button)
        info_layout.addWidget(self.info_label)
        self.info_toggle_widget.setLayout(info_layout)
        self.info_text = QTextEdit()
        self.info_text.setVisible(False)
        self.info_text.setReadOnly(True)
        self.info_text.setStyleSheet("""
            QTextEdit {
            background-color: #7f5af0;
            color: white;
            font-size: 12.5px;
            font-family: 'Segoe UI', sans-serif;
            border: none;
            border-radius: 10px;
            padding: 12px;
            }
        """)
        self.info_text.setText(
            "To get your App Password:\n"
            "1. Go to your email provider’s settings\n"
            "2. Enable 2FA\n"
            "3. Generate an app-specific password\n\n"
            "To get your Gemini API Key:\n"
            "1. Visit: https://aistudio.google.com/app/apikey\n"
            "2. Generate and copy your API Key"
        )

        self.login_button = QPushButton("Login")
        self.login_button.clicked.connect(self.login)

        right_layout.addWidget(login_title)
        right_layout.addWidget(self.email_input)
        right_layout.addWidget(self.pass_input)
        right_layout.addWidget(self.api_input)
        right_layout.addWidget(self.info_toggle_widget)
        right_layout.addWidget(self.info_text)
        right_layout.addStretch()
        right_layout.addWidget(self.login_button)

        main_layout.addWidget(left_panel)
        main_layout.addWidget(right_panel)

    def toggle_info(self):
        self.info_text.setVisible(not self.info_text.isVisible())

    def animate_button(self, button):
        animation = QPropertyAnimation(button, b"geometry")
        geometry = button.geometry()
        animation.setDuration(150)
        animation.setStartValue(geometry.adjusted(-4, -4, 4, 4))
        animation.setEndValue(geometry)
        animation.start()
        button.animation = animation

    def login(self):
        self.animate_button(self.login_button)
        email = self.email_input.text().strip()
        password = self.pass_input.text().strip()
        api_key = self.api_input.text().strip()
        if email and password and api_key:
            self.hide()
            self.dashboard = email_sender1.EmailDashboard(email, password, api_key)
            self.dashboard.show()
        else:
            self.info_text.setVisible(True)
            self.info_text.setText("⚠ Please fill in all fields.")

    def stylesheet(self):
        return """
        QWidget {
            background-color: #0d0c1d;
            font-family: 'Segoe UI', sans-serif;
            font-size: 15px;
        }

        QLineEdit, QTextEdit {
            background-color: rgba(255, 255, 255, 0.06);
            border: 1px solid rgba(255, 255, 255, 0.12);
            border-radius: 14px;
            padding: 14px;
            color: white;
        }

        QTextEdit {
            min-height: 90px;
        }

        QPushButton {
            background-color: rgba(138, 110, 255, 0.25);
            color: white;
            font-weight: bold;
            border: 1px solid #8a6eff;
            border-radius: 14px;
            padding: 14px;
        }

        QPushButton:hover {
            background-color: rgba(122, 93, 252, 0.35);
        }

        QPushButton:pressed {
            background-color: rgba(99, 68, 245, 0.5);
        }

        QLabel {
            color: white;
        }
        """


if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = LoginWindow()
    win.show()
    sys.exit(app.exec_())