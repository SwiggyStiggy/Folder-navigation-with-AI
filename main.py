import sys
import openai
import speech_recognition as sr
import os
import dotenv
from PyQt6.QtWidgets import QApplication, QWidget, QLabel, QPushButton, QVBoxLayout
from PyQt6.QtCore import Qt, QThread, pyqtSignal

dotenv.load_dotenv()
openai.api_key = os.getenv('OPENAI_API_KEY')

class VoiceNavigatorGUI(QWidget):
    def __init__(self):
        super().__init__()
        self.init_ui()
        self.speech_recognizer = SpeechRecognizer(self)
        self.speech_recognizer.command_recognized.connect(self.handle_recognized_command)

    def init_ui(self):
        self.setWindowTitle("Voice Navigator")
        self.setGeometry(100, 100, 800, 600)

        # Status label
        self.status_label = QLabel("Resting", self)
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)

        # Toggle button
        self.toggle_button = QPushButton("Start Listening", self)
        self.toggle_button.clicked.connect(self.toggle_listening)

        # Recognized speech display
        self.recognized_speech_label = QLabel("", self)
        self.recognized_speech_label.setAlignment(Qt.AlignmentFlag.AlignCenter)

        # Layout
        layout = QVBoxLayout()
        layout.addWidget(self.status_label)
        layout.addWidget(self.toggle_button)
        layout.addWidget(self.recognized_speech_label)
        self.setLayout(layout)

        self.listening = False  # State variable
        self.speech_thread = None  # Speech thread

    def toggle_listening(self):
        if self.listening:
            self.status_label.setText("Resting")
            self.toggle_button.setText("Start Listening")
            self.speech_recognizer.stop_listening()
        else:
            self.status_label.setText("Listening...")
            self.toggle_button.setText("Stop Listening")
            self.speech_recognizer.start_listening()

        self.listening = not self.listening

    def handle_recognized_command(self, command):
        self.recognized_speech_label.setText(command)  # Display recognized command
        self.status_label.setText("Processing...")

        # Send the recognized command to OpenAI for interpretation
        response = self.process_command_with_openai(command)

        # Handle the OpenAI response and navigate file system
        self.handle_file_navigation(response)

    def process_command_with_openai(self, command):
        # Send the recognized command to OpenAI for interpretation
        response = openai.ChatCompletion.create(
            model="gpt-3.5-turbo",  
            messages=[
                {"role": "system", "content": "You are a helpful assistant that interprets folder navigation commands."},
                {"role": "user", "content": command}  # The user's speech-to-text command
            ],
            max_tokens=1000
        )

        # Extract and return the interpretation from the response
        return response['choices'][0]['message']['content'].strip()


    def handle_file_navigation(self, response):
        # Example: Let's say the AI interpreted the command as "Open folder named Davit in drive F"
        if "open folder" in response.lower():
            folder_name = response.split("named")[-1].strip()  # Extract folder name
            self.open_folder(folder_name)
        elif "backtrack" in response.lower():
            self.backtrack_folder()
        else:
            self.status_label.setText("Command not understood")

    def open_folder(self, folder_name):
        # Example logic to open folder in file explorer
        folder_path = f"F:\\{folder_name}"  # Modify this based on command parsing
        if os.path.exists(folder_path):
            os.startfile(folder_path)
            self.status_label.setText("Location found!")
        else:
            self.status_label.setText("Location not found!")

    def backtrack_folder(self):
        # Logic to backtrack (go up one directory)
        # This can be adjusted based on your specific requirements
        current_directory = os.getcwd()
        parent_directory = os.path.dirname(current_directory)
        os.chdir(parent_directory)
        self.status_label.setText("Backtracked to previous folder")

class SpeechRecognizer(QThread):
    command_recognized = pyqtSignal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.recognizer = sr.Recognizer()
        self.microphone = sr.Microphone()

    def start_listening(self):
        self.start()

    def stop_listening(self):
        self.quit()

    def run(self):
        with self.microphone as source:
            self.recognizer.adjust_for_ambient_noise(source)
            print("Listening...")
            audio = self.recognizer.listen(source)
        
        try:
            print("Recognizing...")
            recognized_command = self.recognizer.recognize_google(audio)
            print(f"Command recognized: {recognized_command}")
            self.command_recognized.emit(recognized_command)
        except sr.UnknownValueError:
            print("Sorry, I couldn't understand the audio")
        except sr.RequestError:
            print("Could not request results from Google Speech Recognition service")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = VoiceNavigatorGUI()
    window.show()
    sys.exit(app.exec())
