import sys
import os
import openai
import dotenv
import subprocess
import speech_recognition as sr
from PyQt6.QtWidgets import QApplication, QWidget, QLabel, QPushButton, QVBoxLayout, QHBoxLayout
from PyQt6.QtCore import Qt, QThread, pyqtSignal

dotenv.load_dotenv()
openai.api_key = os.getenv('OPENAI_API_KEY')

# -----------------------------
# Speech Recognition Thread
# -----------------------------
class SpeechRecognizerThread(QThread):
    command_recognized = pyqtSignal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self._running = True
        self.recognizer = sr.Recognizer()
        self.microphone = sr.Microphone()

    def run(self):
        with self.microphone as source:
            # Adjust for ambient noise
            self.recognizer.adjust_for_ambient_noise(source)
            while self._running:
                try:
                    print("Listening...")
                    audio = self.recognizer.listen(source, timeout=5)
                    command = self.recognizer.recognize_google(audio)
                    print(f"Command recognized: {command}")
                    self.command_recognized.emit(command)
                except sr.WaitTimeoutError:
                    # No speech within the timeout, continue listening
                    continue
                except sr.UnknownValueError:
                    print("Could not understand audio")
                except sr.RequestError:
                    print("Speech recognition service error")

    def stop(self):
        self._running = False

# -----------------------------
# Main Application Window (Window 1)
# -----------------------------
class VoiceNavigatorGUI(QWidget):
    def __init__(self):
        super().__init__()
        self.init_ui()

        # Initialize the second window (file explorer) once on startup
        if not hasattr(self, 'pc_opened'):
            self.open_this_pc()
            self.pc_opened = True
        self.current_path = None
        self.speech_thread = None

    def init_ui(self):
        pass

    def open_this_pc(self):
        try:
            subprocess.run(["explorer.exe", "/e,", ""])
        except FileNotFoundError as e:
            print(f"Error opening 'This PC': {e}")

    def init_ui(self):
        self.setWindowTitle("AI File Navigator")
        self.setGeometry(100, 100, 800, 800)  # Overall window: 800x800 pixels

        # --- Left side: State Display and Dynamic Button ---
        # State Display: shows status messages
        self.state_label = QLabel("Resting...", self)
        self.state_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.state_label.setFixedSize(200, 200)  # 2:2 ratio (200x200 pixels)

        # Dynamic Button: toggles between "Listen" and "Rest"
        self.toggle_button = QPushButton("Listen", self)
        self.toggle_button.setFixedSize(200, 200)  # 2:2 ratio
        self.toggle_button.clicked.connect(self.toggle_listening)

        left_layout = QVBoxLayout()
        left_layout.addWidget(self.state_label)
        left_layout.addWidget(self.toggle_button)

        # --- Right side: Recognized Speech Screen Display (RSSD) ---
        self.rssd_label = QLabel("Your command will be displayed here...", self)
        self.rssd_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.rssd_label.setFixedSize(400, 400)  # 4:4 ratio

        # --- Main Layout ---
        main_layout = QHBoxLayout()
        main_layout.addLayout(left_layout)
        main_layout.addWidget(self.rssd_label)
        self.setLayout(main_layout)

        self.listening = False  # Application state

    def toggle_listening(self):
        if self.listening:
            # Stop listening
            self.state_label.setText("Resting...")
            self.toggle_button.setText("Listen")
            self.stop_listening()
        else:
            # Start listening
            self.state_label.setText("Listening...")
            self.toggle_button.setText("Rest")
            self.start_listening()
        self.listening = not self.listening

    def start_listening(self):
        self.speech_thread = SpeechRecognizerThread()
        self.speech_thread.command_recognized.connect(self.handle_recognized_command)
        self.speech_thread.start()

    def stop_listening(self):
        if self.speech_thread:
            self.speech_thread.stop()
            self.speech_thread.quit()
            self.speech_thread.wait()

    def handle_recognized_command(self, command):
        # Display the recognized command in the RSSD
        self.rssd_label.setText(command)
        self.state_label.setText("Processing...")
        # Process the command with OpenAI to get a structured response
        structured_command = self.process_command_with_openai(command)
        print("Structured command:", structured_command)
        # Use the structured command to navigate the file explorer
        self.handle_file_navigation(structured_command)

    def process_command_with_openai(self, command):
        """
        Uses OpenAI's Chat API to parse the voice command into a structured command.
        The prompt instructs the AI to output one of:
          - OPEN_DRIVE <drive_letter>
          - OPEN_FOLDER <folder_name>
          - BACKTRACK
          - INVALID
        """
        prompt_instructions = (
            "You are a helpful assistant that converts human file navigation commands into a structured command. "
            "For a given command, output exactly one of the following responses:\n"
            "1. OPEN_DRIVE <drive_letter>    (for commands like 'open drive D')\n"
            "2. OPEN_FOLDER <folder_name>    (for commands like 'open folder named Downloads')\n"
            "3. BACKTRACK                    (for commands like 'backtrack' or 'go up one level')\n"
            "4. INVALID                      (if the command cannot be parsed)\n"
            "Do not include any extra text."
        )

        # Call the ChatCompletion API with a system message and the user's command.
        try:
            response = openai.ChatCompletion.create(
                model="gpt-3.5-turbo",
                messages=[
                    {"role": "system", "content": prompt_instructions},
                    {"role": "user", "content": command}
                ],
                max_tokens=1000
            )
            structured_response = response['choices'][0]['message']['content'].strip()
            return structured_response
        except Exception as e:
            print("OpenAI API error:", e)
            return "INVALID"

    def handle_file_navigation(self, structured_command):
        """
        Parses the structured command and updates the file explorer accordingly.
        """
        if structured_command.startswith("OPEN_DRIVE"):
            parts = structured_command.split()
            if len(parts) >= 2:
                drive_letter = parts[1]
                # Ensure drive letter ends with colon and backslash
                path = drive_letter if drive_letter.endswith(":\\") else drive_letter + ":\\"
                self.current_path = path
                self.open_path(path)
                self.state_label.setText("Location found!")
            else:
                self.state_label.setText("Invalid drive command.")

        elif structured_command.startswith("OPEN_FOLDER"):
            parts = structured_command.split(maxsplit=1)
            if len(parts) == 2:
                folder_name = parts[1]
                if self.current_path is None:
                    # If no drive/folder is currently open, show an error.
                    self.state_label.setText("No drive selected. Please open a drive first.")
                else:
                    path = os.path.join(self.current_path, folder_name)
                    if os.path.isdir(path):
                        self.current_path = path
                        self.open_path(path)
                        self.state_label.setText("Location found!")
                    else:
                        self.state_label.setText("Location not found!")
            else:
                self.state_label.setText("Invalid folder command.")

        elif structured_command == "BACKTRACK":
            if self.current_path is None:
                # Already at "This PC"
                self.state_label.setText("Already at 'This PC' view.")
            else:
                parent = os.path.dirname(self.current_path.rstrip("\\"))
                # If parent becomes empty or the same, go to "This PC"
                if not parent or parent == self.current_path:
                    self.current_path = None
                    self.open_this_pc()
                    self.state_label.setText("Returned to 'This PC'.")
                else:
                    self.current_path = parent
                    self.open_path(parent)
                    self.state_label.setText("Location found!")
        else:
            self.state_label.setText("Command not understood.")

    def open_path(self, path):
        """
        Opens the given path in the default file explorer.
        """
        try:
            os.startfile(path)
        except Exception as e:
            print("Error opening path:", e)
            self.state_label.setText("Error opening location.")

# -----------------------------
# Main entry point
# -----------------------------
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = VoiceNavigatorGUI()
    window.show()
    sys.exit(app.exec())
