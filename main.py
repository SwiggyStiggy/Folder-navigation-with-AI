import sys
import os
import openai
import dotenv
import subprocess
import win32com.client
import speech_recognition as sr
from PyQt6.QtWidgets import QApplication, QWidget, QLabel, QPushButton, QVBoxLayout, QHBoxLayout
from PyQt6.QtCore import Qt, QThread, pyqtSignal

dotenv.load_dotenv()
openai.api_key = os.getenv('OPENAI_API_KEY')

try:
    import win32com.client
except ImportError:
    print("Please install pywin32 package: pip install pywin32")

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

        self.recognizer.energy_threshold = 300

    def run(self):
        with self.microphone as source:
            self.recognizer.adjust_for_ambient_noise(source)
            while self._running:
                try:
                    print("Listening...")
                    audio = self.recognizer.listen(source, timeout=5)
                    command = self.recognizer.recognize_google(audio)
                    print(f"Command recognized: {command}")
                    self.command_recognized.emit(command)
                except sr.WaitTimeoutError:
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

        if not hasattr(self, 'pc_opened'):
            self.open_this_pc()
            self.pc_opened = True
        self.current_path = None
        self.speech_thread = None

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

        self.listening = False

    def open_this_pc(self):
        try:
            subprocess.run(["explorer.exe", "/e,", ""])
        except FileNotFoundError as e:
            print(f"Error opening 'This PC': {e}")

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
        self.rssd_label.setText(command)
        self.state_label.setText("Processing...")
        structured_command = self.process_command_with_openai(command)
        print("Structured command:", structured_command)
        if structured_command.startswith("RUN_EXECUTABLE"):
            self.execute_command(structured_command)
        elif structured_command.startswith("OPEN_FILE"):  # NEW: Handle OPEN_FILE command
            parts = structured_command.split(maxsplit=1)
            if len(parts) == 2:
                filename = parts[1]
                self.open_file(filename)
            else:
                self.state_label.setText("Invalid file command.")
        else:
            self.handle_file_navigation(structured_command)

    def process_command_with_openai(self, command):
        prompt_instructions = (
            "You are an AI that converts human file navigation commands into structured commands. "
            "For a given input, return only ONE of the following structured commands:\n\n"
            "1. OPEN_DRIVE <drive_letter>    (e.g., 'Open D drive' → OPEN_DRIVE D)\n"
            "2. OPEN_FOLDER <folder_name>    (e.g., 'Go to Documents' → OPEN_FOLDER Documents)\n"
            "3. RUN_EXECUTABLE <filename>    (e.g., 'Run app dot exe' → RUN_EXECUTABLE app.exe)\n"
            "4. OPEN_PATH <full_path>        (e.g., 'Open C:/Users/Admin/Desktop' → OPEN_PATH C:/Users/Admin/Desktop)\n"
            "5. SEARCH_FILE <filename>       (e.g., 'Find my resume' → SEARCH_FILE resume)\n"
            "6. BACKTRACK                    (e.g., 'Go back one step' → BACKTRACK)\n"
            "7. INVALID                      (if the command cannot be understood)\n"
            "8. OPEN_FILE <filename>         (e.g., 'Open report dot pdf' → OPEN_FILE report.pdf)\n"  # NEW: Added option for opening any file type\n"
            "\n"
            "Note: if a user says underscore in the sentence they mean → _\n"
            "Note: if a user says dot in the sentence they mean → .\n"
            "Note: If a user provides a folder name that contains spaces, KEEP THE SPACES INTACT. Do not replace spaces with underscores. Only return the folder name as the user said it.\n"
            "IMPORTANT: Always return only the structured command. Do not add any explanations."
        )

        try:
            response = openai.ChatCompletion.create(
                model="gpt-3.5-turbo",
                messages=[
                    {"role": "system", "content": prompt_instructions},
                    {"role": "user", "content": command}
                ],
                max_tokens=50
            )
            structured_response = response['choices'][0]['message']['content'].strip()
            return structured_response
        except Exception as e:
            print("OpenAI API error:", e)
            return "INVALID"

    def execute_command(self, command):
        if command.startswith("RUN_EXECUTABLE"):
            _, exe_name = command.split(" ", 1)
            if not exe_name.endswith(".exe"):
                exe_name += ".exe"
            folder_path = self.current_path if self.current_path else os.getcwd()
            exe_path = os.path.join(folder_path, exe_name)
            if os.path.isfile(exe_path):
                try:
                    subprocess.Popen(exe_path, shell=True)
                    print(f"Launching {exe_name} from {folder_path}...")
                except Exception as e:
                    print(f"Error launching {exe_name}: {e}")
            else:
                files_in_folder = os.listdir(folder_path)
                matching_files = [f for f in files_in_folder if f.lower() == exe_name.lower()]
                if matching_files:
                    exe_path = os.path.join(folder_path, matching_files[0])
                    subprocess.Popen(exe_path, shell=True)
                    print(f"Launching {matching_files[0]} from {folder_path}...")
                else:
                    print(f"Executable {exe_name} not found in {folder_path}.")

    def handle_file_navigation(self, structured_command):
        if structured_command.startswith("OPEN_DRIVE"):
            parts = structured_command.split()
            if len(parts) >= 2:
                drive_letter = parts[1]
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
                self.state_label.setText("Already at 'This PC' view.")
            else:
                parent = os.path.dirname(self.current_path.rstrip("\\"))                              
                if not parent or parent == self.current_path:                   
                    self.current_path = None                 
                    self.open_this_pc()                  
                    self.state_label.setText("Returned to 'This PC'.")
                else:\
                    self.current_path = parent
                self.open_path(parent)        
                self.state_label.setText("Location found!")
        elif structured_command.startswith("OPEN_PATH"):
            self.current_path = structured_command.split(" ", 1)[1]
            self.open_path(self.current_path)
        else:
            self.state_label.setText("Command not understood.")
            
    # NEW: Added method to open any file with its default application
    def open_file(self, filename):
        if not self.current_path:
            self.state_label.setText("No directory selected. Please navigate to a folder first.")
            return
        
        file_path = os.path.join(self.current_path, filename)
        
        if not os.path.exists(file_path):
            self.state_label.setText(f"Error: '{filename}' not found in {self.current_path}")
            return
        
        try:
            if os.name == 'nt':
                os.startfile(file_path)  # Windows
            elif os.uname().sysname == 'Darwin':
                subprocess.call(['open', file_path])  # macOS
            else:
                subprocess.call(['xdg-open', file_path])  # Linux
            print(f"Opened: {filename}")
            self.state_label.setText(f"Opened {filename} successfully!")
        except Exception as e:
            print(f"Error opening file: {e}")
            self.state_label.setText("Failed to open file.")
            
    def open_path(self, path):
        try:
            shell = win32com.client.Dispatch("Shell.Application")
            windows = shell.Windows()
            found = False
            for window in windows:
                if window.LocationName != "" and window.Name.lower() in ["file explorer", "explorer"]:
                    window.Navigate(path)
                    found = True
                    print(f"Navigating existing explorer window to: {path}")
                    break
            if not found:
                subprocess.Popen(["explorer", path])
                print(f"Opening new explorer window to: {path}")
        except Exception as e:
            print("Error opening path:", e)
            self.state_label.setText("Error opening location.")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = VoiceNavigatorGUI()
    window.show()
    sys.exit(app.exec())
