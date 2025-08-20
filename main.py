import sys
import os
import pyperclip
import keyboard
import mouse
from PyQt5 import QtWidgets, QtCore, QtGui
import openai
import time
import win32gui
import win32con
import win32process
import psutil
import threading
import json
import http.client
import logging
import shutil
import re

APP_PID = os.getpid()
DOUBLE_TAP_MAX_DELAY = 0.35  # seconds between taps
last_shift_time = 0
DEBUG = bool(os.environ.get('REPHRASER_DEBUG'))

def is_own_window_focused():
    try:
        hwnd = win32gui.GetForegroundWindow()
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        return pid == APP_PID
    except Exception as e:
        print(f"[is_own_window_focused] Exception: {e}")
        return False

def is_supported_app_focused():
    try:
        hwnd = win32gui.GetForegroundWindow()
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        proc = psutil.Process(pid)
        exe = proc.name().lower()
        return exe in settings.get('supported_apps', [])
    except Exception as e:
        debug_print('[DEBUG] is_supported_app_focused error:', e)
        return False

if DEBUG:
    http.client.HTTPConnection.debuglevel = 0
    logging.basicConfig()
    logging.getLogger().setLevel(logging.DEBUG)
    requests_log = logging.getLogger("urllib3")
    requests_log.setLevel(logging.DEBUG)
    requests_log.propagate = True

def debug_print(*args, **kwargs):
    if DEBUG:
        print(*args, **kwargs)

SETTINGS_FILE = './assets/settings.json'
DEFAULT_SETTINGS = {
    'api_key': '',
    'api_url': 'https://api.openai.com/v1',
    'model': 'gpt-3.5-turbo',
    'prompt': 'You are a helpful assistant that rephrases text in a clear and concise way.',
    'supported_apps': ['outlook.exe', 'notepad.exe', 'chrome.exe']
}
settings = {}

def load_settings():
    global settings
    loaded = DEFAULT_SETTINGS.copy()
    if os.path.exists(SETTINGS_FILE):
        try:
            with open(SETTINGS_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
                loaded.update(data)
        except Exception as e:
            print(f"[load_settings] Error: {e}")
    settings.clear()
    settings.update(loaded)
    openai.api_key = settings['api_key']
    openai.base_url = settings['api_url']

def save_settings():
    try:
        with open(SETTINGS_FILE, 'w', encoding='utf-8') as f:
            json.dump(settings, f, indent=2)
    except Exception as e:
        print(f"[save_settings] Error: {e}")
    load_settings()

load_settings()

def get_icon_path():
    # Look for icon files in the ./assets directory first
    base_dir = os.path.dirname(sys.executable if getattr(sys, 'frozen', False) else os.path.abspath(sys.argv[0]))
    assets_dir = os.path.join(base_dir, 'assets')
    ico_path = os.path.join(assets_dir, 'icon.ico')
    png_path = os.path.join(assets_dir, 'icon.png')
    if os.path.exists(ico_path):
        return ico_path
    if os.path.exists(png_path):
        return png_path
    # fallback to base_dir if not found in assets
    ico_path_base = os.path.join(base_dir, 'icon.ico')
    png_path_base = os.path.join(base_dir, 'icon.png')
    if os.path.exists(ico_path_base):
        return ico_path_base
    return png_path_base

class FloatingButton(QtWidgets.QWidget):
    overlay_created = QtCore.pyqtSignal(object)

    def __init__(self, selected_text, source_hwnd, parent=None):
        super().__init__(parent)
        self.selected_text = selected_text
        self.source_hwnd = source_hwnd
        self.setWindowFlags(
            QtCore.Qt.FramelessWindowHint |
            QtCore.Qt.WindowStaysOnTopHint |
            QtCore.Qt.Tool
        )
        self.setAttribute(QtCore.Qt.WA_TranslucentBackground)
        self.init_ui()
        QtCore.QTimer.singleShot(5000, self.close)

    def init_ui(self):
        layout = QtWidgets.QVBoxLayout()
        self.button = QtWidgets.QPushButton()
        self.button.setIcon(QtGui.QIcon(get_icon_path()))
        self.button.setIconSize(QtCore.QSize(40, 40))
        self.button.setFixedSize(65, 65)
        self.button.setStyleSheet('border: none; background: transparent;')
        self.button.clicked.connect(self.rephrase_text)
        layout.addWidget(self.button)
        self.setLayout(layout)
        self.setFixedSize(65, 65)
    def rephrase_text(self):
        self.button.setEnabled(False)
        overlay = RephraseOverlay(self.selected_text, self.source_hwnd)
        overlay.show_near_cursor()
        self.overlay_created.emit(overlay)
        self.close()

    def show_near_cursor(self):
        pos = QtGui.QCursor.pos()
        self.move(pos.x() + 20, pos.y())
        self.show()
        debug_print('[DEBUG] FloatingButton shown at', pos.x() + 20, pos.y())

class RephraseWorker(QtCore.QThread):
    result_ready = QtCore.pyqtSignal(str, bool)

    def __init__(self, selected_text):
        super().__init__()
        self.selected_text = selected_text

    def run(self):
        #try:
        debug_print('[DEBUG] api_key and api_url', settings['api_key'], settings['api_url'])
        openai.api_key = settings['api_key']
        openai.base_url = settings['api_url']

        lines = self.selected_text.split('\n')
        reconstructed_lines = list(lines)
        
        lines_to_rephrase_map = {}  # Maps original index to the line content
        for idx, line in enumerate(lines):
            if line.strip() and not (
                line.strip().startswith('#') or
                line.strip().startswith('%') or
                self.is_code_like(line)
            ):
                lines_to_rephrase_map[idx] = line

        if not lines_to_rephrase_map:
            self.result_ready.emit(self.selected_text, False)
            return

        lines_to_send = list(lines_to_rephrase_map.values())
        input_json_str = json.dumps({"lines_to_rephrase": lines_to_send})

        system_prompt = (
            settings['prompt']
            #+ " You will be given a JSON object with a key 'lines_to_rephrase' containing a list of strings. "
            #+ "Your task is to rephrase each string in the list. "
            #+ "You MUST respond with a JSON object that contains a single key, 'rephrased_lines', "
            #+ "which is a list of the rephrased strings. "
            #+ "The returned list must have the exact same number of items as the input list."
            #+ "Most of the time, these lines are all part of the same email or text. "
            + "You will be given a JSON object with a key 'lines_to_rephrase' containing a list of strings. "
            + "These strings are all part of the same email or message and must be understood in that shared context."
            + "Your task is to rephrase each line while preserving the meaning and tone appropriate to the overall message. "
            + "Pay attention to how the lines relate to one another to maintain consistency, flow, and coherence."
            + "You MUST respond with a JSON object containing a single key, 'rephrased_lines', which is a list of the rephrased strings. "
            + "The output list should have the exact same number of items, in the same order, as the input list."
        )

        messages = [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": input_json_str}
        ]
        
        response = openai.chat.completions.create(
            model=settings.get('model', 'gpt-3.5-turbo'),
            messages=messages,
            max_tokens=1024, # Increased max_tokens for JSON overhead
            temperature=0.7,
            timeout=20.0, # Increased timeout for potentially longer processing
            # response_format={"type": "json_object"} # Ideal, but might not be supported by all endpoints
        )
        
        reply_content = response.choices[0].message.content.strip()
        debug_print('[DEBUG] Raw OpenAI response:\n', reply_content)

        # Extract JSON from the reply, which might be wrapped in markdown
        match = re.search(r"\{.*\}", reply_content, re.DOTALL)
        if match:
            json_str = match.group(0)
        else:
            self.result_ready.emit(f"Error: Model did not return valid JSON.\n\n{reply_content}", True)
            return

        try:
            response_data = json.loads(json_str)
            rephrased_lines = response_data.get("rephrased_lines", [])
        except json.JSONDecodeError:
            self.result_ready.emit(f"Error: Failed to decode JSON from model response.\n\n{reply_content}", True)
            return

        if not isinstance(rephrased_lines, list) or len(rephrased_lines) != len(lines_to_send):
            error_msg = (
                f"Error: Rephrased data is invalid or has a mismatched number of lines "
                f"({len(rephrased_lines)}) than expected ({len(lines_to_send)})."
            )
            self.result_ready.emit(f"{error_msg}\n\n{reply_content}", True)
            return

        # Reconstruct the text
        rephrased_lines_iter = iter(rephrased_lines)
        debug_print('[DEBUG] rephrased_lines_iter: ', rephrased_lines_iter)
        for index in lines_to_rephrase_map.keys():
            try:
                rephrased_line = next(rephrased_lines_iter)
                debug_print('[DEBUG] reconstructed_lines[index]: ', reconstructed_lines[index])
                debug_print('[DEBUG] rephrased_line: ', rephrased_line)
                reconstructed_lines[index] = rephrased_line
            except StopIteration:
                debug_print(f"[DEBUG] StopIteration at index {index}. Mismatch between lines to rephrase and rephrased lines.")
                break

        final_text = '\n'.join(reconstructed_lines)
        self.result_ready.emit(final_text, False)

        #except Exception as e:
        #debug_print('[DEBUG] error', e)
        #self.result_ready.emit(f"Error: {str(e)}", True)

    def is_code_like(self, line):
        stripped = line.strip()
        return (
            stripped.endswith(':') or
            (stripped and (stripped.startswith('def ') or stripped.startswith('class '))) or
            ('=' in line and not line.strip().startswith('//')) or
            ('import ' in line) or
            ('print(' in line) or
            (stripped.startswith('for ') or stripped.startswith('while ') or stripped.startswith('if '))
        )

class RephraseOverlay(QtWidgets.QWidget):
    def __init__(self, selected_text, source_hwnd, parent=None):
        super().__init__(parent)
        self.selected_text = selected_text
        self.setWindowFlags(QtCore.Qt.FramelessWindowHint | QtCore.Qt.WindowStaysOnTopHint | QtCore.Qt.Tool)
        self.setAttribute(QtCore.Qt.WA_TranslucentBackground)
        self.prev_hwnd = source_hwnd
        self.init_ui()
        self.get_rephrased_text()
        # Fade effect
        self.opacity_effect = QtWidgets.QGraphicsOpacityEffect(self)
        self.setGraphicsEffect(self.opacity_effect)
        self.fade_anim = QtCore.QPropertyAnimation(self.opacity_effect, b"opacity")
        self.fade_anim.setDuration(300)
        self.fade_anim.setStartValue(0)
        self.fade_anim.setEndValue(1)
        self.fade_anim.start()
        # Auto-close timer (10 seconds, always starts on show)
        self.auto_close_timer = QtCore.QTimer(self)
        self.auto_close_timer.setSingleShot(True)
        self.auto_close_timer.timeout.connect(self.on_auto_close_timeout)
        self.setMouseTracking(True)
        self.timer_expired = False
        self.auto_close_timer.start(10000)

    def on_auto_close_timeout(self):
        self.timer_expired = True
        if not self.underMouse():
            self.fade_and_close()
        # else: wait for leaveEvent to close

    def enterEvent(self, event):
        super().enterEvent(event)

    def leaveEvent(self, event):
        if self.timer_expired:
            self.fade_and_close()
        super().leaveEvent(event)

    def fade_and_close(self):
        self.fade_anim = QtCore.QPropertyAnimation(self.opacity_effect, b"opacity")
        self.fade_anim.setDuration(400)
        self.fade_anim.setStartValue(1)
        self.fade_anim.setEndValue(0)
        self.fade_anim.finished.connect(self.close)
        self.fade_anim.start()

    def showEvent(self, event):
        # Fade in when shown
        self.opacity_effect.setOpacity(0)
        self.fade_anim = QtCore.QPropertyAnimation(self.opacity_effect, b"opacity")
        self.fade_anim.setDuration(300)
        self.fade_anim.setStartValue(0)
        self.fade_anim.setEndValue(1)
        self.fade_anim.start()
        super().showEvent(event)

    def closeEvent(self, event):
        self.auto_close_timer.stop()
        QtCore.QTimer.singleShot(100, self.clear_clipboard)
        super().closeEvent(event)

    def clear_clipboard(self):
        try:
            pyperclip.copy('')
            debug_print('[DEBUG] Clipboard cleared on overlay close.')
        except Exception as e:
            debug_print(f'[DEBUG] Failed to clear clipboard on overlay close: {e}')

    def init_ui(self):
        layout = QtWidgets.QVBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
        self.frame = QtWidgets.QFrame()
        self.frame.setStyleSheet('background: #e0ffe0; border-radius: 16px;')
        frame_layout = QtWidgets.QVBoxLayout(self.frame)
        frame_layout.setContentsMargins(10, 10, 10, 10)
        frame_layout.setSpacing(0)
        close_layout = QtWidgets.QHBoxLayout()
        close_layout.setContentsMargins(0, 0, 0, 0)
        close_layout.setSpacing(0)
        close_layout.addStretch()
        close_btn = QtWidgets.QPushButton('✕')
        close_btn.setFixedSize(32, 32)
        close_btn.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        close_btn.setStyleSheet('border: none; background: rgba(255,255,255,0.01); font-size: 18px; color: #888; padding: 2px; margin: 0px;')
        close_btn.clicked.connect(self.close)
        close_layout.addWidget(close_btn)
        frame_layout.addLayout(close_layout)
        self.text_label = QtWidgets.QLabel()
        self.text_label.setWordWrap(True)
        self.text_label.setAlignment(QtCore.Qt.AlignTop | QtCore.Qt.AlignLeft)
        self.text_label.setStyleSheet("background: transparent; font-size: 14px; padding: 0px;")
        self.text_label.setSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Minimum)
        self.text_label.installEventFilter(self)
        frame_layout.addWidget(self.text_label, alignment=QtCore.Qt.AlignTop)
        self.instruction_label = QtWidgets.QLabel('(Click on the green area to replace the selected text).')
        self.instruction_label.setAlignment(QtCore.Qt.AlignCenter)
        self.instruction_label.setStyleSheet('color: #666; font-size: 11px; padding-top: 8px; background: transparent;')
        frame_layout.addWidget(self.instruction_label, alignment=QtCore.Qt.AlignBottom)
        self.loading_label = QtWidgets.QLabel('Loading...')
        self.loading_label.setAlignment(QtCore.Qt.AlignCenter)
        self.loading_label.setStyleSheet('color: #888; font-size: 16px; background: transparent;')
        frame_layout.addWidget(self.loading_label, alignment=QtCore.Qt.AlignVCenter)
        self.loading_label.hide()
        layout.addWidget(self.frame)
        self.setLayout(layout)
        self.setMinimumSize(200, 60)
        self.setMaximumSize(1200, 800)

    def get_rephrased_text(self):
        self.text_label.hide()
        self.instruction_label.hide()
        self.loading_label.show()
        self.worker = RephraseWorker(self.selected_text)
        self.worker.result_ready.connect(self.on_result_ready)
        self.worker.start()

    def on_result_ready(self, result, is_error):
        debug_print('[DEBUG] on_result_ready called with:', repr(result), 'is_error:', is_error)
        # Always clean tags before display
        if isinstance(result, str):
            result = re.sub(r"\[\[REPHRASE:\s*\d+\]\]\s*", "", result, flags=re.IGNORECASE | re.MULTILINE)
        self.loading_label.hide()
        if is_error:
            debug_print('[DEBUG] Setting error text in label:', repr(result))
            self.text_label.setText(result)
            self.text_label.setStyleSheet("background: #ffe0e0; padding: 8px; border-radius: 16px; font-size: 14px;")
        else:
            debug_print('[DEBUG] Setting normal text in label:', repr(result))
            self.text_label.setText(result)
            self.text_label.setStyleSheet("background: transparent; font-size: 14px;")
        self.text_label.show()
        self.instruction_label.show()
        self.adjust_size_to_text()
        self.show()
        self.raise_()
        self.activateWindow()

    def adjust_size_to_text(self):
        font = self.text_label.font()
        metrics = QtGui.QFontMetrics(font)
        lines = self.text_label.text().splitlines() or ['']
        max_line_width = max((metrics.width(line) for line in lines), default=200)
        width = max(max_line_width + 40, 200)
        content_height = max(metrics.height() * len(lines) + 40, 60)
        self.resize(width, content_height)

    def eventFilter(self, obj, event):
        if obj == self.text_label and event.type() == QtCore.QEvent.MouseButtonPress:
            if hasattr(self, 'auto_close_timer'):
                self.auto_close_timer.stop()
            rephrased = self.text_label.text()
            pyperclip.copy('')
            time.sleep(0.05)
            pyperclip.copy(rephrased)
            debug_print('[DEBUG] Copied rephrased text to clipboard')
            self.hide()
            self.close()
            try:
                if self.prev_hwnd:
                    win32gui.ShowWindow(self.prev_hwnd, win32con.SW_SHOW)
                    win32gui.SetForegroundWindow(self.prev_hwnd)
                    time.sleep(0.1)
                keyboard.press_and_release('ctrl+v')
                debug_print('[DEBUG] Pasted rephrased text.')
                notif = NotificationWindow('Rephrased text has been pasted.')
                notif.show()
                time.sleep(0.1)
                pyperclip.copy('')
            except Exception as e:
                debug_print(f'[DEBUG] Failed to paste: {e}')
                notif = NotificationWindow('Rephrased text copied!<br>Could not paste automatically.')
                notif.show()
            return True
        return super().eventFilter(obj, event)

    def show_near_cursor(self):
        pos = QtGui.QCursor.pos()
        self.adjust_size_to_text()
        self.move(pos.x() + 10, pos.y() + 10)
        self.show()
        self.raise_()
        self.activateWindow()
        debug_print('[DEBUG] RephraseOverlay shown at', pos.x() + 10, pos.y() + 10)

class SelectionListener(QtCore.QObject):
    request_show_rephrase_overlay = QtCore.pyqtSignal(str, int)

    def __init__(self, app):
        super().__init__()
        self.app = app
        self.overlay = None
        self.request_show_rephrase_overlay.connect(self.show_rephrase_overlay)


    def trigger_rephrase(self):
        if not is_supported_app_focused():
            debug_print('[DEBUG] Hotkey triggered, but not a supported app.')
            return

        source_hwnd = win32gui.GetForegroundWindow()
        old_clip = ''
        try:
            old_clip = pyperclip.paste()
        except Exception as e:
            debug_print(f"[DEBUG] Could not get clipboard content: {e}")

        time.sleep(0.05)
        keyboard.press_and_release('ctrl+c')
        time.sleep(0.05)

        text = ''
        try:
            text = pyperclip.paste()
        except Exception as e:
            debug_print(f"[DEBUG] Could not paste: {e}")

        if text and text != old_clip:
            debug_print('[DEBUG] Hotkey pressed, showing rephrase overlay for:', text[:50])
            self.request_show_rephrase_overlay.emit(text, source_hwnd)
        else:
            debug_print('[DEBUG] No selection copied, not showing overlay.')
        
        try:
            pyperclip.copy(old_clip)
        except Exception as e:
            debug_print(f"[DEBUG] Could not restore clipboard: {e}")
        if not is_supported_app_focused():
            debug_print('[DEBUG] Not a supported app, not showing button.')
            return
        if keyboard.is_pressed('ctrl'):
            debug_print('[DEBUG] Ctrl is currently pressed, skipping simulated copy to avoid key state issues.')
            return
        
        source_hwnd = win32gui.GetForegroundWindow()
        old_clip = ''
        try:
            old_clip = pyperclip.paste()
        except Exception as e:
            debug_print(f"[DEBUG] Could not get clipboard content: {e}")

        #unique_marker = f"GRephraser-{time.time()}"
        #try:
        #    pyperclip.copy(unique_marker)
        # except Exception as e:
        #    debug_print(f"[DEBUG] Could not set unique marker to clipboard: {e}")
        #    return

        time.sleep(0.05)
        keyboard.press_and_release('ctrl+c')

        text = ''
        text = pyperclip.paste()
        

    def show_rephrase_overlay(self, text, source_hwnd):
        debug_print('[DEBUG] show_rephrase_overlay called with:', repr(text))
        if self.overlay is not None:
            try:
                self.overlay.close()
            except Exception as e:
                debug_print('[DEBUG] Error closing previous overlay:', e)
            self.overlay = None
        self.overlay = RephraseOverlay(text, source_hwnd)
        self.overlay.show_near_cursor()

def get_startup_shortcut_path():
    startup_dir = os.path.join(os.environ['APPDATA'], r'Microsoft\Windows\Start Menu\Programs\Startup')
    if getattr(sys, 'frozen', False):
        exe_path = sys.executable
        base_dir = os.path.dirname(exe_path)
    else:
        exe_path = os.path.abspath(sys.argv[0])
        base_dir = os.path.dirname(exe_path) 
    ico_path = os.path.join(base_dir, 'icon.ico')
    png_path = os.path.join(base_dir, 'icon.png')
    shortcut_name = 'GRephraser.lnk'
    return os.path.join(startup_dir, shortcut_name), exe_path, ico_path, png_path

def enable_startup():
    shortcut_path, exe_path, ico_path, png_path = get_startup_shortcut_path()
    working_dir = os.path.dirname(exe_path)
    import pythoncom
    from win32com.shell import shell, shellcon # type: ignore
    shell_link = pythoncom.CoCreateInstance(shell.CLSID_ShellLink, None, pythoncom.CLSCTX_INPROC_SERVER, shell.IID_IShellLink)
    shell_link.SetPath(exe_path)
    shell_link.SetWorkingDirectory(working_dir)
    shell_link.SetDescription('GRephraser')
    if os.path.exists(ico_path):
        shell_link.SetIconLocation(ico_path, 0)
    elif os.path.exists(png_path):
        shell_link.SetIconLocation(png_path, 0)
    else:
        shell_link.SetIconLocation(exe_path, 0)
    persist_file = shell_link.QueryInterface(pythoncom.IID_IPersistFile)
    persist_file.Save(shortcut_path, 0)

def disable_startup():
    shortcut_path, _, _, _ = get_startup_shortcut_path()
    if os.path.exists(shortcut_path):
        os.remove(shortcut_path)

def is_startup_enabled():
    shortcut_path, _, _, _ = get_startup_shortcut_path()
    return os.path.exists(shortcut_path)

class ModelFetchWorker(QtCore.QThread):
    models_ready = QtCore.pyqtSignal(list, str)

    def __init__(self, api_key, api_url):
        super().__init__()
        self.api_key = api_key
        self.api_url = api_url

    def run(self):
        try:
            client = openai.OpenAI(api_key=self.api_key, base_url=self.api_url)
            models = client.models.list()
            model_ids = sorted([model.id for model in models.data])
            self.models_ready.emit(model_ids, "")
        except Exception as e:
            self.models_ready.emit([], str(e))

class SettingsWindow(QtWidgets.QMainWindow):
    PREDEFINED_APPS = {
        "Microsoft Outlook": "outlook.exe",
        "Google Chrome": "chrome.exe",
        "Notepad": "notepad.exe",
        "Visual Studio Code": "code.exe",
        "Slack": "slack.exe",
        "Microsoft Word": "winword.exe"
    }

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle('Settings')
        self.setMinimumSize(450, 400)
        self.setWindowIcon(QtGui.QIcon(get_icon_path()))
        self.tabs = QtWidgets.QTabWidget()
        self.general_tab = QtWidgets.QWidget()
        self.parameters_tab = QtWidgets.QWidget()
        self.apps_tab = QtWidgets.QWidget()
        self.tabs.addTab(self.general_tab, 'General')
        self.tabs.addTab(self.parameters_tab, 'Parameters')
        self.tabs.addTab(self.apps_tab, 'Applications')
        self.init_general_tab()
        self.init_parameters_tab()
        self.init_apps_tab()
        btn_box = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Save | QtWidgets.QDialogButtonBox.Cancel)
        btn_box.accepted.connect(self.save_and_close)
        btn_box.rejected.connect(self.close)
        layout = QtWidgets.QVBoxLayout()
        layout.addWidget(self.tabs)
        layout.addWidget(btn_box)
        container = QtWidgets.QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)
        self.load_current_settings()
        self.worker = None

    def init_general_tab(self):
        layout = QtWidgets.QVBoxLayout()
        self.startup_checkbox = QtWidgets.QCheckBox('Start this application automatically at Windows startup')
        self.startup_checkbox.setChecked(is_startup_enabled())
        layout.addWidget(self.startup_checkbox)
        layout.addStretch()
        self.general_tab.setLayout(layout)

    def init_parameters_tab(self):
        layout = QtWidgets.QFormLayout()
        self.api_key_edit = QtWidgets.QLineEdit()
        self.api_key_edit.setEchoMode(QtWidgets.QLineEdit.Password)
        self.api_url_edit = QtWidgets.QLineEdit()
        
        model_layout = QtWidgets.QHBoxLayout()
        self.model_combo = QtWidgets.QComboBox()
        self.model_combo.setMinimumWidth(200)
        self.fetch_models_btn = QtWidgets.QPushButton("Fetch Models")
        self.fetch_models_btn.clicked.connect(self.fetch_models)
        model_layout.addWidget(self.model_combo)
        model_layout.addWidget(self.fetch_models_btn)
        
        self.prompt_edit = QtWidgets.QPlainTextEdit()
        
        layout.addRow('API Key:', self.api_key_edit)
        layout.addRow('API URL:', self.api_url_edit)
        layout.addRow('Model:', model_layout)
        layout.addRow('Prompt:', self.prompt_edit)
        self.parameters_tab.setLayout(layout)

    def fetch_models(self):
        api_key = self.api_key_edit.text().strip()
        api_url = self.api_url_edit.text().strip()
        if not api_key or not api_url:
            QtWidgets.QMessageBox.warning(self, "API Info Missing", "Please enter both an API Key and API URL.")
            return
        
        self.fetch_models_btn.setText("Fetching...")
        self.fetch_models_btn.setEnabled(False)
        
        self.worker = ModelFetchWorker(api_key, api_url)
        self.worker.models_ready.connect(self.on_models_fetched)
        self.worker.start()

    def on_models_fetched(self, models, error_str):
        self.fetch_models_btn.setText("Fetch Models")
        self.fetch_models_btn.setEnabled(True)
        
        if error_str:
            QtWidgets.QMessageBox.critical(self, "Error", f"Failed to fetch models:\n{error_str}")
            return
        
        self.model_combo.clear()
        self.model_combo.addItems(models)
        
        current_model = settings.get('model')
        if current_model in models:
            self.model_combo.setCurrentText(current_model)
        
        QtWidgets.QMessageBox.information(self, "Success", f"Successfully fetched {len(models)} models.")

    def init_apps_tab(self):
        layout = QtWidgets.QVBoxLayout()
        label = QtWidgets.QLabel("Enable GRephraser for these applications:")
        layout.addWidget(label)
        self.app_checkboxes = {}
        for display_name, exe_name in self.PREDEFINED_APPS.items():
            checkbox = QtWidgets.QCheckBox(display_name)
            self.app_checkboxes[exe_name] = checkbox
            layout.addWidget(checkbox)
        layout.addStretch()
        self.apps_tab.setLayout(layout)

    def load_current_settings(self):
        self.api_key_edit.setText(settings.get('api_key', ''))
        self.api_url_edit.setText(settings.get('api_url', ''))
        
        # Add the current model to the combo box, even if it's not in the fetched list yet
        current_model = settings.get('model', 'gpt-3.5-turbo')
        self.model_combo.clear()
        self.model_combo.addItem(current_model)
        self.model_combo.setCurrentText(current_model)

        self.prompt_edit.setPlainText(settings.get('prompt', ''))
        
        enabled_apps = settings.get('supported_apps', [])
        for exe_name, checkbox in self.app_checkboxes.items():
            checkbox.setChecked(exe_name in enabled_apps)

    def save_and_close(self):
        settings['api_key'] = self.api_key_edit.text().strip()
        settings['api_url'] = self.api_url_edit.text().strip()
        settings['model'] = self.model_combo.currentText()
        settings['prompt'] = self.prompt_edit.toPlainText().strip()
        
        enabled_apps = []
        for exe_name, checkbox in self.app_checkboxes.items():
            if checkbox.isChecked():
                enabled_apps.append(exe_name)
        settings['supported_apps'] = enabled_apps
        
        save_settings()
        
        if self.startup_checkbox.isChecked():
            try:
                enable_startup()
            except Exception as e:
                debug_print('[DEBUG] Failed to enable startup:', e)
        else:
            try:
                disable_startup()
            except Exception as e:
                debug_print('[DEBUG] Failed to disable startup:', e)
        self.close()

class SystemTrayIcon(QtWidgets.QSystemTrayIcon):
    def __init__(self, app, parent=None):
        icon = QtGui.QIcon(get_icon_path())
        super().__init__(icon, parent)
        self.app = app
        self.setToolTip('ChatGPT Rephraser')
        menu = QtWidgets.QMenu(parent)
        settings_action = menu.addAction('Settings')
        settings_action.triggered.connect(self.show_settings)
        exit_action = menu.addAction('Exit')
        exit_action.triggered.connect(self.exit_app)
        self.setContextMenu(menu)
        self.activated.connect(self.on_activated)
        self.settings_window = None
        self.show()

    def show_settings(self):
        if self.settings_window is None or not self.settings_window.isVisible():
            self.settings_window = SettingsWindow()
            self.settings_window.show()
        else:
            self.settings_window.raise_()
            self.settings_window.activateWindow()

    def exit_app(self):
        mouse.unhook_all()
        keyboard.unhook_all()
        QtCore.QCoreApplication.quit()

    def on_activated(self, reason):
        if reason == QtWidgets.QSystemTrayIcon.Trigger:
            self.contextMenu().popup(QtGui.QCursor.pos())

class NotificationWindow(QtWidgets.QWidget):
    def __init__(self, message, duration=2000, parent=None):
        super().__init__(parent)
        self.setWindowFlags(QtCore.Qt.FramelessWindowHint | QtCore.Qt.WindowStaysOnTopHint | QtCore.Qt.Tool)
        self.setAttribute(QtCore.Qt.WA_TranslucentBackground)
        layout = QtWidgets.QVBoxLayout()
        label = QtWidgets.QLabel(message)
        label.setStyleSheet("background: #ffffe0; padding: 12px; border-radius: 8px; font-size: 14px; color: #333;")
        layout.addWidget(label)
        self.setLayout(layout)
        self.adjustSize()
        screen = QtWidgets.QApplication.primaryScreen().availableGeometry()
        self.move(screen.center() - self.rect().center())
        QtCore.QTimer.singleShot(duration, self.close)

class DoubleCtrlListener:
    def __init__(self, callback):
        self.callback = callback
        self.last_ctrl_press_time = 0
        keyboard.on_press_key("ctrl", self.on_ctrl_press, suppress=False)

    def on_ctrl_press(self, key_event):
        current_time = time.time()
        if current_time - self.last_ctrl_press_time < 0.3:
            self.callback()
            # Reset the timer to prevent immediate re-triggering
            self.last_ctrl_press_time = 0
        else:
            self.last_ctrl_press_time = current_time

class GlobalPasteHotkey(QtCore.QObject):
    def __init__(self, parent=None):
        super().__init__(parent)
        #keyboard.add_hotkey('ctrl+shift+v', self.paste_clipboard)

    def paste_clipboard(self):
        debug_print('[DEBUG] Sending Ctrl+V')
        keyboard.press_and_release('ctrl+v')

def main():
    global hidden_main
    app = QtWidgets.QApplication(sys.argv)
    app.setWindowIcon(QtGui.QIcon(get_icon_path()))
    app.setQuitOnLastWindowClosed(False)
    hidden_main = QtWidgets.QMainWindow()
    hidden_main.setWindowIcon(QtGui.QIcon(get_icon_path()))
    hidden_main.setWindowTitle('GRephraser')
    hidden_main.setGeometry(-10000, -10000, 100, 100)
    hidden_main.show()
    hidden_main.hide()
    tray = SystemTrayIcon(app)
    listener = SelectionListener(app)
    paste_hotkey = GlobalPasteHotkey()
    
    # Set up the global hotkey for rephrasing
    double_ctrl_listener = DoubleCtrlListener(listener.trigger_rephrase)
    debug_print('[DEBUG] Registered double ctrl listener')

    try:
        sys.exit(app.exec_())
    except KeyboardInterrupt:
        debug_print('[DEBUG] KeyboardInterrupt caught, exiting gracefully.')


if __name__ == '__main__':
    main()
