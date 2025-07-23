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

def is_own_window_focused():
    try:
        hwnd = win32gui.GetForegroundWindow()
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        return pid == APP_PID
    except Exception as e:
        print(f"[is_own_window_focused] Exception: {e}")
        return False

SUPPORTED_APPS = [
    'outlook.exe', 'notepad.exe', 'chrome.exe'
]

def is_supported_app_focused():
    try:
        hwnd = win32gui.GetForegroundWindow()
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        proc = psutil.Process(pid)
        exe = proc.name().lower()
        return exe in SUPPORTED_APPS
    except Exception as e:
        debug_print('[DEBUG] is_supported_app_focused error:', e)
        return False

DEBUG = bool(os.environ.get('REPHRASER_DEBUG'))

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
    'prompt': 'You are a helpful assistant that rephrases text in a clear and concise way.'
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
        try:
            debug_print('[DEBUG] api_key and api_url', settings['api_key'], settings['api_url'])
            openai.api_key = settings['api_key']
            openai.base_url = settings['api_url']

            lines = self.selected_text.split('\n')
            processed_lines = []
            code_lines_indices = []
            for idx, line in enumerate(lines):
                if (
                    line.strip().startswith('#') or
                    line.strip().startswith('%') or
                    self.is_code_like(line)
                ):
                    processed_lines.append(line)
                    code_lines_indices.append(idx)
                else:
                    processed_lines.append(f"[[REPHRASE:{idx}]] {line}")
            text_for_rephrase = '\n'.join(processed_lines)
            system_prompt = (
                settings['prompt']
                + " DO NOT change or rephrase any line that starts with #, %, or appears to be code. "
                + "Only rephrase lines that start with [[REPHRASE:IDX]]. "
                + "For any such line, rephrase only the part after the marker and keep code/context lines as they are."
            )

            messages = [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": f"Rephrase the following text as per instructions:\n\n{text_for_rephrase}"}
            ]
            response = openai.chat.completions.create(
                model=settings.get('model', 'gpt-3.5-turbo'),
                messages=messages,
                max_tokens=500,
                temperature=0.7,
                timeout=15.0
            )
            reply = response.choices[0].message.content.strip()
            debug_print('[DEBUG] Raw OpenAI response:\n', reply)
            # Remove all [[REPHRASE:idx]] tags from the reply (robust)
            cleaned_reply = re.sub(r"\[\[REPHRASE:\s*\d+\]\]\s*", "", reply, flags=re.IGNORECASE | re.MULTILINE)
            self.result_ready.emit(cleaned_reply, False)
        except Exception as e:
            debug_print('[DEBUG] error', e)
            # Clean any tags from the error string if present
            error_str = str(e)
            cleaned_error = re.sub(r"\[\[REPHRASE:\s*\d+\]\]\s*", "", error_str, flags=re.IGNORECASE | re.MULTILINE)
            self.result_ready.emit(f"Error: {cleaned_error}", True)

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
        super().closeEvent(event)

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
            result = re.sub(r"\[\[REPHRASE:\s*\d+\]\]\s*", "", result, flags=re.IGNORECASE | re.MULTILINE).lstrip('\n')
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
                    win32gui.ShowWindow(self.prev_hwnd, win32con.SW_RESTORE)
                    win32gui.SetForegroundWindow(self.prev_hwnd)
                    time.sleep(0.1)
                keyboard.press_and_release('ctrl+v')
                debug_print('[DEBUG] Pasted rephrased text.')
                notif = NotificationWindow('Rephrased text has been pasted.')
                notif.show()
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
    request_show_button = QtCore.pyqtSignal(str, int)

    def __init__(self, app):
        super().__init__()
        self.app = app
        self.last_text = ''
        self.button = None
        self.overlay = None  # Track the current overlay
        self.mouse_down_pos = None
        self.request_show_button.connect(self.show_button)
        mouse.on_button(self.on_mouse_down, buttons=mouse.LEFT, types=mouse.DOWN)
        mouse.on_button(self.on_mouse_release, buttons=mouse.LEFT, types=mouse.UP)
        keyboard.on_release(self.on_key_release)

    def on_mouse_down(self, *args, **kwargs):
        self.mouse_down_pos = mouse.get_position()

    def on_mouse_release(self, *args, **kwargs):
        if is_own_window_focused():
            return
        mouse_up_pos = mouse.get_position()
        if self.mouse_down_pos and (abs(mouse_up_pos[0] - self.mouse_down_pos[0]) > 3 or abs(mouse_up_pos[1] - self.mouse_down_pos[1]) > 3):
            time.sleep(0.25)
            self.try_show_button_with_retry(retries=5, delay=0.25)
        self.mouse_down_pos = None

    def on_key_release(self, event):
        if is_own_window_focused():
            return
        if event.name in ['left', 'right', 'up', 'down', 'a'] and (keyboard.is_pressed('shift') or keyboard.is_pressed('ctrl')):
            self.try_show_button()

    def try_show_button(self):
        # Use the more reliable retry mechanism
        self.try_show_button_with_retry()

    def try_show_button_with_retry(self, retries=10, delay=0.05):
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

        try:
            pyperclip.copy('')
        except Exception as e:
            debug_print(f"[DEBUG] Could not clear clipboard: {e}")
            return # Cannot proceed if we can't clear the clipboard

        time.sleep(0.05) # Small delay before sending keys
        keyboard.press_and_release('ctrl+c')
        
        text = ''
        for attempt in range(retries):
            time.sleep(delay)
            try:
                text = pyperclip.paste()
            except Exception as e:
                debug_print(f"[DEBUG] Could not paste on attempt {attempt+1}: {e}")
                continue # Try again

            debug_print(f'[DEBUG] Clipboard content (attempt {attempt+1}, len={len(text)}):', repr(text))
            if text:
                break
        
        if text.strip() and len(text.strip()) >= 100:
            debug_print('[DEBUG] Scheduling floating button for:', text[:50])
            self.request_show_button.emit(text, source_hwnd)
        else:
            try:
                pyperclip.copy(old_clip)
            except Exception as e:
                debug_print(f"[DEBUG] Could not restore clipboard: {e}")

            if not text.strip():
                debug_print(f'[DEBUG] Failed to get selection from clipboard. Clipboard restored.')
            else:
                debug_print(f'[DEBUG] Selection too short (len={len(text.strip())}), not showing button. Clipboard restored.')

    def show_button(self, text, source_hwnd):
        debug_print('[DEBUG] show_button called with:', repr(text))
        # Close any existing overlay (if any)
        if self.overlay is not None:
            try:
                self.overlay.close()
            except Exception as e:
                debug_print('[DEBUG] Error closing previous overlay:', e)
            self.overlay = None
        if self.button is not None:
            self.button.close()
        self.button = FloatingButton(text, source_hwnd)
        self.button.overlay_created.connect(self.track_overlay)
        self.button.show_near_cursor()
        self.last_text = ''

    def track_overlay(self, overlay):
        self.overlay = overlay

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
    import pythoncom
    from win32com.shell import shell, shellcon
    shell_link = pythoncom.CoCreateInstance(shell.CLSID_ShellLink, None, pythoncom.CLSCTX_INPROC_SERVER, shell.IID_IShellLink)
    shell_link.SetPath(exe_path)
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

class SettingsWindow(QtWidgets.QMainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle('Settings')
        self.setMinimumSize(400, 300)
        self.setWindowIcon(QtGui.QIcon(get_icon_path()))
        self.tabs = QtWidgets.QTabWidget()
        self.general_tab = QtWidgets.QWidget()
        self.parameters_tab = QtWidgets.QWidget()
        self.tabs.addTab(self.general_tab, 'General')
        self.tabs.addTab(self.parameters_tab, 'Parameters')
        self.init_general_tab()
        self.init_parameters_tab()
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
        self.api_url_edit = QtWidgets.QLineEdit()
        self.model_edit = QtWidgets.QLineEdit()
        self.prompt_edit = QtWidgets.QPlainTextEdit()
        layout.addRow('API Key:', self.api_key_edit)
        layout.addRow('API URL:', self.api_url_edit)
        layout.addRow('Model:', self.model_edit)
        layout.addRow('Prompt:', self.prompt_edit)
        self.parameters_tab.setLayout(layout)

    def load_current_settings(self):
        self.api_key_edit.setText(settings.get('api_key', ''))
        self.api_url_edit.setText(settings.get('api_url', ''))
        self.model_edit.setText(settings.get('model', 'gpt-3.5-turbo'))
        self.prompt_edit.setPlainText(settings.get('prompt', ''))

    def save_and_close(self):
        settings['api_key'] = self.api_key_edit.text().strip()
        settings['api_url'] = self.api_url_edit.text().strip()
        settings['model'] = self.model_edit.text().strip() or 'gpt-3.5-turbo'
        settings['prompt'] = self.prompt_edit.toPlainText().strip()
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

class GlobalPasteHotkey(QtCore.QObject):
    def __init__(self, parent=None):
        super().__init__(parent)
        keyboard.add_hotkey('ctrl+shift+v', self.paste_clipboard)

    def paste_clipboard(self):
        debug_print('[DEBUG] Global hotkey Ctrl+Shift+V pressed, sending Ctrl+V')
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
    try:
        sys.exit(app.exec_())
    except KeyboardInterrupt:
        debug_print('[DEBUG] KeyboardInterrupt caught, exiting gracefully.')

if __name__ == '__main__':
    main()
