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
http.client.HTTPConnection.debuglevel = 1

import logging
logging.basicConfig()
logging.getLogger().setLevel(logging.DEBUG)
requests_log = logging.getLogger("urllib3")
requests_log.setLevel(logging.DEBUG)
requests_log.propagate = True


APP_PID = os.getpid()

# Helper to check if our own window is focused
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

def debug_print(*args, **kwargs):
    if DEBUG:
        print(*args, **kwargs)

SETTINGS_FILE = 'settings.json'
DEFAULT_SETTINGS = {
    'api_key': '',
    'api_url': 'https://api.openai.com/v1',
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
    # After saving, reload to ensure consistency
    load_settings()

# Initial load
load_settings()

class FloatingButton(QtWidgets.QWidget):
    def __init__(self, selected_text, parent=None):
        super().__init__(parent)
        self.selected_text = selected_text
        self.setWindowFlags(
            QtCore.Qt.FramelessWindowHint |
            QtCore.Qt.WindowStaysOnTopHint |
            QtCore.Qt.Tool
        )
        self.setAttribute(QtCore.Qt.WA_TranslucentBackground)
        self.init_ui()
        # Auto-hide after 5 seconds
        QtCore.QTimer.singleShot(5000, self.close)

    def init_ui(self):
        layout = QtWidgets.QVBoxLayout()
        self.button = QtWidgets.QPushButton()
        self.button.setIcon(QtGui.QIcon('icon.png'))
        self.button.setIconSize(QtCore.QSize(40, 40))  # <-- Make icon bigger
        self.button.setFixedSize(65, 65)               # <-- Make button bigger
        self.button.setStyleSheet('border: none; background: transparent;')
        self.button.clicked.connect(self.rephrase_text)
        layout.addWidget(self.button)
        self.setLayout(layout)
        self.setFixedSize(65, 65)                      # <-- Make widget bigger

    def rephrase_text(self):
        self.button.setEnabled(False)
        self.overlay = RephraseOverlay(self.selected_text)
        self.overlay.show_near_cursor()
        self.close()

    def show_near_cursor(self):
        pos = QtGui.QCursor.pos()
        self.move(pos.x() + 20, pos.y())  # 20px to the right, same y
        self.show()
        debug_print('[DEBUG] FloatingButton shown at', pos.x() + 20, pos.y())

class RephraseWorker(QtCore.QThread):
    result_ready = QtCore.pyqtSignal(str, bool)  # (result, is_error)

    def __init__(self, selected_text):
        super().__init__()
        self.selected_text = selected_text

    def run(self):
        try:
            debug_print('[DEBUG] api_key and api_url', settings['api_key'], settings['api_url'])
            openai.api_key = settings['api_key']
            openai.base_url = settings['api_url']
            response = openai.chat.completions.create(
                model="gpt-3.5-turbo",
                messages=[
                    {"role": "system", "content": settings['prompt']},
                    {"role": "user", "content": f"Rephrase the following text: {self.selected_text}"}
                ],
                max_tokens=500,
                temperature=0.7
            )
            debug_print('[DEBUG] response', response);
            rephrased = response.choices[0].message.content.strip()
            self.result_ready.emit(rephrased, False)
        except Exception as e:
            debug_print('[DEBUG] error', e)
            self.result_ready.emit(f"Error: {str(e)}", True)

class RephraseOverlay(QtWidgets.QWidget):
    def __init__(self, selected_text, parent=None):
        super().__init__(parent)
        self.selected_text = selected_text
        self.setWindowFlags(QtCore.Qt.FramelessWindowHint | QtCore.Qt.WindowStaysOnTopHint | QtCore.Qt.Tool)
        self.setAttribute(QtCore.Qt.WA_TranslucentBackground)
        self.prev_hwnd = win32gui.GetForegroundWindow()
        self.init_ui()
        self.get_rephrased_text()

    def init_ui(self):
        layout = QtWidgets.QVBoxLayout()
        # Add close button (X) at the top right
        close_btn = QtWidgets.QPushButton('✕')
        close_btn.setFixedSize(24, 24)
        close_btn.setStyleSheet('border: none; background: transparent; font-size: 16px; color: #888;')
        close_btn.clicked.connect(self.close)
        close_layout = QtWidgets.QHBoxLayout()
        close_layout.addStretch()
        close_layout.addWidget(close_btn)
        layout.addLayout(close_layout)
        self.text_label = QtWidgets.QLabel()
        self.text_label.setWordWrap(True)
        self.text_label.setStyleSheet("background: #e0ffe0; padding: 8px; border-radius: 6px; font-size: 14px;")
        layout.addWidget(self.text_label)
        self.setLayout(layout)
        self.setMinimumSize(200, 60)
        self.setMaximumSize(1200, 800)
        self.text_label.installEventFilter(self)

    def get_rephrased_text(self):
        self.worker = RephraseWorker(self.selected_text)
        self.worker.result_ready.connect(self.on_result_ready)
        self.worker.start()

    def on_result_ready(self, result, is_error):
        if is_error:
            self.text_label.setText(result)
            self.text_label.setStyleSheet("background: #ffe0e0; padding: 8px; border-radius: 6px; font-size: 14px;")
        else:
            self.text_label.setText(result)
            self.text_label.setStyleSheet("background: #e0ffe0; padding: 8px; border-radius: 6px; font-size: 14px;")
        self.adjust_size_to_text()
        self.show()
        self.raise_()
        self.activateWindow()

    def adjust_size_to_text(self):
        font = self.text_label.font()
        metrics = QtGui.QFontMetrics(font)
        lines = self.text_label.text().splitlines() or ['']
        max_line_width = max((metrics.width(line) for line in lines), default=200)
        width = min(max(max_line_width + 40, 200), 1200)
        height = min(max(metrics.height() * len(lines) + 40, 60), 800)
        self.resize(width, height)

    def eventFilter(self, obj, event):
        if obj == self.text_label and event.type() == QtCore.QEvent.MouseButtonPress:
            rephrased = self.text_label.text()
            pyperclip.copy('')  # Clear clipboard first
            time.sleep(0.05)
            pyperclip.copy(rephrased)
            debug_print('[DEBUG] Copied rephrased text to clipboard')
            self.hide()
            self.close()
            notif = NotificationWindow('Rephrased text copied!<br>Click back into your app and press Ctrl+V to paste.')
            notif.show()
            return True
        return super().eventFilter(obj, event)

    def show_near_cursor(self):
        # Show click catcher first
        # if not self.click_catcher:
        #     self.click_catcher = ClickCatcher(self)
        pos = QtGui.QCursor.pos()
        self.move(pos.x() + 10, pos.y() + 10)
        self.show()
        self.raise_()
        self.activateWindow()
        debug_print('[DEBUG] RephraseOverlay shown at', pos.x() + 10, pos.y() + 10)

class SelectionListener(QtCore.QObject):
    request_show_button = QtCore.pyqtSignal(str)

    def __init__(self, app):
        super().__init__()
        self.app = app
        self.last_text = ''
        self.button = None
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
        # Only trigger if mouse was dragged (distance > threshold)
        if self.mouse_down_pos and (abs(mouse_up_pos[0] - self.mouse_down_pos[0]) > 3 or abs(mouse_up_pos[1] - self.mouse_down_pos[1]) > 3):
            self.try_show_button()
        self.mouse_down_pos = None

    def on_key_release(self, event):
        if is_own_window_focused():
            return
        # Only trigger on likely selection keys
        if event.name in ['left', 'right', 'up', 'down', 'a'] and (keyboard.is_pressed('shift') or keyboard.is_pressed('ctrl')):
            self.try_show_button()

    def try_show_button(self):
        if not is_supported_app_focused():
            debug_print('[DEBUG] Not a supported app, not showing button.')
            return
        old_clip = pyperclip.paste()
        keyboard.press_and_release('ctrl+c')
        time.sleep(0.1)
        text = pyperclip.paste()
        debug_print('[DEBUG] Clipboard content:', repr(text))
        if text != old_clip:
            pyperclip.copy(old_clip)
        # Only show button for long enough text
        if text.strip() and len(text.strip()) >= 100:
            debug_print('[DEBUG] Scheduling floating button for:', text[:50])
            self.request_show_button.emit(text)
        else:
            debug_print('[DEBUG] Selection too short, not showing button.')

    def show_button(self, text):
        debug_print('[DEBUG] show_button called with:', repr(text))
        if self.button is not None:
            self.button.close()
        self.button = FloatingButton(text)
        self.button.show_near_cursor()
        # Reset last_text so repeated selections work
        self.last_text = ''

class SettingsWindow(QtWidgets.QMainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle('Settings')
        self.setMinimumSize(400, 300)
        self.setWindowIcon(QtGui.QIcon('icon.png'))  # Set custom icon
        self.tabs = QtWidgets.QTabWidget()
        self.general_tab = QtWidgets.QWidget()
        self.parameters_tab = QtWidgets.QWidget()
        self.tabs.addTab(self.general_tab, 'General')
        self.tabs.addTab(self.parameters_tab, 'Parameters')
        self.init_general_tab()
        self.init_parameters_tab()
        btn_box = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Save | QtWidgets.QDialogButtonBox.Cancel)
        btn_box.accepted.connect(self.save_and_close)
        btn_box.rejected.connect(self.close)  # Only close the window
        layout = QtWidgets.QVBoxLayout()
        layout.addWidget(self.tabs)
        layout.addWidget(btn_box)
        container = QtWidgets.QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)
        self.load_current_settings()

    def init_general_tab(self):
        layout = QtWidgets.QVBoxLayout()
        # Placeholder for future general settings
        layout.addStretch()
        self.general_tab.setLayout(layout)

    def init_parameters_tab(self):
        layout = QtWidgets.QFormLayout()
        self.api_key_edit = QtWidgets.QLineEdit()
        self.api_url_edit = QtWidgets.QLineEdit()
        self.prompt_edit = QtWidgets.QPlainTextEdit()
        layout.addRow('API Key:', self.api_key_edit)
        layout.addRow('API URL:', self.api_url_edit)
        layout.addRow('Prompt:', self.prompt_edit)
        self.parameters_tab.setLayout(layout)

    def load_current_settings(self):
        self.api_key_edit.setText(settings.get('api_key', ''))
        self.api_url_edit.setText(settings.get('api_url', ''))
        self.prompt_edit.setPlainText(settings.get('prompt', ''))

    def save_and_close(self):
        settings['api_key'] = self.api_key_edit.text().strip()
        settings['api_url'] = self.api_url_edit.text().strip()
        settings['prompt'] = self.prompt_edit.toPlainText().strip()
        save_settings()
        self.close()

class SystemTrayIcon(QtWidgets.QSystemTrayIcon):
    def __init__(self, app, parent=None):
        icon = QtGui.QIcon('icon.png')
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
        # Do not set parent to hidden_main
        if self.settings_window is None or not self.settings_window.isVisible():
            self.settings_window = SettingsWindow()  # parent=None
            self.settings_window.show()
        else:
            self.settings_window.raise_()
            self.settings_window.activateWindow()

    def exit_app(self):
        # Unhook all listeners before exit
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
        # Center on screen
        screen = QtWidgets.QApplication.primaryScreen().availableGeometry()
        self.move(screen.center() - self.rect().center())
        QtCore.QTimer.singleShot(duration, self.close)

# Add a global hotkey for paste
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
    app.setWindowIcon(QtGui.QIcon('icon.png'))  # Set global app icon for taskbar
    app.setQuitOnLastWindowClosed(False)  # Prevent app from quitting when dialogs close
    hidden_main = QtWidgets.QMainWindow()
    hidden_main.setWindowIcon(QtGui.QIcon('icon.png'))
    hidden_main.setWindowTitle('GRephraser')
    hidden_main.setGeometry(-10000, -10000, 100, 100)  # Move off-screen
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