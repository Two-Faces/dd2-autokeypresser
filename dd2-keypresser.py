"""
DD2 Auto KeyPresser - Automatic key pressing tool for Dungeon Defenders 2
"""
import os
import sys
import configparser
import threading
import time
import ctypes

import psutil
import win32gui
import win32process
import win32api
import win32con
from pystray import Icon, Menu as TrayMenu, MenuItem as TrayMenuItem
from PIL import Image, ImageDraw

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QListWidget, QFrame
)
from PyQt5.QtCore import Qt, QTimer, pyqtSignal, QObject
from PyQt5.QtGui import QColor, QLinearGradient, QPainter, QBrush, QPen, QIcon


# =============================================================================
# Configuration
# =============================================================================

def resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller"""
    try:
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


def load_config():
    """Load configuration from config.ini"""
    config = configparser.ConfigParser()
    encodings = ['utf-8-sig', 'utf-8', 'cp1251', 'cp1252', 'latin-1']

    try:
        cfg_path = resource_path('config.ini')
        if not os.path.exists(cfg_path):
            return config

        with open(cfg_path, 'rb') as f:
            raw = f.read()

        for enc in encodings:
            try:
                config.read_string(raw.decode(enc))
                return config
            except (UnicodeDecodeError, configparser.Error):
                continue
    except Exception:
        pass

    return config


config = load_config()
START_HOTKEY = config.get('settings', 'start_hotkey', fallback='F8')
STOP_HOTKEY = config.get('settings', 'stop_hotkey', fallback='F9')
GAME_NAME = config.get('settings', 'game_name', fallback='DunDefGame.exe')


# =============================================================================
# Virtual Key Codes
# =============================================================================

VK_CODE = {
    'backspace': 0x08, 'tab': 0x09, 'clear': 0x0C, 'enter': 0x0D,
    'shift': 0x10, 'ctrl': 0x11, 'alt': 0x12, 'pause': 0x13,
    'caps_lock': 0x14, 'esc': 0x1B, 'space': 0x20,
    'page_up': 0x21, 'page_down': 0x22, 'end': 0x23, 'home': 0x24,
    'left': 0x25, 'up': 0x26, 'right': 0x27, 'down': 0x28,
    'print_screen': 0x2C, 'insert': 0x2D, 'delete': 0x2E,
    '0': 0x30, '1': 0x31, '2': 0x32, '3': 0x33, '4': 0x34,
    '5': 0x35, '6': 0x36, '7': 0x37, '8': 0x38, '9': 0x39,
    'a': 0x41, 'b': 0x42, 'c': 0x43, 'd': 0x44, 'e': 0x45,
    'f': 0x46, 'g': 0x47, 'h': 0x48, 'i': 0x49, 'j': 0x4A,
    'k': 0x4B, 'l': 0x4C, 'm': 0x4D, 'n': 0x4E, 'o': 0x4F,
    'p': 0x50, 'q': 0x51, 'r': 0x52, 's': 0x53, 't': 0x54,
    'u': 0x55, 'v': 0x56, 'w': 0x57, 'x': 0x58, 'y': 0x59, 'z': 0x5A,
    'numpad_0': 0x60, 'numpad_1': 0x61, 'numpad_2': 0x62,
    'numpad_3': 0x63, 'numpad_4': 0x64, 'numpad_5': 0x65,
    'numpad_6': 0x66, 'numpad_7': 0x67, 'numpad_8': 0x68, 'numpad_9': 0x69,
    'multiply': 0x6A, 'add': 0x6B, 'separator': 0x6C,
    'subtract': 0x6D, 'decimal': 0x6E, 'divide': 0x6F,
    'f1': 0x70, 'f2': 0x71, 'f3': 0x72, 'f4': 0x73, 'f5': 0x74,
    'f6': 0x75, 'f7': 0x76, 'f8': 0x77, 'f9': 0x78, 'f10': 0x79,
    'f11': 0x7A, 'f12': 0x7B,
    'num_lock': 0x90, 'scroll_lock': 0x91,
    'left_shift': 0xA0, 'right_shift': 0xA1,
    'left_control': 0xA2, 'right_control': 0xA3,
    'left_menu': 0xA4, 'right_menu': 0xA5,
}

VK_TO_NAME = {v: k.upper() for k, v in VK_CODE.items()}

QT_KEY_TO_VK = {
    Qt.Key_Backspace: 0x08, Qt.Key_Tab: 0x09, Qt.Key_Return: 0x0D, Qt.Key_Enter: 0x0D,
    Qt.Key_Shift: 0x10, Qt.Key_Control: 0x11, Qt.Key_Alt: 0x12, Qt.Key_Pause: 0x13,
    Qt.Key_CapsLock: 0x14, Qt.Key_Escape: 0x1B, Qt.Key_Space: 0x20,
    Qt.Key_PageUp: 0x21, Qt.Key_PageDown: 0x22, Qt.Key_End: 0x23, Qt.Key_Home: 0x24,
    Qt.Key_Left: 0x25, Qt.Key_Up: 0x26, Qt.Key_Right: 0x27, Qt.Key_Down: 0x28,
    Qt.Key_Print: 0x2C, Qt.Key_Insert: 0x2D, Qt.Key_Delete: 0x2E,
    Qt.Key_0: 0x30, Qt.Key_1: 0x31, Qt.Key_2: 0x32, Qt.Key_3: 0x33, Qt.Key_4: 0x34,
    Qt.Key_5: 0x35, Qt.Key_6: 0x36, Qt.Key_7: 0x37, Qt.Key_8: 0x38, Qt.Key_9: 0x39,
    Qt.Key_A: 0x41, Qt.Key_B: 0x42, Qt.Key_C: 0x43, Qt.Key_D: 0x44, Qt.Key_E: 0x45,
    Qt.Key_F: 0x46, Qt.Key_G: 0x47, Qt.Key_H: 0x48, Qt.Key_I: 0x49, Qt.Key_J: 0x4A,
    Qt.Key_K: 0x4B, Qt.Key_L: 0x4C, Qt.Key_M: 0x4D, Qt.Key_N: 0x4E, Qt.Key_O: 0x4F,
    Qt.Key_P: 0x50, Qt.Key_Q: 0x51, Qt.Key_R: 0x52, Qt.Key_S: 0x53, Qt.Key_T: 0x54,
    Qt.Key_U: 0x55, Qt.Key_V: 0x56, Qt.Key_W: 0x57, Qt.Key_X: 0x58, Qt.Key_Y: 0x59,
    Qt.Key_Z: 0x5A,
    Qt.Key_F1: 0x70, Qt.Key_F2: 0x71, Qt.Key_F3: 0x72, Qt.Key_F4: 0x73, Qt.Key_F5: 0x74,
    Qt.Key_F6: 0x75, Qt.Key_F7: 0x76, Qt.Key_F8: 0x77, Qt.Key_F9: 0x78, Qt.Key_F10: 0x79,
    Qt.Key_F11: 0x7A, Qt.Key_F12: 0x7B,
    Qt.Key_NumLock: 0x90, Qt.Key_ScrollLock: 0x91,
}


# =============================================================================
# Stylesheet
# =============================================================================

STYLESHEET = """
QMainWindow {
    background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
        stop:0 #0f0f23, stop:1 #1a1a3e);
}

QLabel {
    color: #ffffff;
    background: transparent;
}

QLabel#title {
    font-size: 18px;
    font-weight: bold;
}

QLabel#section {
    font-size: 11px;
    color: #8892a0;
}

QLabel#status {
    font-size: 11px;
    color: #8892a0;
}

QLabel#footer {
    font-size: 10px;
    color: #6a7280;
}

QLineEdit {
    background-color: #2c3e50;
    border: none;
    border-radius: 15px;
    padding: 8px 15px;
    color: #ffffff;
    font-size: 13px;
    font-family: Consolas;
}

QLineEdit:focus {
    background-color: #34495e;
}

QPushButton {
    border: none;
    border-radius: 16px;
    padding: 10px 20px;
    font-size: 11px;
    font-weight: bold;
    color: #ffffff;
}

QPushButton#capture {
    background-color: #2980b9;
}
QPushButton#capture:hover {
    background-color: #3498db;
}
QPushButton#capture:disabled {
    background-color: #1a4a6e;
    color: #666666;
}

QPushButton#interval {
    background-color: #3d5166;
    padding: 8px 12px;
    border-radius: 14px;
}
QPushButton#interval:hover {
    background-color: #4a6278;
}

QPushButton#start {
    background-color: #00a854;
    font-size: 12px;
    padding: 12px 25px;
    border-radius: 18px;
}
QPushButton#start:hover {
    background-color: #00c964;
}
QPushButton#start:disabled {
    background-color: #555555;
    color: #888888;
}

QPushButton#stop {
    background-color: #555555;
    font-size: 12px;
    padding: 12px 25px;
    border-radius: 18px;
    color: #888888;
}
QPushButton#stop:enabled {
    background-color: #e74c3c;
    color: #ffffff;
}
QPushButton#stop:enabled:hover {
    background-color: #ff6b5b;
}

QListWidget {
    background-color: #1e2333;
    border: 1px solid #3d5166;
    border-radius: 14px;
    padding: 8px;
    color: #ffffff;
    font-family: Consolas;
    font-size: 11px;
}

QListWidget::item {
    padding: 4px;
}

QListWidget::item:selected {
    background-color: #3d5166;
    border-radius: 6px;
}

QFrame#keyDisplay {
    background-color: #0f3460;
    border-radius: 14px;
    padding: 5px;
}

QLabel#keyText {
    color: #ffdd57;
    font-size: 14px;
    font-weight: bold;
    font-family: Consolas;
}
"""


# =============================================================================
# Signal Emitter for thread-safe UI updates
# =============================================================================

class SignalEmitter(QObject):
    update_ui = pyqtSignal()
    update_processes = pyqtSignal()
    update_status = pyqtSignal(bool)


# =============================================================================
# Game Overlay Widget
# =============================================================================

class GameOverlay(QWidget):
    """Transparent overlay that shows status on top of the game window"""

    OVERLAY_WIDTH = 170
    OVERLAY_HEIGHT = 72
    MARGIN = 12

    def __init__(self):
        super().__init__(None)
        self.game_hwnd = None
        self.is_active = False
        self.key_name = ""
        self.interval = 0

        self._setup_window()
        self._setup_ui()
        self._setup_timer()
        self.hide()

    def _setup_window(self):
        """Configure window flags and attributes"""
        self.setWindowFlags(
            Qt.FramelessWindowHint |
            Qt.WindowStaysOnTopHint |
            Qt.Tool |
            Qt.WindowTransparentForInput
        )
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setAttribute(Qt.WA_ShowWithoutActivating)
        self.setFixedSize(self.OVERLAY_WIDTH, self.OVERLAY_HEIGHT)

    def _setup_ui(self):
        """Create overlay UI elements"""
        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 8, 12, 8)
        layout.setSpacing(2)

        # Title
        self.title_label = QLabel("DD2 KeyPresser")
        self.title_label.setStyleSheet("""
            QLabel {
                color: rgba(255, 255, 255, 140);
                font-size: 9px;
                font-weight: bold;
                font-family: Segoe UI;
                background: transparent;
            }
        """)
        layout.addWidget(self.title_label)

        # Status
        self.status_label = QLabel("Stopped")
        self.status_label.setStyleSheet(self._get_status_style(False))
        layout.addWidget(self.status_label)

        # Key info
        self.info_label = QLabel("No key set")
        self.info_label.setStyleSheet(self._get_info_style(False))
        layout.addWidget(self.info_label)

    def _setup_timer(self):
        """Setup position tracking timer"""
        self.position_timer = QTimer(self)
        self.position_timer.timeout.connect(self.update_position)
        self.position_timer.start(50)

    @staticmethod
    def _get_status_style(active):
        opacity = 255 if active else 200
        return f"""
            QLabel {{
                color: rgba(255, 255, 255, {opacity});
                font-size: 14px;
                font-weight: bold;
                font-family: Segoe UI;
                background: transparent;
            }}
        """

    @staticmethod
    def _get_info_style(active):
        opacity = 220 if active else 150
        return f"""
            QLabel {{
                color: rgba(255, 255, 255, {opacity});
                font-size: 12px;
                font-family: Consolas;
                background: transparent;
            }}
        """

    def paintEvent(self, event):
        """Draw semi-transparent rounded background with border"""
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)

        rect = self.rect().adjusted(1, 1, -1, -1)
        gradient = QLinearGradient(0, 0, 0, rect.height())

        if self.is_active:
            gradient.setColorAt(0, QColor(16, 185, 129, 240))
            gradient.setColorAt(1, QColor(5, 150, 105, 240))
            border_color = QColor(52, 211, 153, 200)
        else:
            gradient.setColorAt(0, QColor(55, 65, 81, 230))
            gradient.setColorAt(1, QColor(31, 41, 55, 230))
            border_color = QColor(75, 85, 99, 180)

        painter.setBrush(QBrush(gradient))
        painter.setPen(Qt.NoPen)
        painter.drawRoundedRect(rect, 12, 12)

        painter.setBrush(Qt.NoBrush)
        painter.setPen(QPen(border_color, 1.5))
        painter.drawRoundedRect(rect, 12, 12)

    def set_status(self, active, key_name="", interval=0):
        """Update overlay status"""
        self.is_active = active
        self.key_name = key_name
        self.interval = interval

        self.status_label.setText("● ACTIVE" if active else "○ Stopped")
        self.status_label.setStyleSheet(self._get_status_style(active))
        self.info_label.setStyleSheet(self._get_info_style(active))
        self.info_label.setText(f"{key_name}  ·  {interval}ms" if key_name else "No key set")
        self.update()

    def set_game_hwnd(self, hwnd):
        """Set the game window handle to track"""
        self.game_hwnd = hwnd
        if hwnd:
            self.update_position()
            self.show()
            self.raise_()
        else:
            self.hide()

    def update_position(self):
        """Update overlay position to follow game window"""
        if not self.game_hwnd:
            if self.isVisible():
                self.hide()
            return

        try:
            if not win32gui.IsWindow(self.game_hwnd):
                self.game_hwnd = None
                self.hide()
                return

            if not win32gui.IsWindowVisible(self.game_hwnd):
                self.hide()
                return

            # Only show when game is in focus
            if win32gui.GetForegroundWindow() != self.game_hwnd:
                if self.isVisible():
                    self.hide()
                return

            # Position at top-left of game client area
            point = win32gui.ClientToScreen(self.game_hwnd, (0, 0))
            self.move(point[0] + self.MARGIN, point[1] + self.MARGIN)

            if not self.isVisible():
                self.show()
                self.raise_()

        except Exception:
            self.game_hwnd = None
            self.hide()


# =============================================================================
# Main Application Window
# =============================================================================

class GameKeyPresserApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self._init_state()
        self._setup_window()
        self._setup_ui()
        self._setup_signals()
        self._start_services()

    def _init_state(self):
        """Initialize state variables"""
        self.selected_processes = {}
        self.key_to_press = ""
        self.key_vk_code = None
        self.press_interval = 100
        self.is_pressing = False
        self.stop_event = threading.Event()
        self.is_capturing = False
        self._hotkeys_registered = False
        self.game_hwnd = None
        self.overlay = GameOverlay()

    def _setup_window(self):
        """Configure main window"""
        self.setWindowTitle("DD2 Auto KeyPresser")
        self.setFixedSize(420, 460)
        self.setStyleSheet(STYLESHEET)

        try:
            self.setWindowIcon(QIcon(resource_path("app_icon.ico")))
        except Exception:
            pass

    def _setup_signals(self):
        """Setup signal connections"""
        self.signals = SignalEmitter()
        self.signals.update_ui.connect(self._on_stop_ui_update)
        self.signals.update_processes.connect(self._update_process_list)
        self.signals.update_status.connect(self._update_status)

    def _start_services(self):
        """Start background services"""
        self._setup_hotkeys()
        self._start_process_monitor()
        self._setup_tray()

    def _setup_ui(self):
        """Build the user interface"""
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        layout.setContentsMargins(20, 20, 20, 15)
        layout.setSpacing(12)

        self._build_header(layout)
        self._build_key_section(layout)
        self._build_interval_section(layout)
        self._build_control_buttons(layout)
        self._build_process_section(layout)
        self._build_footer(layout)

    def _build_header(self, layout):
        """Build header with title and status"""
        header = QHBoxLayout()

        title = QLabel("DD2 Auto KeyPresser")
        title.setObjectName("title")
        header.addWidget(title)
        header.addStretch()

        self.status_dot = QLabel("●")
        self.status_dot.setStyleSheet("color: #8892a0; font-size: 10px;")
        header.addWidget(self.status_dot)

        self.status_label = QLabel("Stopped")
        self.status_label.setObjectName("status")
        header.addWidget(self.status_label)

        layout.addLayout(header)

    def _build_key_section(self, layout):
        """Build key capture section"""
        label = QLabel("Key")
        label.setObjectName("section")
        layout.addWidget(label)

        row = QHBoxLayout()

        self.capture_btn = QPushButton("Capture Key")
        self.capture_btn.setObjectName("capture")
        self.capture_btn.setFixedSize(120, 34)
        self.capture_btn.setCursor(Qt.PointingHandCursor)
        self.capture_btn.clicked.connect(self._start_capture_key)
        row.addWidget(self.capture_btn)
        row.addStretch()

        key_frame = QFrame()
        key_frame.setObjectName("keyDisplay")
        key_frame.setFixedSize(90, 34)
        key_layout = QHBoxLayout(key_frame)
        key_layout.setContentsMargins(10, 0, 10, 0)

        self.key_display = QLabel("—")
        self.key_display.setObjectName("keyText")
        self.key_display.setAlignment(Qt.AlignCenter)
        key_layout.addWidget(self.key_display)
        row.addWidget(key_frame)

        layout.addLayout(row)

    def _build_interval_section(self, layout):
        """Build interval selection section"""
        label = QLabel("Interval (ms)")
        label.setObjectName("section")
        layout.addWidget(label)

        row = QHBoxLayout()

        self.interval_input = QLineEdit("100")
        self.interval_input.setFixedSize(75, 34)
        self.interval_input.setAlignment(Qt.AlignCenter)
        row.addWidget(self.interval_input)

        for ms in [50, 100, 250, 500]:
            btn = QPushButton(str(ms))
            btn.setObjectName("interval")
            btn.setFixedSize(52, 32)
            btn.setCursor(Qt.PointingHandCursor)
            btn.clicked.connect(lambda _, m=ms: self._set_interval(m))
            row.addWidget(btn)

        row.addStretch()
        layout.addLayout(row)

    def _build_control_buttons(self, layout):
        """Build start/stop control buttons"""
        row = QHBoxLayout()

        self.start_btn = QPushButton(f"START  [{START_HOTKEY}]")
        self.start_btn.setObjectName("start")
        self.start_btn.setFixedHeight(42)
        self.start_btn.setCursor(Qt.PointingHandCursor)
        self.start_btn.setEnabled(False)
        self.start_btn.clicked.connect(self.start_pressing)
        row.addWidget(self.start_btn)

        self.stop_btn = QPushButton(f"STOP  [{STOP_HOTKEY}]")
        self.stop_btn.setObjectName("stop")
        self.stop_btn.setFixedHeight(42)
        self.stop_btn.setCursor(Qt.PointingHandCursor)
        self.stop_btn.setEnabled(False)
        self.stop_btn.clicked.connect(self.stop_pressing)
        row.addWidget(self.stop_btn)

        layout.addLayout(row)

    def _build_process_section(self, layout):
        """Build process list section"""
        header = QHBoxLayout()

        label = QLabel("Processes")
        label.setObjectName("section")
        header.addWidget(label)
        header.addStretch()

        self.process_count = QLabel("Found: 0")
        self.process_count.setObjectName("section")
        header.addWidget(self.process_count)

        layout.addLayout(header)

        self.process_list = QListWidget()
        self.process_list.setFixedHeight(100)
        layout.addWidget(self.process_list)

    def _build_footer(self, layout):
        """Build footer with hotkey info"""
        layout.addStretch()

        footer = QLabel(f"Hotkeys: {START_HOTKEY} = start, {STOP_HOTKEY} = stop")
        footer.setObjectName("footer")
        footer.setAlignment(Qt.AlignCenter)
        layout.addWidget(footer)

    # -------------------------------------------------------------------------
    # Key Capture
    # -------------------------------------------------------------------------

    def _set_interval(self, ms):
        """Set the press interval"""
        self.interval_input.setText(str(ms))
        self.press_interval = ms
        self.overlay.set_status(self.is_pressing, self.key_to_press, ms)

    def _vk_to_display_name(self, vk_code):
        """Get display name for a VK code"""
        if vk_code in VK_TO_NAME:
            return VK_TO_NAME[vk_code]
        if 0x30 <= vk_code <= 0x39 or 0x41 <= vk_code <= 0x5A:
            return chr(vk_code)
        return f"KEY_{vk_code}"

    def _start_capture_key(self):
        """Start key capture mode"""
        if self.is_capturing:
            return
        self.is_capturing = True
        self.capture_btn.setText("Press a key...")
        self.key_display.setText("...")
        self.setFocus()
        self.activateWindow()

    def keyPressEvent(self, event):
        """Handle key press events for key capture"""
        if not self.is_capturing:
            super().keyPressEvent(event)
            return

        qt_key = event.key()
        if qt_key in (Qt.Key_Shift, Qt.Key_Control, Qt.Key_Alt, Qt.Key_Meta):
            return

        vk_code = QT_KEY_TO_VK.get(qt_key) or event.nativeVirtualKey()

        if vk_code:
            start_vk = VK_CODE.get(START_HOTKEY.lower())
            stop_vk = VK_CODE.get(STOP_HOTKEY.lower())

            if vk_code not in [start_vk, stop_vk]:
                self.key_vk_code = vk_code
                display_name = self._vk_to_display_name(vk_code)
                self.key_to_press = display_name
                self._finish_capture(display_name)

    def _finish_capture(self, display_name):
        """Finish key capture mode"""
        self.is_capturing = False
        self.capture_btn.setText("Capture Key")
        self.key_display.setText(display_name)

        try:
            self.press_interval = int(self.interval_input.text())
        except ValueError:
            self.press_interval = 100

        self.overlay.set_status(self.is_pressing, display_name, self.press_interval)

    # -------------------------------------------------------------------------
    # Hotkeys
    # -------------------------------------------------------------------------

    def _setup_hotkeys(self):
        """Register global hotkeys"""
        if self._hotkeys_registered:
            return

        def hotkey_thread():
            try:
                user32 = ctypes.windll.user32
                import ctypes.wintypes as wintypes

                vk_start = VK_CODE.get(START_HOTKEY.lower())
                vk_stop = VK_CODE.get(STOP_HOTKEY.lower())

                if vk_start and vk_stop:
                    user32.RegisterHotKey(None, 1, 0, vk_start)
                    user32.RegisterHotKey(None, 2, 0, vk_stop)
                    self._hotkeys_registered = True

                    msg = wintypes.MSG()
                    while user32.GetMessageW(ctypes.byref(msg), None, 0, 0):
                        if msg.message == 0x0312:  # WM_HOTKEY
                            if msg.wParam == 1:
                                QTimer.singleShot(0, self.start_pressing)
                            elif msg.wParam == 2:
                                QTimer.singleShot(0, self.stop_pressing)
            except Exception:
                pass

        threading.Thread(target=hotkey_thread, daemon=True).start()

    # -------------------------------------------------------------------------
    # Key Pressing
    # -------------------------------------------------------------------------

    def _find_game_windows(self):
        """Find all windows belonging to monitored processes"""
        pid_to_hwnd = {}

        def enum_callback(hwnd, _):
            try:
                if win32gui.IsWindow(hwnd):
                    _, pid = win32process.GetWindowThreadProcessId(hwnd)
                    if pid in self.selected_processes and pid not in pid_to_hwnd:
                        pid_to_hwnd[pid] = hwnd
            except Exception:
                pass
            return True

        try:
            win32gui.EnumWindows(enum_callback, None)
        except Exception:
            pass

        return list(pid_to_hwnd.values())

    def _send_key_to_window(self, hwnd, vk_code):
        """Send key press to a window"""
        if not vk_code:
            return
        try:
            scan_code = win32api.MapVirtualKey(vk_code, 0)
            lparam_down = (scan_code << 16) | 1
            lparam_up = (scan_code << 16) | 0xC0000001
            win32api.PostMessage(hwnd, win32con.WM_KEYDOWN, vk_code, lparam_down)
            win32api.PostMessage(hwnd, win32con.WM_KEYUP, vk_code, lparam_up)
        except Exception:
            pass

    def start_pressing(self):
        """Start the key pressing loop"""
        if self.is_pressing or not self.key_vk_code or not self.selected_processes:
            return

        try:
            self.press_interval = int(self.interval_input.text())
            if self.press_interval < 10:
                self.press_interval = 10
        except ValueError:
            return

        self.is_pressing = True
        self.stop_event.clear()

        self.start_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)
        self.capture_btn.setEnabled(False)
        self.interval_input.setEnabled(False)
        self.signals.update_status.emit(True)

        threading.Thread(target=self._press_loop, daemon=True).start()

    def stop_pressing(self):
        """Stop the key pressing loop"""
        if not self.is_pressing:
            return

        self.stop_event.set()
        self.is_pressing = False
        self.signals.update_ui.emit()

    def _press_loop(self):
        """Main key pressing loop (runs in thread)"""
        vk_code = self.key_vk_code
        interval = self.press_interval / 1000.0

        while not self.stop_event.is_set():
            windows = self._find_game_windows()
            if not windows:
                self.stop_event.wait(0.5)
                continue

            for hwnd in windows:
                if self.stop_event.is_set():
                    break
                self._send_key_to_window(hwnd, vk_code)

            self.stop_event.wait(interval)

    # -------------------------------------------------------------------------
    # UI Updates
    # -------------------------------------------------------------------------

    def _on_stop_ui_update(self):
        """Update UI when stopping"""
        self.start_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)
        self.capture_btn.setEnabled(True)
        self.interval_input.setEnabled(True)
        self.signals.update_status.emit(False)

    def _update_status(self, active):
        """Update status indicator"""
        if active:
            self.status_dot.setStyleSheet("color: #00d26a; font-size: 10px;")
            self.status_label.setText("Active")
            self.status_label.setStyleSheet("color: #00d26a;")
        else:
            self.status_dot.setStyleSheet("color: #8892a0; font-size: 10px;")
            self.status_label.setText("Stopped")
            self.status_label.setStyleSheet("color: #8892a0;")

        self.overlay.set_status(active, self.key_to_press, self.press_interval)

    # -------------------------------------------------------------------------
    # Process Monitoring
    # -------------------------------------------------------------------------

    def _start_process_monitor(self):
        """Start process monitoring thread"""
        threading.Thread(target=self._monitor_processes, daemon=True).start()

    def _monitor_processes(self):
        """Monitor for game processes (runs in thread)"""
        previous_pids = set()

        while True:
            try:
                current_pids = set()
                for proc in psutil.process_iter(['pid', 'name']):
                    try:
                        if proc.info['name'] and proc.info['name'].lower() == GAME_NAME.lower():
                            current_pids.add(proc.info['pid'])
                    except (psutil.NoSuchProcess, psutil.AccessDenied):
                        continue

                if current_pids != previous_pids:
                    self.selected_processes = {
                        pid: f"{GAME_NAME} (PID: {pid})" for pid in current_pids
                    }
                    self.signals.update_processes.emit()
                    previous_pids = current_pids.copy()
            except Exception:
                pass

            time.sleep(2)

    def _update_process_list(self):
        """Update the process list UI"""
        self.process_list.clear()

        if self.selected_processes:
            for info in self.selected_processes.values():
                self.process_list.addItem(f"  > {info}")
            self.process_count.setText(f"Found: {len(self.selected_processes)}")

            if not self.is_pressing:
                self.start_btn.setEnabled(True)
            self._update_game_hwnd()
        else:
            self.process_list.addItem(f"  Waiting for {GAME_NAME}...")
            self.process_count.setText("Found: 0")

            if not self.is_pressing:
                self.start_btn.setEnabled(False)
            self.game_hwnd = None
            self.overlay.set_game_hwnd(None)

    def _update_game_hwnd(self):
        """Find game window handle and update overlay"""
        if not self.selected_processes:
            self.game_hwnd = None
            self.overlay.set_game_hwnd(None)
            return

        for pid in self.selected_processes.keys():
            result = []

            def enum_callback(hwnd, result):
                try:
                    if win32gui.IsWindow(hwnd) and win32gui.IsWindowVisible(hwnd):
                        _, window_pid = win32process.GetWindowThreadProcessId(hwnd)
                        if window_pid == pid and win32gui.GetWindowText(hwnd):
                            result.append(hwnd)
                except Exception:
                    pass
                return True

            try:
                win32gui.EnumWindows(enum_callback, result)
                if result:
                    self.game_hwnd = result[0]
                    self.overlay.set_game_hwnd(self.game_hwnd)
                    self.overlay.set_status(
                        self.is_pressing,
                        self.key_to_press,
                        self.press_interval
                    )
                    return
            except Exception:
                pass

        self.game_hwnd = None
        self.overlay.set_game_hwnd(None)

    # -------------------------------------------------------------------------
    # System Tray
    # -------------------------------------------------------------------------

    def _setup_tray(self):
        """Setup system tray icon"""
        try:
            image = Image.open(resource_path("app_icon.ico"))
        except Exception:
            image = Image.new('RGB', (64, 64), (45, 62, 80))
            draw = ImageDraw.Draw(image)
            draw.ellipse((8, 8, 56, 56), fill=(52, 152, 219))
            draw.text((22, 20), "K", fill=(255, 255, 255))

        menu = TrayMenu(
            TrayMenuItem('Show', self._show_from_tray, default=True),
            TrayMenuItem('Start', lambda: QTimer.singleShot(0, self.start_pressing)),
            TrayMenuItem('Stop', lambda: QTimer.singleShot(0, self.stop_pressing)),
            TrayMenu.SEPARATOR,
            TrayMenuItem('Exit', self._exit_app)
        )

        self.tray_icon = Icon("DD2 KeyPresser", image, "DD2 KeyPresser", menu)
        threading.Thread(target=self.tray_icon.run, daemon=True).start()

    def _show_from_tray(self, icon=None, item=None):
        """Show window from tray"""
        QTimer.singleShot(0, self._show_window)

    def _show_window(self):
        """Show and activate window"""
        self.show()
        self.activateWindow()
        self.raise_()

    def closeEvent(self, event):
        """Handle window close - minimize to tray"""
        event.ignore()
        self.hide()
        if hasattr(self, 'tray_icon') and self.tray_icon:
            self.tray_icon.notify("DD2 KeyPresser", "Minimized to tray")

    def _exit_app(self, icon=None, item=None):
        """Exit application"""
        self.stop_event.set()

        if hasattr(self, 'overlay') and self.overlay:
            self.overlay.close()

        if hasattr(self, 'tray_icon') and self.tray_icon:
            try:
                self.tray_icon.stop()
            except Exception:
                pass

        QApplication.quit()


# =============================================================================
# Entry Point
# =============================================================================

def main():
    # Enable DPI awareness
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass

    app = QApplication(sys.argv)
    app.setStyle('Fusion')

    window = GameKeyPresserApp()
    window.show()

    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
