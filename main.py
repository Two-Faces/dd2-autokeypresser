import os
import sys
import configparser
import tkinter as tk
from tkinter import ttk, messagebox
import psutil
import win32gui
import win32api
import win32con
import win32process
import keyboard
import threading
import time
import ttkbootstrap as ttkb
from ttkbootstrap.constants import *
from pystray import Icon, Menu as TrayMenu, MenuItem as TrayMenuItem
from PIL import Image, ImageDraw

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

# Load configuration
config = configparser.ConfigParser()
config.read(resource_path('config.ini'))

START_HOTKEY = config.get('settings', 'start_hotkey')
STOP_HOTKEY = config.get('settings', 'stop_hotkey')
GAME_NAME = config.get('settings', 'game_name')

VK_CODE = {
    'backspace': 0x08,
    'tab': 0x09,
    'clear': 0x0C,
    'enter': 0x0D,
    'shift': 0x10,
    'ctrl': 0x11,
    'alt': 0x12,
    'pause': 0x13,
    'caps_lock': 0x14,
    'esc': 0x1B,
    'space': 0x20,
    'page_up': 0x21,
    'page_down': 0x22,
    'end': 0x23,
    'home': 0x24,
    'left': 0x25,
    'up': 0x26,
    'right': 0x27,
    'down': 0x28,
    'print_screen': 0x2C,
    'insert': 0x2D,
    'delete': 0x2E,
    '0': 0x30,
    '1': 0x31,
    '2': 0x32,
    '3': 0x33,
    '4': 0x34,
    '5': 0x35,
    '6': 0x36,
    '7': 0x37,
    '8': 0x38,
    '9': 0x39,
    'a': 0x41,
    'b': 0x42,
    'c': 0x43,
    'd': 0x44,
    'e': 0x45,
    'f': 0x46,
    'g': 0x47,
    'h': 0x48,
    'i': 0x49,
    'j': 0x4A,
    'k': 0x4B,
    'l': 0x4C,
    'm': 0x4D,
    'n': 0x4E,
    'o': 0x4F,
    'p': 0x50,
    'q': 0x51,
    'r': 0x52,
    's': 0x53,
    't': 0x54,
    'u': 0x55,
    'v': 0x56,
    'w': 0x57,
    'x': 0x58,
    'y': 0x59,
    'z': 0x5A,
    'numpad_0': 0x60,
    'numpad_1': 0x61,
    'numpad_2': 0x62,
    'numpad_3': 0x63,
    'numpad_4': 0x64,
    'numpad_5': 0x65,
    'numpad_6': 0x66,
    'numpad_7': 0x67,
    'numpad_8': 0x68,
    'numpad_9': 0x69,
    'multiply': 0x6A,
    'add': 0x6B,
    'separator': 0x6C,
    'subtract': 0x6D,
    'decimal': 0x6E,
    'divide': 0x6F,
    'f1': 0x70,
    'f2': 0x71,
    'f3': 0x72,
    'f4': 0x73,
    'f5': 0x74,
    'f6': 0x75,
    'f7': 0x76,
    'f8': 0x77,
    'f9': 0x78,
    'f10': 0x79,
    'f11': 0x7A,
    'f12': 0x7B,
    'f13': 0x7C,
    'f14': 0x7D,
    'f15': 0x7E,
    'f16': 0x7F,
    'f17': 0x80,
    'f18': 0x81,
    'f19': 0x82,
    'f20': 0x83,
    'f21': 0x84,
    'f22': 0x85,
    'f23': 0x86,
    'f24': 0x87,
    'num_lock': 0x90,
    'scroll_lock': 0x91,
    'left_shift': 0xA0,
    'right_shift': 0xA1,
    'left_control': 0xA2,
    'right_control': 0xA3,
    'left_menu': 0xA4,
    'right_menu': 0xA5,
}

class GameKeyPresserApp:
    def __init__(self, root):
        self.root = root
        self.root.title("DD2 KeyPresser")
        self.root.iconbitmap(resource_path("app_icon.ico"))  # Установка иконки окна
        self.root.resizable(False, False)  # Запрет на изменение размера окна

        self.selected_processes = {}
        self.key_to_press = tk.StringVar(value="Press a key...")
        self.press_interval = tk.StringVar(value='100')  # Default interval 100ms

        self.create_widgets()

        # Добавляем бинды для горячих клавиш
        keyboard.add_hotkey(START_HOTKEY, self.start_pressing)
        keyboard.add_hotkey(STOP_HOTKEY, self.stop_pressing)

        # Запускаем поток для отслеживания процессов
        self.monitor_thread = threading.Thread(
            target=self.monitor_processes, daemon=True)
        self.monitor_thread.start()

        # Добавляем функциональность трея
        #self.tray_icon = None
        #self.create_tray_icon()

    def create_widgets(self):
        # Создаем фрейм для всех элементов интерфейса
        frame = ttkb.Frame(self.root, padding="10")
        frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        self.capture_key_button = ttkb.Button(
            frame, text="Capture Key (Press a key...)", command=self.capture_key, bootstyle="info")
        self.capture_key_button.grid(row=0, column=0, columnspan=4, padx=5, pady=5, sticky=(tk.W, tk.E))

        ttkb.Label(frame, text="Press Interval (ms):").grid(
            row=1, column=0, padx=5, pady=5, sticky=tk.W)
        self.interval_entry = ttkb.Entry(
            frame, textvariable=self.press_interval, width=10)
        self.interval_entry.grid(row=1, column=1, padx=5, pady=5, sticky=(tk.W, tk.E))

        self.start_button = ttkb.Button(
            frame, text=f"Start ({START_HOTKEY})", command=self.start_pressing, bootstyle="primary", width=10)
        self.start_button.grid(row=1, column=2, padx=5, pady=5)

        self.stop_button = ttkb.Button(
            frame, text=f"Stop ({STOP_HOTKEY})", command=self.stop_pressing, state=tk.DISABLED, bootstyle="danger", width=10)
        self.stop_button.grid(row=1, column=3, padx=5, pady=5)

        ttkb.Label(frame, text="Selected Processes:").grid(
            row=2, column=0, padx=5, pady=5, sticky=tk.W)
        self.selected_process_frame = ttkb.Frame(frame)
        self.selected_process_frame.grid(
            row=2, column=1, columnspan=3, padx=5, pady=5, sticky=(tk.W, tk.E))

        self.no_process_label = ttkb.Label(self.selected_process_frame, text="No processes available")
        self.no_process_label.pack(fill=tk.X, pady=2)

        # Устанавливаем растяжение колонок
        frame.columnconfigure(0, weight=1)
        frame.columnconfigure(1, weight=1)
        frame.columnconfigure(2, weight=1)
        frame.columnconfigure(3, weight=1)

    def capture_key(self):
        self.capture_key_button.config(text="Capture Key (Press a key...)")
        self.capture_key_button.update()

        key = keyboard.read_event(suppress=True)
        if key.event_type == keyboard.KEY_DOWN:
            self.key_to_press.set(key.name.upper())
            self.capture_key_button.config(text=f"Capture Key ({key.name.upper()})")

    def find_game_window(self, pid):
        def callback(hwnd, param):
            if win32gui.IsWindowVisible(hwnd) and win32gui.IsWindowEnabled(hwnd):
                _, found_pid = win32process.GetWindowThreadProcessId(hwnd)
                if found_pid == param:
                    windows.append(hwnd)

        windows = []
        win32gui.EnumWindows(callback, pid)
        return windows

    def send_key_to_window(self, hwnd, key):
        vk_code = VK_CODE.get(key.lower(), None)
        if vk_code:
            win32api.PostMessage(hwnd, win32con.WM_KEYDOWN, vk_code, 0)
            time.sleep(0.05)
            win32api.PostMessage(hwnd, win32con.WM_KEYUP, vk_code, 0)

    def start_pressing(self):
        if not self.key_to_press.get() or self.key_to_press.get() == "Press a key...":
            messagebox.showwarning(
                "Input Error", "Please enter a key to press.")
            return

        try:
            interval = int(self.press_interval.get()) / 1000.0  # Convert to seconds
        except ValueError:
            messagebox.showwarning(
                "Input Error", "Please enter a valid interval in milliseconds.")
            return

        self.stop_event = threading.Event()
        self.press_thread = threading.Thread(
            target=self.press_key, args=(interval,))
        self.press_thread.start()

        self._toggle_inputs(state=tk.DISABLED)
        self.start_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.NORMAL)
        self.root.protocol("WM_DELETE_WINDOW", self.hide_window)

    def press_key(self, interval):
        hwnds = []
        for pid in self.selected_processes.keys():
            hwnds += self.find_game_window(pid)

        hwnds = [hwnd for hwnd in hwnds if hwnd]

        if not hwnds:
            messagebox.showerror("Error", "Could not find any game windows.")
            self.stop_pressing()
            return

        key = self.key_to_press.get().upper()

        while not self.stop_event.is_set():
            for hwnd in hwnds:
                self.send_key_to_window(hwnd, key)
            time.sleep(interval)

    def stop_pressing(self):
        if hasattr(self, 'stop_event'):
            self.stop_event.set()
        self.root.after(0, self._toggle_inputs, tk.NORMAL)
        self.root.after(0, self.start_button.config, {"state": tk.NORMAL})
        self.root.after(0, self.stop_button.config, {"state": tk.DISABLED})

    def _toggle_inputs(self, state):
        self.capture_key_button.config(state=state)
        self.interval_entry.config(state=state)
        self.start_button.config(state=state)
        self.stop_button.config(state=tk.DISABLED)

    def monitor_processes(self):
        previous_processes = set()
        while True:
            current_processes = {p.pid for p in psutil.process_iter(['pid', 'name']) if p.info['name'].lower() == GAME_NAME.lower()}
            new_processes = current_processes - previous_processes
            terminated_processes = previous_processes - current_processes

            for pid in new_processes:
                print(f"Adding new process: {pid}")  # Debugging line
                self.selected_processes[pid] = f"{GAME_NAME} (PID: {pid})"
                self.update_selected_processes()

            for pid in terminated_processes:
                if pid in self.selected_processes:
                    print(f"Removing terminated process: {pid}")  # Debugging line
                    del self.selected_processes[pid]
                    self.update_selected_processes()

            previous_processes = current_processes
            time.sleep(1)

    def update_selected_processes(self):
        self.root.after(0, self._update_selected_processes)

    def _update_selected_processes(self):
        print("Updating process list in the UI")  # Debugging line
        for widget in self.selected_process_frame.winfo_children():
            widget.destroy()

        if not self.selected_processes:
            self.no_process_label = ttkb.Label(self.selected_process_frame, text="No processes available")
            self.no_process_label.pack(fill=tk.X, pady=2)
        else:
            self.no_process_label.pack_forget()
            for pid, process_info in self.selected_processes.items():
                frame = ttkb.Frame(self.selected_process_frame)
                frame.pack(fill=tk.X, pady=2)
                label = ttkb.Label(frame, text=process_info)
                label.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.selected_process_frame.update_idletasks()  # Update the UI explicitly

    def hide_window(self):
        self.root.withdraw()
        self.show_tray_message("DD2 KeyPresser", "Application is still running. Right-click the tray icon for options.")

    def show_tray_message(self, title, message):
        if self.tray_icon:
            self.tray_icon.show_message(title, message)

    def create_tray_icon(self):
        # Create a simple icon for the tray
        image = Image.new('RGB', (64, 64), (0, 0, 0))
        draw = ImageDraw.Draw(image)
        draw.rectangle((0, 0, 63, 63), outline=(255, 255, 255), width=2)
        draw.text((10, 25), "DD2", fill=(255, 255, 255))

        menu = TrayMenu(
            TrayMenuItem('Start', self.on_tray_start),
            TrayMenuItem('Stop', self.on_tray_stop),
            TrayMenuItem('Exit', self.on_tray_exit)
        )

        self.tray_icon = Icon("DD2 KeyPresser", image, "DD2 KeyPresser", menu)
        self.tray_icon.run_detached()

    def on_tray_start(self, icon, item):
        self.root.after(0, self.start_pressing)

    def on_tray_stop(self, icon, item):
        self.root.after(0, self.stop_pressing)

    def on_tray_exit(self, icon, item):
        self.root.after(0, self.exit_app)

    def on_tray_double_click(self, icon, item):
        self.root.after(0, self.show_window)

    def show_window(self):
        self.root.deiconify()
        self.root.lift()

    def exit_app(self):
        self.root.after(0, self.stop_pressing)
        self.root.after(0, self.tray_icon.stop)
        self.root.after(0, self.root.quit)

    def on_closing(self):
        self.stop_pressing()
        self.root.destroy()

if __name__ == "__main__":
    root = ttkb.Window(themename="darkly")
    app = GameKeyPresserApp(root)
    root.mainloop()
