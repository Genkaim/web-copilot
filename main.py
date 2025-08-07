import sys
import os
import ctypes
import threading
import ctypes.wintypes

from PyQt6.QtCore import Qt, QUrl, QStandardPaths, pyqtSignal, QTimer
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QFrame, QSystemTrayIcon, QMenu, QStyle
)
from PyQt6.QtGui import QAction
from PyQt6.QtGui import QGuiApplication, QIcon
from PyQt6.QtWebEngineWidgets import QWebEngineView
from PyQt6.QtWebEngineCore import QWebEngineProfile, QWebEnginePage


class WeiboWindow(QMainWindow):
    EDGE_MARGIN = 10
    toggle_signal = pyqtSignal()

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Copilot")
        self.setMouseTracking(True)
        self._resizing = False
        self._always_on_top = False  # 置顶状态标志

        self.toggle_signal.connect(self.toggle_visibility)

        # --- 设置缓存和存储路径，保证路径存在 ---
        cache_path = os.path.join(
            QStandardPaths.writableLocation(QStandardPaths.StandardLocation.CacheLocation),
            "web_cache"
        )
        storage_path = os.path.join(
            QStandardPaths.writableLocation(QStandardPaths.StandardLocation.AppDataLocation),
            "web_storage"
        )
        os.makedirs(cache_path, exist_ok=True)
        os.makedirs(storage_path, exist_ok=True)

        # 创建持久化的 WebEngineProfile
        self.profile = QWebEngineProfile("MyProfile", self)
        self.profile.setCachePath(cache_path)
        self.profile.setPersistentStoragePath(storage_path)
        self.profile.setPersistentCookiesPolicy(
            QWebEngineProfile.PersistentCookiesPolicy.ForcePersistentCookies
        )

        # 可选：设置自定义 UserAgent
        # self.profile.setHttpUserAgent("Your Custom User Agent String")

        # 创建页面并加载
        self.page = QWebEnginePage(self.profile, self)
        self.web_view = QWebEngineView()
        self.web_view.setPage(self.page)
        self.setCentralWidget(self.web_view)
        self.web_view.load(QUrl("https://chat.deepseek.com"))

        # 右侧细线，用于视觉提示可拖拽
        self.right_edge_line = QFrame(self)
        self.right_edge_line.setFrameShape(QFrame.Shape.VLine)
        self.right_edge_line.setLineWidth(1)
        self.right_edge_line.setStyleSheet("background-color: gray;")
        self.right_edge_line.show()

    def show_custom_size(self, width=900):
        screen = QGuiApplication.primaryScreen()
        available_geometry = screen.availableGeometry()
        # 设置无边框且置顶
        self.setWindowFlags(
            Qt.WindowType.FramelessWindowHint |
            Qt.WindowType.WindowStaysOnTopHint
        )
        self.setMouseTracking(True)
        left = 0
        top = available_geometry.top()
        height = available_geometry.height()
        self.setGeometry(left, top, width, height)
        self.show()

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self.right_edge_line.setGeometry(
            self.width() - 1, 0, 1, self.height()
        )

    def mouseMoveEvent(self, event):
        x = event.pos().x()
        # 更改鼠标样式
        if self.width() - self.EDGE_MARGIN <= x <= self.width():
            self.setCursor(Qt.CursorShape.SizeHorCursor)
        else:
            self.setCursor(Qt.CursorShape.ArrowCursor)

        if self._resizing:
            global_pos = self.mapToGlobal(event.pos())
            new_width = global_pos.x() - self.geometry().x()
            if new_width >= 300:
                self.resize(new_width, self.height())

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            x = event.pos().x()
            if self.width() - self.EDGE_MARGIN <= x <= self.width():
                self._resizing = True

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self._resizing = False

    def toggle_visibility(self):
        if self.isVisible():
            self.hide()
        else:
            self.show()
            self.raise_()
            self.activateWindow()
            # 如果已经设置了置顶，则保持置顶
            if self._always_on_top:
                self.setWindowFlags(
                    self.windowFlags() |
                    Qt.WindowType.WindowStaysOnTopHint
                )
                self.show()

    def toggle_always_on_top(self):
        """切换窗口置顶状态"""
        self._always_on_top = not self._always_on_top
        if self._always_on_top:
            self.setWindowFlags(
                self.windowFlags() |
                Qt.WindowType.WindowStaysOnTopHint
            )
        else:
            self.setWindowFlags(
                self.windowFlags() &
                ~Qt.WindowType.WindowStaysOnTopHint
            )
        self.show()
        self.raise_()
        self.activateWindow()

    def closeEvent(self, event):
        # 拦截关闭，隐藏窗口，保留缓存
        event.ignore()
        self.hide()


# 全局热键设置
HOTKEY_ID = 0x0001
MOD_CONTROL = 0x0002
MOD_SHIFT = 0x0004
VK_C = 0x43


def listen_copilot_key(window_ref: WeiboWindow):
    if not ctypes.windll.user32.RegisterHotKey(
        None, HOTKEY_ID, MOD_CONTROL | MOD_SHIFT, VK_C
    ):
        error_code = ctypes.GetLastError()
        print(f"❌ 注册热键失败（Ctrl+Shift+C），错误码：{error_code}")
        return

    try:
        msg = ctypes.wintypes.MSG()
        while ctypes.windll.user32.GetMessageA(
            ctypes.byref(msg), None, 0, 0
        ) != 0:
            if msg.message == 0x0312 and msg.wParam == HOTKEY_ID:
                window_ref.toggle_signal.emit()
            ctypes.windll.user32.TranslateMessage(ctypes.byref(msg))
            ctypes.windll.user32.DispatchMessageA(ctypes.byref(msg))
    finally:
        ctypes.windll.user32.UnregisterHotKey(None, HOTKEY_ID)


def create_tray_icon(app: QApplication, window: WeiboWindow) -> QSystemTrayIcon:
    tray = QSystemTrayIcon(parent=app)
    icon_path = os.path.join(os.path.dirname(__file__), "icon.ico")
    if os.path.exists(icon_path):
        tray.setIcon(QIcon(icon_path))
    else:
        tray.setIcon(app.style().standardIcon(
            QStyle.StandardPixmap.SP_ComputerIcon
        ))

    tray.setToolTip("Copilot")

    # 为菜单和动作指定 parent 防止被回收
    menu = QMenu(parent=tray)
    show_action = QAction("显示窗口", parent=tray)
    toggle_top_action = QAction("切换置顶", parent=tray)
    quit_action = QAction("退出程序", parent=tray)

    show_action.triggered.connect(window.show)
    toggle_top_action.triggered.connect(window.toggle_always_on_top)
    quit_action.triggered.connect(app.quit)

    menu.addAction(show_action)
    menu.addAction(toggle_top_action)
    menu.addSeparator()
    menu.addAction(quit_action)

    tray.setContextMenu(menu)
    tray.show()
    return tray


if __name__ == "__main__":
    # 打包后隐藏控制台（仅在 exe 下生效）
    if getattr(sys, 'frozen', False):
        ctypes.windll.user32.ShowWindow(
            ctypes.windll.kernel32.GetConsoleWindow(), 0
        )

    app = QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(False)
    app.setOrganizationName("Copilot")
    app.setApplicationName("Copilot")

    window = WeiboWindow()
    window.show_custom_size(width=500)
    window.hide()

    tray_icon = create_tray_icon(app, window)
    threading.Thread(
        target=listen_copilot_key, args=(window,), daemon=True
    ).start()

    sys.exit(app.exec())
