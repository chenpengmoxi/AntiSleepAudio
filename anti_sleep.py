import sys
import os
import time
import threading
import json
import winsound
from pathlib import Path
from PyQt6.QtWidgets import (
    QApplication, QSystemTrayIcon, QMenu, QWidget, 
    QFileDialog, QVBoxLayout, QHBoxLayout, QLabel, 
    QPushButton, QLineEdit, QCheckBox, QMessageBox,
    QGroupBox, QGridLayout
)
from PyQt6.QtGui import QIcon, QAction, QFont, QPalette, QColor
from PyQt6.QtCore import Qt, QTimer, pyqtSignal, QObject
from PyQt6.QtMultimedia import QSoundEffect

# 配置路径
CONFIG_DIR = Path(os.environ.get('APPDATA', '')) / 'AntiSleepAudio'
CONFIG_FILE = CONFIG_DIR / 'config.json'
DEFAULT_AUDIO = CONFIG_DIR / 'silent_pulse.wav'

class SignalBridge(QObject):
    """用于线程间通信的信号桥"""
    show_settings = pyqtSignal()
    show_message = pyqtSignal(str, str)

class AudioKeeper:
    """核心防休眠逻辑"""
    
    def __init__(self):
        self.running = False
        self.thread = None
        self.custom_audio_path = None
        self.use_custom_audio = False
        self.interval = 30  # 默认30秒播放一次
        self.volume = 1  # 音量 0-1
        
    def load_config(self, config):
        """从配置加载设置"""
        self.custom_audio_path = config.get('audio_path', '')
        self.use_custom_audio = config.get('use_custom', False)
        self.interval = config.get('interval', 30)
        self.volume = config.get('volume', 1)
        
    def get_config(self):
        """导出当前配置"""
        return {
            'audio_path': self.custom_audio_path,
            'use_custom': self.use_custom_audio,
            'interval': self.interval,
            'volume': self.volume
        }
    
    def _create_silent_wave(self):
        """创建无声的波形文件（如果自定义音频不存在）"""
        if not DEFAULT_AUDIO.exists():
            try:
                import wave
                import struct
                
                # 创建1秒的无声wav
                sample_rate = 44100
                duration = 1
                num_samples = sample_rate * duration
                
                with wave.open(str(DEFAULT_AUDIO), 'w') as wav_file:
                    wav_file.setnchannels(2)  # 立体声
                    wav_file.setsampwidth(2)   # 16位
                    wav_file.setframerate(sample_rate)
                    
                    # 写入静音数据（极小的声音，防止完全静音被优化掉）
                    for _ in range(num_samples):
                        # 极小的振幅，人耳听不到但声卡保持活跃
                        value = int(0.001 * 32767.0)
                        packed_value = struct.pack('h', value)
                        wav_file.writeframes(packed_value)
                        wav_file.writeframes(packed_value)
            except Exception as e:
                print(f"创建默认音频失败: {e}")
                
    def _play_audio(self):
        """播放音频脉冲"""
        try:
            if self.use_custom_audio and self.custom_audio_path and os.path.exists(self.custom_audio_path):
                # 播放用户选择的音频
                if self.custom_audio_path.lower().endswith('.wav'):
                    winsound.PlaySound(self.custom_audio_path, winsound.SND_FILENAME | winsound.SND_ASYNC)
                else:
                    # 对mp3使用QSoundEffect（需要Qt6）
                    sound = QSoundEffect()
                    sound.setSource(f"file:///{self.custom_audio_path.replace(os.sep, '/')}")
                    sound.setVolume(self.volume)
                    sound.play()
            else:
                # 使用默认无声脉冲
                if DEFAULT_AUDIO.exists():
                    winsound.PlaySound(str(DEFAULT_AUDIO), winsound.SND_FILENAME | winsound.SND_ASYNC)
                else:
                    # 备用方案：系统默认提示音（极短）
                    winsound.MessageBeep(winsound.MB_OK)
        except Exception as e:
            print(f"播放音频失败: {e}")
    
    def _loop(self):
        """主循环"""
        self._create_silent_wave()
        while self.running:
            self._play_audio()
            time.sleep(self.interval)
    
    def start(self):
        """启动防休眠"""
        if not self.running:
            self.running = True
            self.thread = threading.Thread(target=self._loop, daemon=True)
            self.thread.start()
            
    def stop(self):
        """停止防休眠"""
        self.running = False
        if self.thread:
            self.thread.join(timeout=1)

class SettingsDialog(QWidget):
    """设置窗口"""
    
    def __init__(self, keeper, parent=None):
        super().__init__(parent)
        self.keeper = keeper
        self.setWindowTitle("防休眠音频设置")
        self.setFixedSize(500, 350)
        
        # 设置深色主题
        self.setStyleSheet("""
            QWidget {
                background-color: #2b2b2b;
                color: #ffffff;
                font-family: "Microsoft YaHei", "Segoe UI", sans-serif;
            }
            QGroupBox {
                border: 2px solid #3d3d3d;
                border-radius: 8px;
                margin-top: 10px;
                padding-top: 10px;
                font-weight: bold;
                color: #4fc3f7;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px;
            }
            QPushButton {
                background-color: #0d7377;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 4px;
                font-weight: 500;
            }
            QPushButton:hover {
                background-color: #14a085;
            }
            QPushButton:pressed {
                background-color: #0a5c5f;
            }
            QLineEdit {
                background-color: #3d3d3d;
                border: 1px solid #555;
                border-radius: 4px;
                padding: 6px;
                color: white;
            }
            QCheckBox {
                spacing: 8px;
            }
            QCheckBox::indicator {
                width: 18px;
                height: 18px;
                border-radius: 3px;
                border: 2px solid #555;
            }
            QCheckBox::indicator:checked {
                background-color: #0d7377;
                border-color: #0d7377;
            }
        """)
        
        layout = QVBoxLayout()
        layout.setSpacing(15)
        layout.setContentsMargins(20, 20, 20, 20)
        
        # 音频源设置组
        audio_group = QGroupBox("音频源设置")
        audio_layout = QGridLayout()
        audio_layout.setSpacing(10)
        
        self.use_custom_cb = QCheckBox("使用自定义音频文件")
        self.use_custom_cb.setChecked(keeper.use_custom_audio)
        audio_layout.addWidget(self.use_custom_cb, 0, 0, 1, 3)
        
        self.path_edit = QLineEdit(keeper.custom_audio_path or "")
        self.path_edit.setPlaceholderText("选择 .mp3 或 .wav 文件...")
        self.path_edit.setEnabled(keeper.use_custom_audio)
        audio_layout.addWidget(self.path_edit, 1, 0, 1, 2)
        
        browse_btn = QPushButton("浏览...")
        browse_btn.clicked.connect(self.browse_file)
        browse_btn.setEnabled(keeper.use_custom_audio)
        audio_layout.addWidget(browse_btn, 1, 2)
        
        # 连接复选框状态
        self.use_custom_cb.toggled.connect(self.path_edit.setEnabled)
        self.use_custom_cb.toggled.connect(browse_btn.setEnabled)
        
        audio_group.setLayout(audio_layout)
        layout.addWidget(audio_group)
        
        # 高级设置组
        adv_group = QGroupBox("高级设置")
        adv_layout = QGridLayout()
        
        adv_layout.addWidget(QLabel("播放间隔（秒）："), 0, 0)
        self.interval_spin = QLineEdit(str(keeper.interval))
        adv_layout.addWidget(self.interval_spin, 0, 1)
        
        adv_layout.addWidget(QLabel("提示：建议保持30-60秒，过短可能干扰正常使用"), 1, 0, 1, 2)
        adv_layout.addWidget(QLabel("音量：通过系统音量控制"), 2, 0, 1, 2)
        
        adv_group.setLayout(adv_layout)
        layout.addWidget(adv_group)
        
        # 状态信息
        status_group = QGroupBox("当前状态")
        status_layout = QVBoxLayout()
        self.status_label = QLabel("● 防休眠运行中" if keeper.running else "○ 已停止")
        self.status_label.setStyleSheet("color: #4fc3f7; font-size: 12px;")
        status_layout.addWidget(self.status_label)
        status_group.setLayout(status_layout)
        layout.addWidget(status_group)
        
        # 按钮区域
        btn_layout = QHBoxLayout()
        
        test_btn = QPushButton("测试播放")
        test_btn.clicked.connect(self.test_audio)
        test_btn.setStyleSheet("background-color: #5c6bc0;")
        
        save_btn = QPushButton("保存设置")
        save_btn.clicked.connect(self.save_settings)
        save_btn.setStyleSheet("background-color: #26a69a;")
        
        cancel_btn = QPushButton("取消")
        cancel_btn.clicked.connect(self.close)
        cancel_btn.setStyleSheet("background-color: #757575;")
        
        btn_layout.addWidget(test_btn)
        btn_layout.addStretch()
        btn_layout.addWidget(save_btn)
        btn_layout.addWidget(cancel_btn)
        
        layout.addLayout(btn_layout)
        layout.addStretch()
        
        self.setLayout(layout)
        
    def browse_file(self):
        """浏览音频文件"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择音频文件", "",
            "音频文件 (*.mp3 *.wav);;MP3文件 (*.mp3);;WAV文件 (*.wav)"
        )
        if file_path:
            self.path_edit.setText(file_path)
            
    def test_audio(self):
        """测试音频播放"""
        path = self.path_edit.text() if self.use_custom_cb.isChecked() else None
        if path and os.path.exists(path):
            try:
                if path.lower().endswith('.wav'):
                    winsound.PlaySound(path, winsound.SND_FILENAME | winsound.SND_ASYNC)
                else:
                    # 使用系统默认播放器测试mp3
                    os.startfile(path)
                QMessageBox.information(self, "测试", "正在播放测试音频...")
            except Exception as e:
                QMessageBox.warning(self, "错误", f"播放失败: {e}")
        else:
            # 测试默认无声脉冲
            self.keeper._play_audio()
            QMessageBox.information(self, "测试", "已发送无声脉冲（保持声卡活跃）")
            
    def save_settings(self):
        """保存设置"""
        try:
            interval = int(self.interval_spin.text())
            if interval < 5 or interval > 300:
                raise ValueError("间隔必须在5-300秒之间")
                
            self.keeper.use_custom_audio = self.use_custom_cb.isChecked()
            self.keeper.custom_audio_path = self.path_edit.text()
            self.keeper.interval = interval
            
            # 重启服务以应用新间隔
            was_running = self.keeper.running
            if was_running:
                self.keeper.stop()
                time.sleep(0.5)
                
            self.keeper.start()
            
            QMessageBox.information(self, "成功", "设置已保存并生效！")
            self.close()
            
        except ValueError as e:
            QMessageBox.warning(self, "输入错误", str(e))

class TrayApplication(QApplication):
    """主应用程序"""
    
    def __init__(self, argv):
        super().__init__(argv)
        self.setQuitOnLastWindowClosed(False)
        
        # 确保配置目录存在
        CONFIG_DIR.mkdir(parents=True, exist_ok=True)
        
        # 初始化核心组件
        self.keeper = AudioKeeper()
        self.load_config()
        
        # 创建托盘图标
        self.tray_icon = QSystemTrayIcon()
        self.update_icon()
        self.tray_icon.setToolTip("防声卡休眠 - 运行中")
        
        # 创建菜单
        self.create_menu()
        
        # 信号桥
        self.signals = SignalBridge()
        self.signals.show_settings.connect(self.show_settings)
        
        # 启动服务
        self.keeper.start()
        self.tray_icon.show()
        
        # 显示启动通知
        self.tray_icon.showMessage(
            "防声卡休眠已启动",
            "正在阻止声卡自动休眠，右键图标进行设置",
            QSystemTrayIcon.MessageIcon.Information,
            3000
        )
        
    def update_icon(self):
        """更新托盘图标（使用内嵌图标或系统图标）"""
        # 使用系统标准图标作为fallback
        style = self.style()
        icon = style.standardIcon(self.style().StandardPixmap.SP_ComputerIcon)
        self.tray_icon.setIcon(icon)
        
    def create_menu(self):
        """创建右键菜单"""
        self.menu = QMenu()
        self.menu.setStyleSheet("""
            QMenu {
                background-color: #2b2b2b;
                color: white;
                border: 1px solid #3d3d3d;
                padding: 5px;
            }
            QMenu::item {
                padding: 8px 25px;
                border-radius: 4px;
            }
            QMenu::item:selected {
                background-color: #0d7377;
            }
            QMenu::separator {
                height: 1px;
                background-color: #3d3d3d;
                margin: 5px 0;
            }
        """)
        
        # 状态显示
        status_action = QAction("● 防休眠运行中", self)
        status_action.setEnabled(False)
        self.menu.addAction(status_action)
        self.menu.addSeparator()
        
        # 设置选项
        settings_action = QAction("⚙ 打开设置...", self)
        settings_action.triggered.connect(self.show_settings)
        self.menu.addAction(settings_action)
        
        # 测试音频
        test_action = QAction("▶ 测试音频", self)
        test_action.triggered.connect(self.keeper._play_audio)
        self.menu.addAction(test_action)
        
        self.menu.addSeparator()
        
        # 开机自启管理
        self.autostart_action = QAction("✓ 开机自启动", self)
        self.autostart_action.setCheckable(True)
        self.autostart_action.setChecked(self.check_autostart())
        self.autostart_action.triggered.connect(self.toggle_autostart)
        self.menu.addAction(self.autostart_action)
        
        self.menu.addSeparator()
        
        # 退出
        quit_action = QAction("✕ 退出", self)
        quit_action.triggered.connect(self.quit_app)
        self.menu.addAction(quit_action)
        
        self.tray_icon.setContextMenu(self.menu)
        self.tray_icon.activated.connect(self.on_tray_activated)
        
    def on_tray_activated(self, reason):
        """处理托盘图标激活事件"""
        if reason == QSystemTrayIcon.ActivationReason.DoubleClick:
            self.show_settings()
            
    def show_settings(self):
        """显示设置窗口"""
        self.settings_dialog = SettingsDialog(self.keeper)
        self.settings_dialog.show()
        self.settings_dialog.activateWindow()
        
    def load_config(self):
        """加载配置文件"""
        if CONFIG_FILE.exists():
            try:
                with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    self.keeper.load_config(config)
            except Exception as e:
                print(f"加载配置失败: {e}")
                
    def save_config(self):
        """保存配置文件"""
        try:
            with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(self.keeper.get_config(), f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"保存配置失败: {e}")
            
    def get_startup_path(self):
        """获取启动目录路径"""
        startup_dir = Path(os.environ.get('APPDATA', '')) / 'Microsoft' / 'Windows' / 'Start Menu' / 'Programs' / 'Startup'
        return startup_dir / 'AntiSleepAudio.lnk'
        
    def check_autostart(self):
        """检查是否已设置开机自启"""
        return self.get_startup_path().exists()
        
    def toggle_autostart(self, enabled):
        """切换开机自启状态"""
        shortcut_path = self.get_startup_path()
        exe_path = sys.executable
        
        if enabled:
            try:
                import winshell
                from win32com.client import Dispatch
                
                shell = Dispatch('WScript.Shell')
                shortcut = shell.CreateShortCut(str(shortcut_path))
                shortcut.Targetpath = exe_path
                shortcut.WorkingDirectory = os.path.dirname(exe_path)
                shortcut.IconLocation = exe_path
                shortcut.save()
                
                self.tray_icon.showMessage("设置成功", "已添加到开机启动项", QSystemTrayIcon.MessageIcon.Information)
            except ImportError:
                # 如果没有winshell，创建简单的bat文件作为fallback
                bat_path = shortcut_path.with_suffix('.bat')
                with open(bat_path, 'w') as f:
                    f.write(f'start "" "{exe_path}"')
                self.tray_icon.showMessage("设置成功", "已创建启动脚本（简易模式）", QSystemTrayIcon.MessageIcon.Information)
            except Exception as e:
                QMessageBox.warning(None, "错误", f"设置开机启动失败: {e}")
                self.autostart_action.setChecked(False)
        else:
            try:
                if shortcut_path.exists():
                    shortcut_path.unlink()
                # 也检查bat文件
                bat_path = shortcut_path.with_suffix('.bat')
                if bat_path.exists():
                    bat_path.unlink()
                self.tray_icon.showMessage("设置成功", "已移除开机启动项", QSystemTrayIcon.MessageIcon.Information)
            except Exception as e:
                QMessageBox.warning(None, "错误", f"移除开机启动失败: {e}")
                
    def quit_app(self):
        """退出应用"""
        self.save_config()
        self.keeper.stop()
        self.tray_icon.hide()
        self.quit()

def main():
    # 高DPI支持
    if hasattr(Qt, 'AA_EnableHighDpiScaling'):
        QApplication.setAttribute(Qt.ApplicationAttribute.AA_EnableHighDpiScaling, True)
    if hasattr(Qt, 'AA_UseHighDpiPixmaps'):
        QApplication.setAttribute(Qt.ApplicationAttribute.AA_UseHighDpiPixmaps, True)
        
    app = TrayApplication(sys.argv)
    sys.exit(app.exec())

if __name__ == '__main__':
main()
