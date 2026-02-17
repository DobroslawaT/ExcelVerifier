# main.py
import sys
import os

# Set Qt plugin path BEFORE importing PyQt5
if hasattr(sys, 'frozen'):
    # Running as compiled exe
    os.environ['QT_PLUGIN_PATH'] = os.path.join(sys._MEIPASS, 'PyQt5', 'Qt5', 'plugins')
    os.environ['QT_QPA_PLATFORM_PLUGIN_PATH'] = os.path.join(sys._MEIPASS, 'PyQt5', 'Qt5', 'plugins', 'platforms')
else:
    # Running as script - set to venv path
    venv_path = os.path.dirname(os.path.dirname(sys.executable))
    plugins_path = os.path.join(venv_path, 'Lib', 'site-packages', 'PyQt5', 'Qt5', 'plugins')
    platforms_path = os.path.join(plugins_path, 'platforms')
    os.environ['QT_PLUGIN_PATH'] = plugins_path
    os.environ['QT_QPA_PLATFORM_PLUGIN_PATH'] = platforms_path

from PyQt5.QtWidgets import QApplication
from PyQt5.QtCore import Qt
from ui.main_window import VerifyApp

if __name__ == "__main__":
    # Enable High DPI scaling for better display on different screen sizes
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)
    
    app = QApplication(sys.argv)
    # This now loads the MainWindow with the Menu Bar
    window = VerifyApp() 
    window.showMaximized()
    sys.exit(app.exec_())