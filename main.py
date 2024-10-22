from PySide6.QtWidgets import QApplication, QSplashScreen
from PySide6.QtCore import QTimer, Qt
from PySide6.QtGui import QPixmap, QIcon
import sys

# Define the LoadingScreen class here to avoid circular imports
class LoadingScreen(QSplashScreen):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        # Customize the loading screen appearance
        self.setWindowFlag(Qt.FramelessWindowHint)
        self.setPixmap(QPixmap("bg.png").scaled(500,500, Qt.KeepAspectRatio))
        
def start_application():
    from enigmapdf import EnigmaPdf 

    mainWin = EnigmaPdf()
    mainWin.show()
    loading_screen.finish(mainWin)

if __name__ == "__main__":
    app = QApplication(sys.argv)


    app_icon = QIcon("pdf_icon.ico")
    app.setWindowIcon(app_icon)

    # Show loading screen immediately
    loading_screen = LoadingScreen()
    loading_screen.show()

    # Set a timer to load the main window after 2 seconds
    QTimer.singleShot(1000, start_application)

    sys.exit(app.exec())