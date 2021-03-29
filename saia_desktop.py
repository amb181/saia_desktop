from PyQt5.QtCore import QUrl, Qt, QSize
from PyQt5.QtGui import QIcon, QMouseEvent, QCursor
from PyQt5.QtWebEngineWidgets import QWebEngineView, QWebEngineSettings
from PyQt5.QtWidgets import QApplication, QMainWindow, QSystemTrayIcon, QMenu, QPushButton, QLabel
import sys, os

class SAIA_chat(QMainWindow):

    def __init__(self):
        super(SAIA_chat, self).__init__()
        # Window content
        self.setWindowFlag(Qt.FramelessWindowHint)
        self.setWindowTitle('Sourcing Artificial Intelligence Assistant')
        self.setFixedSize(380, 545)
        self.UiComponents()
        self.show()

    def UiComponents(self):
        self.setWindowIcon(QIcon('C:\\Program Files\\SAIA\\bin\\favicon.ico'))
        self.browser = QWebEngineView()
        self.browser.setUrl( QUrl('https://sourcingai.internal.ericsson.com/home.php'))
        self.browser.page().settings().setAttribute(QWebEngineSettings.ShowScrollBars, False)
        self.setCentralWidget(self.browser)
        # creating Close push button
        self.button = QPushButton("", self)
        self.button.setToolTip('Close')
        # setting geometry of button
        self.button.setGeometry(350, 0, 30, 30)
        # adding action to a button
        self.button.clicked.connect(self.closeEvent)
        # setting image to the button
        self.button.setIcon(QIcon('C:\\Program Files\\SAIA\\bin\\close_icon.png'))
        self.button.setIconSize(QSize(30, 30))
        # dummy label  to allow drag the window around
        self.label = QLabel(' ', self)
        self.label.resize(350, 30)

    def closeEvent(self, event):
        self.hide()
        # Add system tray icon
        self.trayIcon = QSystemTrayIcon(QIcon('C:\\Program Files\\SAIA\\bin\\favicon.ico'))
        self.trayIcon.setToolTip('Sourcing Artificial Intelligence Assistant')
        self.trayIcon.show()
        self.trayIcon.showMessage("SAIA is still running!", "SAIA will continue running in the background." )
        # Function for tray
        self.menu = QMenu()
        # Re-open
        self.reopen_action = self.menu.addAction('Open')        
        self.reopen_action.triggered.connect(self.Re_Open)
        # Exit
        self.exit_action = self.menu.addAction('Exit')          
        self.exit_action.triggered.connect(app.quit)  
        # Add actions to menu          
        self.trayIcon.setContextMenu(self.menu)
        # Action on-click
        self.trayIcon.activated.connect(self.trayIconActivated) 

    def Re_Open(self):      # Open with right click (menu)
        self.show()
        self.raise_()
        self.trayIcon.hide()

    def trayIconActivated(self, reason):    # Open with left click (click on tray icon)
        if reason == self.trayIcon.Trigger:
            self.show()
            self.raise_()
            self.trayIcon.hide()

    def mousePressEvent(self, event):
        if event.button()==Qt.LeftButton:
            self.m_flag=True
            self.m_Position=event.globalPos()-self.pos() #Get the position of the mouse relative to the window
            event.accept()
            self.setCursor(QCursor(Qt.OpenHandCursor))  #Change mouse icon
            
    def mouseMoveEvent(self, QMouseEvent):
        if Qt.LeftButton and self.m_flag:  
            self.move(QMouseEvent.globalPos()-self.m_Position)#Change window position
            QMouseEvent.accept()
            
    def mouseReleaseEvent(self, QMouseEvent):
        self.m_flag=False
        self.setCursor(QCursor(Qt.ArrowCursor))


if __name__ == '__main__':
    app = QApplication(sys.argv)

    window = SAIA_chat()

    sys.exit(app.exec_())