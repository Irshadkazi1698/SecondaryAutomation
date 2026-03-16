import sys
from PyQt5.QtWidgets import *
from PyQt5.QtGui import QIntValidator,QDoubleValidator,QFont
from PyQt5.QtCore import Qt,QThread,pyqtSignal
from PyQt5 import QtWidgets, QtGui
import os
from CountsFileCleaning import CountsCleaning


class WorkThread(QThread):
    finished = pyqtSignal(str)

    def __init__(self, file_path,colnum):
        super().__init__()
        self.file_path = file_path
        self.colnum = colnum

    def run(self):
        self.finished.emit("Thread started")
        CountsCleaning.LoadingFileAndCleaning(self.file_path,self.colnum)
        CountsCleaning.CopyingFiletoInputFolder()
        self.finished.emit("Processing completed")


class BannerCleaningGUI(QWidget):
    def __init__(self):
      super().__init__()

      self.setWindowTitle("Banner Cleaning Tool")
      self.setGeometry(400, 400, 450, 450)
      self.setWindowIcon(QtGui.QIcon('icon.png'))

      self.BannerFileLabel = QLabel("Banner File Location :",self)
      self.BannerFileLabel.move(30,50)
      self.BannerEdit = QLineEdit(self)
      self.BannerEdit.setFixedWidth(150)
      self.BannerEdit.move(170,50)
      self.BannerButton = QPushButton("Browse",self)
      self.BannerButton.clicked.connect(self.browse_banner_file)
      self.BannerButton.move(330,50)

      self.BannerTextArea = QTextEdit(self)
      self.BannerTextArea.setFixedWidth(350)
      self.BannerTextArea.move(50,100)

      self.BannerButton = QPushButton("Generate",self)
      self.BannerButton.clicked.connect(self.generate_banner)
      self.BannerButton.move(170,350)

    def browse_banner_file(self):
        file_dialog = QFileDialog()
        banner_file, _ = file_dialog.getOpenFileName(self, "Select Banner File", "", "CSV Files (*.csv);;All Files (*)")
        if banner_file:
            self.BannerEdit.setText(banner_file)

    def generate_banner(self):
        try:
            file_path = self.BannerEdit.text()
            print(f"Selected file: {file_path}")
            self.thread = WorkThread(file_path,5)
            self.thread.finished.connect(self.on_thread_finished)
            self.thread.start()
        except Exception as e:
            self.BannerTextArea.append(f"Error: {str(e)}")
    
    def on_thread_finished(self, message):
        self.BannerTextArea.append(message)

if __name__ == '__main__':
   app = QApplication(sys.argv)
   window = BannerCleaningGUI()
   window.show()
   sys.exit(app.exec())