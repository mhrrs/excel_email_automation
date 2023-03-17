import sys
import cooper_emails as ce
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QLabel, QFileDialog

class App(QWidget):

    def __init__(self):
        super().__init__()
        self.title = 'File Upload Example'
        self.left = 10
        self.top = 10
        self.width = 320
        self.height = 120
        self.initUI()

    def initUI(self):
        self.setWindowTitle(self.title)
        self.setGeometry(self.left, self.top, self.width, self.height)

        self.label = QLabel(self)
        self.label.setText("Select a file to upload:")
        self.label.move(20, 20)

        self.button = QPushButton('Upload', self)
        self.button.setToolTip('Click to upload a file')
        self.button.move(20, 60)
        self.button.clicked.connect(self.openFileNameDialog)

        self.show()

    def openFileNameDialog(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        fileName, _ = QFileDialog.getOpenFileName(self,"QFileDialog.getOpenFileName()", "","All Files (*);;Python Files (*.py)", options=options)
        if fileName:
            self.file_path = fileName
            print(fileName)

    def parseFile(self):
        self.openFileNameDialog()
        with open(self.file_path, 'r') as f:
            file_contents = f.read()
            print(file_contents)


if __name__ == '__main__':

    # init pyqt5 application
    app = QApplication(sys.argv)
    ex = App()

    # open file
    ex.openFileNameDialog()

    # Parse selected file
    ce.parseExcelSheet(ex.file_path)
    
    # send emails
    
    sys.exit(app.exec_())
