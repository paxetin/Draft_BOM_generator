import sys
from PyQt5.QtWidgets import (
    QApplication, 
    QPushButton, 
    QVBoxLayout, 
    QHBoxLayout,
    QFileDialog,
    QLineEdit,
    QTextEdit,
    QLabel,
    QWidget, 
    QProgressBar, 
    QDialog
    )
from PyQt5.QtCore import QThread, pyqtSignal
from PyQt5.QtGui import QIcon
from functions import *


class Worker(QThread):
    """
    Worker thread for running tasks in the background.
    """
    progress = pyqtSignal(int)
    finished = pyqtSignal()

    def __init__(self, filePath, fileName, parent=None):
        super(Worker, self).__init__(parent)
        self.filePath = filePath
        self.fileName = fileName


    def run(self):
        """
        Attributes:
            target: The original BOM file path as input.
            designated_name: The file name for generated BOM.
        """

        input_bom = xlsx_loader(self.filePath)

        try:
            parts = input_bom.tolist()
        except:
            raise KeyError('"Part" is not found in axis.')

        searcher = Searcher(by = 'keyword')
        generator = BOMGenerator()

        with pd.ExcelWriter(f"./{self.fileName}.xlsx") as writer:

            # init to keep the draft BOM at the first page.
            generator.BOM.to_excel(writer, sheet_name='draft BOM')
                
            for idx, part in enumerate(parts):
                data = searcher.get(desc_parser(part))
                df = pd.json_normalize(data['SearchResults']['Parts'])
                df.to_excel(writer, sheet_name=str(idx+1), index=False)

                generator.append(df, part)
                self.progress.emit(self.progress_idx(len(parts), idx))

            # update the draft BOM per row.
            generator.BOM.to_excel(writer, sheet_name='draft BOM', index=False)

        self.finished.emit()

    @staticmethod
    def progress_idx(parts_len, idx):
        return int( idx / (parts_len - 1) * 100 )


class CustomStream(object):
    """
    Redirection of sys.stdout to this custom stream.
    """
    def __init__(self, text_widget):
        self.text_widget = text_widget

    def write(self, message):
        self.text_widget.append(message)  # Append message to the QTextEdit

    def flush(self):
        pass

class ProgressDialog(QDialog):
    """
    Separate dialog for progress bar.
    """
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle('Progress')
        self.setGeometry(300, 300, 250, 50)
        self.progress_bar = QProgressBar(self)
        layout = QVBoxLayout()
        layout.addWidget(self.progress_bar)
        self.setLayout(layout)

    def setProgress(self, value):
        self.progress_bar.setValue(value)

class AppWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

        self.setWindowIcon(QIcon('./img/logo.png'))
        self.setWindowTitle("Draft BOM Generator")


    def initUI(self):
        mainLayout = QVBoxLayout()
        pathLayout = QHBoxLayout() 
        inputLayout = QHBoxLayout()

        # File path input
        self.pathName_text = QLabel('Target File Path: ')
        pathLayout.addWidget(self.pathName_text)
        self.filePathInput = QLineEdit(self)
        pathLayout.addWidget(self.filePathInput)

        # Button for selecting file path
        self.selectPathButton = QPushButton('Browse', self)
        self.selectPathButton.clicked.connect(self.select_path)
        pathLayout.addWidget(self.selectPathButton)

        mainLayout.addLayout(pathLayout)  # Add the horizontal layout to the main layout

        self.fileName = QLabel('Saved File name: ')
        inputLayout.addWidget(self.fileName)
        self.fileNameInput = QLineEdit(self)
        inputLayout.addWidget(self.fileNameInput)

        mainLayout.addLayout(inputLayout)

        # Button to activate program
        self.activateButton = QPushButton('Generate', self)
        self.activateButton.clicked.connect(self.startTask)
        mainLayout.addWidget(self.activateButton)
    
        self.setLayout(mainLayout)
        self.setWindowTitle('Draft BOM Search/Generator')
        self.setGeometry(300, 300, 350, 200)
        self.show()

    def select_path(self):
        # Function to handle file path selection
        filePath, _ = QFileDialog.getOpenFileName(self)
        if filePath:
            self.filePathInput.setText(filePath)
        print(f"File path selected: {filePath}")


    def startTask(self):
        self.progress_dialog = ProgressDialog(self)
        filePath = self.filePathInput.text()
        fileName = self.fileNameInput.text()
        self.thread = Worker(filePath, fileName)
        self.thread.progress.connect(self.progress_dialog.setProgress)
        self.thread.finished.connect(self.taskFinished)
        self.progress_dialog.show()
        self.thread.start()

    def taskFinished(self):
        self.progress_dialog.accept()
        print("Task completed!")

def main():
    app = QApplication(sys.argv)
    ex = AppWindow()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()