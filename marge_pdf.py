import sys
import ctypes
from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QVBoxLayout, QPushButton, QFileDialog
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QIcon, QPixmap
from PyQt5.QtWinExtras import QtWin
from PyPDF2 import PdfMerger
import base64
from io import BytesIO

# מזהה ייחודי לאפליקציה
myappid = 'MIT.LEARN_PYQT.pdfmergeapp.1.0'

# מחרוזת Base64 של האייקון 
icon_base64 = "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"

class PDFMergeWindow(QWidget):
    def __init__(self):
        super().__init__()

        # הגדרת החלון
        self.setWindowTitle("מיזוג קבצי PDF")
        self.setWindowIcon(self.load_icon_from_base64(icon_base64))
        self.setGeometry(100, 100, 400, 300)

        # הגדרת האייקון לשורת המשימות
        if sys.platform == 'win32':
            QtWin.setCurrentProcessExplicitAppUserModelID(myappid)

        # יצירת תווית להנחיה
        self.label = QLabel("גרור ושחרר או בחר קבצי PDF למיזוג", self)
        self.label.setAlignment(Qt.AlignCenter)
        self.label.setStyleSheet("border: 2px dashed gray; font-size: 16px; padding: 40px;")

        # יצירת כפתור להוספת קבצים דרך סייר הקבצים
        self.add_files_button = QPushButton('בחר קבצי PDF', self)
        self.add_files_button.clicked.connect(self.open_file_dialog)

        # יצירת כפתור לביצוע המיזוג
        self.merge_button = QPushButton('בחירת שם ונתיב לקובץ הממוזג', self)
        self.merge_button.setEnabled(False)
        self.merge_button.clicked.connect(self.merge_pdfs)

        # כפתורים שיוצגו לאחר המיזוג
        self.merge_new_button = QPushButton('מיזוג קבצים חדשים', self)
        self.merge_new_button.setVisible(False)
        self.merge_new_button.clicked.connect(self.reset_for_new_merge)

        self.merge_elsewhere_button = QPushButton('מזג למיקום אחר', self)
        self.merge_elsewhere_button.setVisible(False)
        self.merge_elsewhere_button.clicked.connect(self.merge_to_another_location)

        self.close_button = QPushButton('סגור את התוכנה', self)
        self.close_button.setVisible(False)
        self.close_button.clicked.connect(self.close_program)

        # פריסת אלמנטים
        layout = QVBoxLayout()
        layout.addWidget(self.label)
        layout.addWidget(self.add_files_button)
        layout.addWidget(self.merge_button)
        layout.addWidget(self.merge_elsewhere_button)
        layout.addWidget(self.merge_new_button)
        layout.addWidget(self.close_button)
        self.setLayout(layout)

        # משתנה לאחסון נתיבי קבצי PDF
        self.pdf_files = []

        # קביעת שהחלון יקבל גרירה ושחרור
        self.setAcceptDrops(True)

    # פונקציה לפתיחת סייר קבצים לבחירת קבצים
    def open_file_dialog(self):
        files, _ = QFileDialog.getOpenFileNames(self, "בחר קבצי PDF", "", "PDF Files (*.pdf)")
        if files:
            self.pdf_files.extend(files)
            self.label.setText(f"{len(self.pdf_files)} קבצי PDF נוספו.")
            self.merge_button.setEnabled(True)

    # פונקציה שמופעלת כשגוררים קובץ לחלון
    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()

    # פונקציה שמופעלת כשמשחררים את הקבצים בחלון
    def dropEvent(self, event):
        if event.mimeData().hasUrls():
            for url in event.mimeData().urls():
                file_path = url.toLocalFile()
                if file_path.endswith('.pdf'):
                    self.pdf_files.append(file_path)
            self.label.setText(f"{len(self.pdf_files)} קבצי PDF נוספו.")
            self.merge_button.setEnabled(True)

    # פונקציה למיזוג קבצי PDF
    def merge_pdfs(self):
        if not self.pdf_files:
            return

        output_file, _ = QFileDialog.getSaveFileName(self, "שמור קובץ ממוזג", "", "PDF Files (*.pdf)")

        if output_file:
            merger = PdfMerger()

            for pdf in self.pdf_files:
                merger.append(pdf)

            with open(output_file, 'wb') as f:
                merger.write(f)

            self.label.setText(f"הקבצים מוזגו בהצלחה!\n{output_file}")
            self.show_post_merge_buttons()

    # פונקציה להצגת כפתורים אחרי המיזוג
    def show_post_merge_buttons(self):
        self.add_files_button.setEnabled(False)
        self.merge_button.setEnabled(False)
        self.merge_new_button.setVisible(True)
        self.merge_elsewhere_button.setVisible(True)
        self.close_button.setVisible(True)
        self.close_button.setFocus()

    # פונקציה לסגירת התוכנה
    def close_program(self):
        self.close()

    # פונקציה למיזוג למיקום אחר
    def merge_to_another_location(self):
        if not self.pdf_files:
            return

        output_file, _ = QFileDialog.getSaveFileName(self, "שמור את הקבצים בשם אחר או במיקום אחר", "", "PDF Files (*.pdf)")

        if output_file:
            merger = PdfMerger()

            for pdf in self.pdf_files:
                merger.append(pdf)

            with open(output_file, 'wb') as f:
                merger.write(f)

            self.label.setText(f"מוזג בהצלחה גם ל:\n{output_file}")

    # פונקציה לאיפוס עבור מיזוג קבצים חדשים
    def reset_for_new_merge(self):
        self.pdf_files = []
        self.label.setText("גרור ושחרר קבצי PDF כאן")
        self.merge_button.setEnabled(False)
        self.merge_new_button.setVisible(False)
        self.merge_elsewhere_button.setVisible(False)
        self.close_button.setVisible(False)

    # פונקציה לטעינת אייקון ממחרוזת Base64
    def load_icon_from_base64(self, base64_string):
        pixmap = QPixmap()
        pixmap.loadFromData(base64.b64decode(base64_string))
        return QIcon(pixmap)

# הפונקציה הראשית שמריצה את היישום
def main():
    if sys.platform == 'win32':
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)

    app = QApplication(sys.argv)
    window = PDFMergeWindow()
    window.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()
