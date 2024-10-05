

# import aspose.words as aw
# import subprocess


# def test():
#     # Convert .docx to .html
#   doc = aw.Document("osman.docx")
#   doc.save("osmanramadan.html")

# # Convert .html to .pdf using wkhtmltopdf
#   try:
#     subprocess.run(['wkhtmltopdf', '--enable-local-file-access', 'osmanramadan.html', 'lol.pdf'], check=True)
#     print("Successfully converted HTML to PDF.")
#   except subprocess.CalledProcessError as e:
#     print(f"An error occurred during PDF conversion: {e}")



              #  d = QMessageBox(parent=self.windowCreating)  # Set the parent to self.windowCreating
              #  d.setWindowTitle("فشل")  # Set the title for the warning message box
              #  d.setText("لم يتم اختيار اسما لقاعدة البيانات")  # Set the warning message text
              #  d.setIcon(QMessageBox.Icon.Warning)  # Set the icon to Warning
              #  d.exec()  # Execute the dialog to show it
            
            # desktopPath = QStandardPaths.writableLocation(QStandardPaths.StandardLocation.DesktopLocation)
            # database_name, ok = QInputDialog.getText(self.windowCreating, "اكتب قاعدة البيانات", "اكتب اسم قاعدة البيانات:")

            # if not  database_name:
                # return
            
            # filePath = QFileDialog.getExistingDirectory(self.windowCreating, "اختار مسارا", desktopPath)
            # if len(filePath) > 0:
            #   database_name =f"{database_name}.db"
            #   destination_path = os.path.join(filePath, database_name)
            # else:
            #     return

            # if os.path.exists(destination_path):
            #  confirm = QMessageBox.question(
            #     self,
            #     "تنبيه",
            #     f"الملف موجود بالفعل هل تريد اعادة انشائه؟",
            #     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            #  )
            #  if confirm == QMessageBox.StandardButton.No:



# # Import Required libraries
# from PyQt6.QtWidgets import QGridLayout, QHBoxLayout, QApplication, QListWidgetItem, QListWidget, QSplitter, QWidget, QVBoxLayout, QPushButton, QLineEdit, QLabel, QMessageBox, QItemDelegate, QTextEdit, QFrame, QFileDialog, QScrollArea, QMainWindow, QTableWidget, QTableWidgetItem, QRadioButton, QProgressBar, QInputDialog
# from PyQt6.QtGui import QIcon, QFont, QPixmap
# from PyQt6.QtCore import Qt, pyqtSignal
# import sys
# import os
# from docx2pdf import convert

# # Set desktop path
# desktopPath = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')


# # Create a new class for the new window
# class NewWindow(QWidget):
#     def __init__(self, parentWindow):
#         super().__init__()

#         self.parentWindow = parentWindow  # Reference to the parent window

#         self.setWindowTitle("New Page")
#         self.setGeometry(300, 300, 600, 400)  # Set window size and position

#         # Add a simple label
#         label = QLabel("This is the new window.", self)
#         label.setFont(QFont('Arial', 20))
#         label.move(50, 50)

#         # Add the "Export as PDF" button
#         self.exportPdfButton = QPushButton("تصدير PDF", self)
#         self.exportPdfButton.move(50, 100)  # Position the button
#         self.exportPdfButton.clicked.connect(self.exportPdf)

#     def exportPdf(self):
#         # Call the export function from the parent window
#         self.parentWindow.exportSummaryScreenAsPdf(fromWhere="Pdf")


# # Main application window
# class Test(QMainWindow):
#     def __init__(self):
#         super().__init__()

#         # Initialize your UI components here
#         self.initUI()

#     def initUI(self):
#         # Button to open the new window
#         self.pdfExport = QPushButton("فتح الصفحة الجديدة", self)
#         self.pdfExport.clicked.connect(lambda: self.openNewWindow())
#         self.setCentralWidget(self.pdfExport)  # For simplicity, just adding the button as the central widget

#     def exportSummaryScreenAsPdf(self,fromWhere="Pdf"):
#         # Open a file dialog to select the DOCX file
#         file, _ = QFileDialog.getOpenFileName(self, "اختر ملفا", desktopPath, filter="Database File (*.docx)")
#         # Check if a file was selected
#         if file:
#             try:
#                 # Get the name and folder of the selected file
#                 folder = os.path.dirname(file)
#                 nameFile = os.path.basename(file).replace('.docx', '')
#                 pdfname = os.path.join(folder, f"{nameFile}.pdf")

#                 convert(file, pdfname)

#                 # # Show success message
#                 # success_msg = QMessageBox(self)
#                 # success_msg.setWindowTitle("نجاح")
#                 # success_msg.setText("تم التصدير إلى PDF بنجاح")
#                 # success_msg.setIcon(QMessageBox.Icon.Information)
#                 # success_msg.exec_()

#             except Exception as e:
#                 # Show error message
#                 # pass
#                 error_msg = QMessageBox(parent=self)
#                 error_msg.setWindowTitle("خطأ")
#                 error_msg.setText(f"حدث خطأ: {str(e)}")
#                 error_msg.setIcon(QMessageBox.Icon.Warning)
#                 error_msg.exec_()

#     def openNewWindow(self):
#         # Create an instance of the NewWindow and pass the current window as a reference
#         self.newWindow = NewWindow(self)
#         self.newWindow.show()


# if __name__ == "__main__":
#     app = QApplication(sys.argv)
#     window = Test()
#     window.show()
#     sys.exit(app.exec())


