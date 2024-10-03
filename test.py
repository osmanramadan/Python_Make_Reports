import aspose.words as aw
import subprocess


def test():
    # Convert .docx to .html
  doc = aw.Document("osman.docx")
  doc.save("osmanramadan.html")

# Convert .html to .pdf using wkhtmltopdf
  try:
    subprocess.run(['wkhtmltopdf', '--enable-local-file-access', 'osmanramadan.html', 'lol.pdf'], check=True)
    print("Successfully converted HTML to PDF.")
  except subprocess.CalledProcessError as e:
    print(f"An error occurred during PDF conversion: {e}")


            #    d = QMessageBox(parent=self.windowCreating)  # Set the parent to self.windowCreating
            #    d.setWindowTitle("فشل")  # Set the title for the warning message box
            #    d.setText("لم يتم اختيار اسما لقاعدة البيانات")  # Set the warning message text
            #    d.setIcon(QMessageBox.Icon.Warning)  # Set the icon to Warning
            #    d.exec()  # Execute the dialog to show it
