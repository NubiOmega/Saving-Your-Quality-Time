# Utilities\dragndrop_files_func.py
from PyQt6 import QtCore, QtWidgets

def dragEnterEvent(self, event):
    if event.mimeData().hasUrls():
        event.accept()
    else:
        event.ignore()

def dragMoveEvent(self, event):
    if event.mimeData().hasUrls():
        event.accept()
    else:
        event.ignore()

def dropEvent(self, event):
    if event.mimeData().hasUrls():
        event.setDropAction(QtCore.Qt.DropAction.CopyAction)
        event.accept()
        urls = event.mimeData().urls()
        for url in urls:
            if url.isLocalFile() and url.toLocalFile().endswith('.xlsx'):
                file_path = url.toLocalFile()
                if not self.fileExists(file_path):
                    # Tambahkan item ke listItems_filesSource
                    item = QtWidgets.QListWidgetItem(file_path)
                    item.setToolTip(file_path)
                    self.ui.listItems_filesSource.addItem(item)
                else:
                    QtWidgets.QMessageBox.warning(self, "File Duplikat", f"File '{file_path}' sudah ada dalam daftar.")
    else:
        event.ignore()

def fileExists(self, file_path):
    for row in range(self.ui.listItems_filesSource.count()):
        item = self.ui.listItems_filesSource.item(row)
        if item.toolTip() == file_path:
            return True
    return False
