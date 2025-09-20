import requests
from PyQt5.QtCore import QThread, pyqtSignal

class Thread2(QThread):
    data_ready = pyqtSignal(list)

    def __init__(self, url, parent=None):
        super().__init__(parent)
        self.url = url

    def run(self):
        try:
            resp = requests.get(self.url)
            resp.raise_for_status()
            lines = resp.text.strip().split("\n")

            table_data = []
            for line in lines:
                parts = line.strip().split()
                if len(parts) >= 3:
                    code = parts[0]
                    name = parts[1]
                    last_close = parts[2]
                    table_data.append([code, name, last_close])

            self.data_ready.emit(table_data)

        except Exception as e:
            print("Thread2 Error:", e)
            self.data_ready.emit([])