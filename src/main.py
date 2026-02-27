import sys
from pathlib import Path
from PyQt6.QtWidgets import QApplication
from src.gui.main_window import MainWindow

def main():
    app = QApplication(sys.argv)
    
    # App style setup if needed
    app.setStyle("Fusion")
    
    # QSS Yükleme
    if hasattr(sys, "_MEIPASS"):
        qss_path = Path(sys._MEIPASS) / "src" / "gui" / "style.qss"
    else:
        qss_path = Path(__file__).parent / "gui" / "style.qss"
        
    if qss_path.exists():
        with open(qss_path, "r", encoding="utf-8") as f:
            app.setStyleSheet(f.read())

    window = MainWindow()
    window.show()

    sys.exit(app.exec())

if __name__ == "__main__":
    main()
