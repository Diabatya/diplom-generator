import sys


def main():
    if "--selftest" in sys.argv:
        from app.selftest import run
        run()
        return
    from PySide6.QtWidgets import QApplication

    from app.gui.main_window import MainWindow

    app = QApplication(sys.argv)
    app.setApplicationName("Генератор дипломов")
    win = MainWindow()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
