import sys


def main():
    if "--selftest" in sys.argv:
        from app.selftest import run
        run()
        return
    from PySide6.QtWidgets import QApplication

    from app.gui.main_window import MainWindow
    from app.gui.style import apply as apply_style

    app = QApplication(sys.argv)
    app.setApplicationName("Генератор дипломов")
    apply_style(app)
    win = MainWindow()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
