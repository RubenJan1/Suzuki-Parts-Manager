# utils/theme.py

from PySide6.QtGui import QPalette


def is_dark(widget):
    bg = widget.palette().color(QPalette.Window)
    return bg.lightness() < 128


def apply_theme(widget):
    if is_dark(widget):
        style = """
        QWidget {
            background-color: #111827;
            color: #E5E7EB;
            font-size: 10pt;
        }

        QLabel {
            color: #E5E7EB;
        }

        QLineEdit, QTextEdit, QComboBox, QDoubleSpinBox {
            background-color: #0B1220;
            border: 1px solid #273244;
            padding: 6px;
            border-radius: 6px;
        }

        QPushButton {
            padding: 6px 12px;
            border-radius: 6px;
        }

        QPushButton#primary {
            background-color: #2563EB;
            color: white;
        }

        QPushButton#danger {
            background-color: #DC2626;
            color: white;
        }

        QPushButton#secondary {
            background-color: #1F2937;
            border: 1px solid #273244;
        }
        """
    else:
        style = """
        QWidget {
            background-color: #F3F4F6;
            color: #111827;
            font-size: 10pt;
        }

        QLabel {
            color: #111827;
        }

        QLineEdit, QTextEdit, QComboBox, QDoubleSpinBox {
            background-color: white;
            border: 1px solid #D1D5DB;
            padding: 6px;
            border-radius: 6px;
        }

        QPushButton {
            padding: 6px 12px;
            border-radius: 6px;
        }

        QPushButton#primary {
            background-color: #2563EB;
            color: white;
        }

        QPushButton#danger {
            background-color: #DC2626;
            color: white;
        }

        QPushButton#secondary {
            background-color: white;
            border: 1px solid #D1D5DB;
        }
        """

    widget.setStyleSheet(style)