from PySide6.QtGui import QPalette

def is_dark_mode(widget):
    bg = widget.palette().color(QPalette.Window)
    return bg.lightness() < 128


def apply_theme(widget):
    dark = is_dark_mode(widget)

    if dark:
        style = """
        QWidget {
            background-color: #111827;
            color: #E5E7EB;
            font-size: 10pt;
        }

        QLabel { color: #E5E7EB; }

        QLineEdit, QTextEdit, QComboBox {
            background-color: #0B1220;
            color: #E5E7EB;
            border: 1px solid #273244;
            padding: 6px;
            border-radius: 4px;
        }

        QPushButton {
            padding: 6px 12px;
            border-radius: 4px;
            min-height: 30px;
        }

        QPushButton#primary {
            background-color: #2563EB;
            color: white;
        }

        QPushButton#secondary {
            background-color: #1F2937;
            border: 1px solid #273244;
        }

        QPushButton#danger {
            background-color: #DC2626;
            color: white;
        }

        QTableWidget {
            background-color: #0B1220;
            border: 1px solid #273244;
        }

        QHeaderView::section {
            background-color: #0E172A;
            padding: 6px;
        }
        """
    else:
        style = """
        QWidget {
            background-color: #F3F4F6;
            color: #111827;
            font-size: 10pt;
        }

        QLineEdit, QTextEdit, QComboBox {
            background-color: white;
            border: 1px solid #D1D5DB;
            padding: 6px;
            border-radius: 4px;
        }

        QPushButton#primary {
            background-color: #2563EB;
            color: white;
        }

        QPushButton#secondary {
            background-color: white;
            border: 1px solid #D1D5DB;
        }

        QPushButton#danger {
            background-color: #DC2626;
            color: white;
        }
        """

    widget.setStyleSheet(style)