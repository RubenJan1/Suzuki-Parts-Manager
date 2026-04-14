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

        QMainWindow, QTabWidget::pane {
            background-color: #111827;
        }

        QTabBar::tab {
            background-color: #1F2937;
            color: #E5E7EB;
            border: 1px solid #334155;
            padding: 8px 14px;
            min-height: 18px;
            margin-right: 4px;
            border-top-left-radius: 6px;
            border-top-right-radius: 6px;
        }

        QTabBar::tab:selected {
            background-color: #243244;
            color: #FFFFFF;
        }

        QTabBar::tab:hover {
            background-color: #2B3A4F;
        }

        QLabel {
            color: #E5E7EB;
            background: transparent;
        }

        QLineEdit, QTextEdit, QComboBox, QDoubleSpinBox, QTableWidget, QTreeWidget {
            background-color: #1F2937;
            color: #E5E7EB;
            border: 1px solid #334155;
            border-radius: 6px;
            padding: 6px;
        }

        QLineEdit:focus, QTextEdit:focus, QComboBox:focus, QDoubleSpinBox:focus {
            border: 1px solid #5B8DEF;
        }

        QHeaderView::section {
            background-color: #243244;
            color: #E5E7EB;
            border: 1px solid #334155;
            padding: 6px;
            font-weight: bold;
        }

        QGroupBox {
            border: 1px solid #334155;
            border-radius: 8px;
            margin-top: 12px;
            padding-top: 12px;
            background-color: #0F172A;
        }

        QGroupBox::title {
            subcontrol-origin: margin;
            left: 10px;
            padding: 0 6px;
            color: #E5E7EB;
            font-weight: bold;
            background-color: #0F172A;
        }

        QPushButton {
            background-color: #243244;
            color: #E5E7EB;
            border: 1px solid #334155;
            border-radius: 6px;
            padding: 8px 14px;
            min-height: 18px;
        }

        QPushButton:hover {
            background-color: #2B3A4F;
        }

        QPushButton#primary {
            background-color: #3B82F6;
            color: white;
            border: 1px solid #3B82F6;
        }

        QPushButton#primary:hover {
            background-color: #2563EB;
        }

        QPushButton#secondary {
            background-color: #1F2937;
            color: #E5E7EB;
            border: 1px solid #475569;
        }

        QPushButton#secondary:hover {
            background-color: #273449;
        }

        QPushButton#danger {
            background-color: #DC2626;
            color: white;
            border: 1px solid #DC2626;
        }

        QPushButton#danger:hover {
            background-color: #B91C1C;
        }

        QPushButton:disabled {
            background-color: #374151;
            color: #9CA3AF;
            border: 1px solid #4B5563;
        }

        QRadioButton, QCheckBox {
            spacing: 8px;
            color: #E5E7EB;
            background: transparent;
        }

        QRadioButton::indicator, QCheckBox::indicator {
            width: 16px;
            height: 16px;
        }

        QRadioButton::indicator:unchecked {
            border: 2px solid #94A3B8;
            border-radius: 8px;
            background: #0F172A;
        }

        QRadioButton::indicator:checked {
            border: 2px solid #60A5FA;
            border-radius: 8px;
            background: #60A5FA;
        }

        QCheckBox::indicator:unchecked {
            border: 2px solid #94A3B8;
            border-radius: 3px;
            background: #0F172A;
        }

        QCheckBox::indicator:checked {
            border: 2px solid #60A5FA;
            border-radius: 3px;
            background: #60A5FA;
        }
        """
    else:
        style = """
        QWidget {
            background-color: #F3F4F6;
            color: #111827;
            font-size: 10pt;
        }

        QMainWindow, QTabWidget::pane {
            background-color: #F3F4F6;
        }

        QTabBar::tab {
            background-color: #E5E7EB;
            color: #111827;
            border: 1px solid #CBD5E1;
            padding: 8px 14px;
            min-height: 18px;
            margin-right: 4px;
            border-top-left-radius: 6px;
            border-top-right-radius: 6px;
        }

        QTabBar::tab:selected {
            background-color: #FFFFFF;
            color: #111827;
        }

        QTabBar::tab:hover {
            background-color: #DCE3EA;
        }

        QLabel {
            color: #111827;
            background: transparent;
        }

        QLineEdit, QTextEdit, QComboBox, QDoubleSpinBox, QTableWidget, QTreeWidget {
            background-color: #FFFFFF;
            color: #111827;
            border: 1px solid #CBD5E1;
            border-radius: 6px;
            padding: 6px;
        }

        QLineEdit:focus, QTextEdit:focus, QComboBox:focus, QDoubleSpinBox:focus {
            border: 1px solid #3B82F6;
        }

        QHeaderView::section {
            background-color: #E5E7EB;
            color: #111827;
            border: 1px solid #CBD5E1;
            padding: 6px;
            font-weight: bold;
        }

        QGroupBox {
            border: 1px solid #CBD5E1;
            border-radius: 8px;
            margin-top: 12px;
            padding-top: 12px;
            background-color: #FFFFFF;
        }

        QGroupBox::title {
            subcontrol-origin: margin;
            left: 10px;
            padding: 0 6px;
            color: #111827;
            font-weight: bold;
            background-color: #FFFFFF;
        }

        QPushButton {
            background-color: #FFFFFF;
            color: #111827;
            border: 1px solid #CBD5E1;
            border-radius: 6px;
            padding: 8px 14px;
            min-height: 18px;
        }

        QPushButton:hover {
            background-color: #F8FAFC;
        }

        QPushButton#primary {
            background-color: #2563EB;
            color: white;
            border: 1px solid #2563EB;
        }

        QPushButton#primary:hover {
            background-color: #1D4ED8;
        }

        QPushButton#secondary {
            background-color: #FFFFFF;
            color: #111827;
            border: 1px solid #CBD5E1;
        }

        QPushButton#secondary:hover {
            background-color: #F8FAFC;
        }

        QPushButton#danger {
            background-color: #DC2626;
            color: white;
            border: 1px solid #DC2626;
        }

        QPushButton#danger:hover {
            background-color: #B91C1C;
        }

        QPushButton:disabled {
            background-color: #E5E7EB;
            color: #9CA3AF;
            border: 1px solid #CBD5E1;
        }

        QRadioButton, QCheckBox {
            spacing: 8px;
            color: #111827;
            background: transparent;
        }

        QRadioButton::indicator, QCheckBox::indicator {
            width: 16px;
            height: 16px;
        }

        QRadioButton::indicator:unchecked {
            border: 2px solid #64748B;
            border-radius: 8px;
            background: #FFFFFF;
        }

        QRadioButton::indicator:checked {
            border: 2px solid #2563EB;
            border-radius: 8px;
            background: #2563EB;
        }

        QCheckBox::indicator:unchecked {
            border: 2px solid #64748B;
            border-radius: 3px;
            background: #FFFFFF;
        }

        QCheckBox::indicator:checked {
            border: 2px solid #2563EB;
            border-radius: 3px;
            background: #2563EB;
        }
        """

    widget.setStyleSheet(style)