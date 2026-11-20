"""Excel-like light theme (Fusion style + stylesheet)."""
from __future__ import annotations

from PySide6.QtGui import QColor, QFont, QFontDatabase, QPalette
from PySide6.QtWidgets import QApplication

GREEN = "#217346"
DARK_GREEN = "#185C37"
LIGHT_GREEN = "#E9F5EE"

STYLESHEET = f"""
QMainWindow, QDialog {{ background: #FFFFFF; }}
QWidget#titleStrip {{ background: {GREEN}; }}
QLabel#titleLabel {{ color: white; font-size: 9pt; }}
QLabel#titleRight {{ color: #E8F3EC; }}
QToolButton#qatButton {{ border: none; padding: 4px; color: white; }}
QToolButton#qatButton:hover {{ background: #2E8B57; }}
QToolButton#qatButton:disabled {{ background: transparent; }}

QWidget#ribbonTabStrip {{ background: {GREEN}; }}
QTabBar#ribbonTabs {{ background: {GREEN}; }}
QTabBar#ribbonTabs::tab {{ color: white; background: {GREEN}; border: none;
    padding: 5px 14px 6px 14px; margin: 0; min-width: 36px; }}
QTabBar#ribbonTabs::tab:first {{ background: {DARK_GREEN}; padding-left: 18px; padding-right: 18px; }}
QTabBar#ribbonTabs::tab:hover {{ background: #2E8B57; }}
QTabBar#ribbonTabs::tab:selected {{ background: #F3F3F3; color: {DARK_GREEN}; }}
QStackedWidget#ribbonBody {{ background: #F3F3F3; border-bottom: 1px solid #D4D4D4; }}
QStackedWidget#ribbonBody > QWidget {{ background: #F3F3F3; }}
QFrame#ribbonGroup {{ background: transparent; }}
QLabel#groupCaption {{ color: #666666; font-size: 8pt; padding-bottom: 1px; }}
QFrame#groupSep {{ color: #D4D4D4; background: #D4D4D4; max-width: 1px; margin: 6px 0px 6px 0px; }}
QFrame#rowSep {{ color: #D4D4D4; max-width: 1px; }}
QToolButton#largeButton, QToolButton#smallButton {{ border: 1px solid transparent; border-radius: 2px;
    padding: 1px 3px; color: #262626; background: transparent; }}
QToolButton#largeButton:hover, QToolButton#smallButton:hover {{ background: #C5E5D3; border-color: #A3D6B8; }}
QToolButton#largeButton:checked, QToolButton#smallButton:checked {{ background: #B5DCC6; border-color: #7FBF9B; }}
QToolButton#largeButton:pressed, QToolButton#smallButton:pressed {{ background: #9FD0B4; }}
QToolButton::menu-indicator {{ subcontrol-origin: padding; subcontrol-position: bottom center; }}

QWidget#messageBar {{ background: #FFF4CE; border-bottom: 1px solid #E8D58A; }}
QLabel#messageBold {{ font-weight: bold; color: #3B3B3B; }}
QPushButton#enableContent {{ background: white; border: 1px solid #8A8A8A; padding: 3px 12px; }}
QPushButton#enableContent:hover {{ background: #F3F3F3; }}

QWidget#formulaBar {{ background: #FFFFFF; border-bottom: 1px solid #D4D4D4; }}
QComboBox#nameBox {{ border: 1px solid #C6C6C6; padding: 1px 4px; background: white; }}
QLineEdit#formulaEdit {{ border: 1px solid #C6C6C6; padding: 1px 4px; background: white;
    selection-background-color: #9FD5B7; selection-color: black; }}
QToolButton#fbButton {{ border: none; color: #555555; padding: 2px 5px; }}
QToolButton#fbButton:hover {{ background: #E1E1E1; }}

QWidget#sheetBar {{ background: #F3F3F3; border-top: 1px solid #D4D4D4; }}
QTabBar#sheetTabs::tab {{ background: #F3F3F3; color: #444444; border: none;
    border-right: 1px solid #D4D4D4; padding: 4px 16px; min-width: 50px; }}
QTabBar#sheetTabs::tab:hover {{ background: #E1E1E1; }}
QTabBar#sheetTabs::tab:selected {{ background: white; color: {GREEN}; font-weight: bold;
    border-bottom: 3px solid {GREEN}; }}
QToolButton#sheetNav {{ border: none; color: #666666; padding: 2px 4px; }}
QToolButton#sheetNav:hover {{ background: #E1E1E1; }}
QToolButton#addSheet {{ border: 1px solid transparent; border-radius: 10px; color: #444444;
    font-size: 13pt; padding: 0px 5px; }}
QToolButton#addSheet:hover {{ background: #E1E1E1; }}

QStatusBar {{ background: #F3F3F3; color: #444444; border-top: 1px solid #D4D4D4; }}
QStatusBar QLabel {{ color: #444444; padding: 0 6px; }}
QStatusBar QToolButton {{ border: none; padding: 0 4px; color: #444444; }}
QStatusBar QToolButton:hover {{ background: #E1E1E1; }}
QSlider::groove:horizontal {{ height: 2px; background: #8A8A8A; }}
QSlider::handle:horizontal {{ background: #444444; width: 4px; margin: -6px 0; }}

QWidget#backstage {{ background: white; }}
QWidget#backstageNav {{ background: {GREEN}; }}
QPushButton#bsItem {{ color: white; background: transparent; border: none; text-align: left;
    padding: 9px 22px; font-size: 10pt; }}
QPushButton#bsItem:hover {{ background: #2E8B57; }}
QPushButton#bsItem:checked {{ background: {DARK_GREEN}; }}
QLabel#bsTitle {{ font-size: 22pt; color: #262626; font-weight: 300; }}
QLabel#bsHeading {{ font-size: 11pt; color: #262626; font-weight: bold; }}
QPushButton#bsTile {{ background: white; border: 1px solid #D4D4D4; padding: 10px; text-align: left; }}
QPushButton#bsTile:hover {{ background: #E9F5EE; border-color: {GREEN}; }}
QListWidget#recentList {{ border: none; font-size: 10pt; }}
QListWidget#recentList::item {{ padding: 6px; }}
QListWidget#recentList::item:hover {{ background: #E9F5EE; }}

QPushButton {{ padding: 4px 14px; border: 1px solid #ADADAD; background: #FDFDFD; border-radius: 2px; }}
QPushButton:hover {{ background: #E9F5EE; border-color: {GREEN}; }}
QPushButton:default, QPushButton#primary {{ border: 1px solid {GREEN}; }}
QMenu {{ background: white; border: 1px solid #C6C6C6; }}
QMenu::item {{ padding: 5px 28px 5px 28px; }}
QMenu::item:selected {{ background: #E1E1E1; color: black; }}
QMenu::separator {{ height: 1px; background: #E1E1E1; margin: 3px 8px; }}
QTabWidget::pane {{ border: 1px solid #D4D4D4; }}
QTreeWidget, QListWidget, QPlainTextEdit {{ selection-background-color: #CCE8D8; selection-color: black; }}
QDockWidget::title {{ background: #F3F3F3; padding: 3px; }}
"""


def pick_font_family() -> str:
    fams = set(QFontDatabase.families())
    for fam in ("Segoe UI", "Calibri", "Inter", "Noto Sans", "DejaVu Sans", "Liberation Sans",
                "FreeSans", "Arial"):
        if fam in fams:
            return fam
    return QApplication.font().family()


def apply(app: QApplication) -> None:
    app.setStyle("Fusion")
    pal = QPalette()
    pal.setColor(QPalette.Window, QColor("#FFFFFF"))
    pal.setColor(QPalette.WindowText, QColor("#262626"))
    pal.setColor(QPalette.Base, QColor("#FFFFFF"))
    pal.setColor(QPalette.AlternateBase, QColor("#F7F7F7"))
    pal.setColor(QPalette.Text, QColor("#000000"))
    pal.setColor(QPalette.Button, QColor("#F3F3F3"))
    pal.setColor(QPalette.ButtonText, QColor("#262626"))
    pal.setColor(QPalette.Highlight, QColor(GREEN))
    pal.setColor(QPalette.HighlightedText, QColor("#FFFFFF"))
    pal.setColor(QPalette.ToolTipBase, QColor("#FFFFFF"))
    pal.setColor(QPalette.ToolTipText, QColor("#262626"))
    pal.setColor(QPalette.PlaceholderText, QColor("#8A8A8A"))
    app.setPalette(pal)
    f = QFont(pick_font_family())
    f.setPointSizeF(9)
    app.setFont(f)
    app.setStyleSheet(STYLESHEET)
