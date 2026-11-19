# PyInstaller spec: builds a single portable PortableExcel.exe
# Build:  pyinstaller --noconfirm PortableExcel.spec
from PyInstaller.utils.hooks import collect_submodules

hidden = (collect_submodules("portable_excel") + collect_submodules("openpyxl")
          + collect_submodules("oletools") + collect_submodules("msoffcrypto")
          + collect_submodules("barcode") + collect_submodules("segno")
          + ["olefile", "et_xmlfile", "PIL.Image", "PIL.PngImagePlugin"])

a = Analysis(
    ["run.py"],
    pathex=["."],
    hiddenimports=hidden,
    excludes=["tkinter", "PySide6.QtWebEngineCore", "PySide6.QtQml", "PySide6.QtQuick",
              "PySide6.Qt3DCore", "PySide6.QtMultimedia", "PySide6.QtPdf"],
)
pyz = PYZ(a.pure)
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    name="PortableExcel",
    console=False,
    upx=False,
)
