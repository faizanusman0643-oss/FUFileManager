# PortableExcel

Excel jaisa portable spreadsheet app. Ek hi `PortableExcel.exe` file hai, jo bina install kiye USB ya kisi bhi folder se chalti hai. Microsoft Excel ki zaroorat nahi.

## Download
GitHub par **Releases → "PortableExcel (latest Windows .exe)"** se `PortableExcel.exe` download karo.
Ya phir **Actions → "Build PortableExcel (Windows .exe)"** ke aakhri run mein se artifact download karo.

## Password
- App khulte hi password maangti hai. Default password: **12345643**
- Password badalne ke liye: **File → Change app password…**
- Naya password `PortableExcel.ini` mein save hota hai, jo exe ke saath wale folder mein banti hai. Is tarah setting USB ke saath chalti hai.
- Code mein password plain text mein nahi hai, sirf uska PBKDF2 hash rakha gaya hai.

## Features
- Excel jaisa look: ribbon (File, Home, Insert, Formulas, Data, View, Developer), formula bar, Name Box, sheet tabs, status bar (Sum/Average/Count), zoom.
- 190+ formulas: SUM, IF, VLOOKUP, XLOOKUP, INDEX/MATCH, COUNTIFS, SUMIFS, TEXT, dates waghera.
- Files: `.xlsx`, `.xlsm`, `.xltx`, `.xltm` aur `.csv` open/save. Password wali Excel files bhi khul jati hain, aur password ke saath save bhi ho sakti hain.
- Formatting: bold/italic/underline, font, rang, fill, alignment, number formats, column width.
- Editing: copy/paste (Excel se bhi), fill handle, Ctrl+D/Ctrl+R, undo/redo (100 steps), sort, find & replace, rows/columns insert aur delete, freeze panes.
- **VBA macros:**
  - `.xlsm` kholne par "Enable Content" bar aati hai.
  - Macros dialog: Alt+F8.
  - VBA Editor: Alt+F11, jismein Immediate window aur F5 se run hota hai.
  - Events chalte hain: Workbook_Open, Worksheet_Change, BeforeSave waghera.
  - Sub/Function, loops, Select Case, With, On Error, arrays, Dictionary, Collection, class modules, Range/Cells/Worksheets, WorksheetFunction, MsgBox/InputBox, file I/O.

## Limitations (abhi nahi hai)
- UserForms, ActiveX controls, Windows API (`Declare`) calls, charts, borders, merged cells, comments, conditional formatting.
- PortableExcel mein edit kiye gaye macros sirf PortableExcel mein chalte hain. Asal Excel us file mein purane macros hi chalayega.

## Developers ke liye
```
pip install -r requirements.txt
python run.py                        # app chalao
python -m pytest tests -q            # tests (QT_QPA_PLATFORM=offscreen PORTABLEEXCEL_NO_LOCK=1)
python tools/screenshots.py          # screenshots/ folder mein har feature ki tasveer
pyinstaller --noconfirm PortableExcel.spec   # dist/PortableExcel.exe
```
