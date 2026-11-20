"""Synthetic tests for user-workbook compatibility features: worksheet UDFs,
WithEvents, MSForms.DataObject, Windows API emulation, printing / PDF."""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from portable_excel.model import Workbook  # noqa: E402
from portable_excel.engine import FormulaEngine  # noqa: E402
from portable_excel.forms import new_form, new_control  # noqa: E402
from portable_excel.vba import VBARuntime  # noqa: E402
from test_userforms import FormHost, FakeView, click  # noqa: E402


class CompatView(FakeView):
    def set_window_flags(self, **k):
        self.calls.append(("flags", k))

    def set_window_state(self, s):
        self.calls.append(("state", s))


class CompatHost(FormHost):
    def __init__(self, scripts=None):
        super().__init__(scripts)
        self.printed = []
        self.pdfs = []
        self.clip = None

    def create_form_view(self, inst):
        v = CompatView(self, inst)
        self.views.append(v)
        return v

    def print_range(self, sheet, r1, c1, r2, c2, preview=False, copies=1):
        self.printed.append((sheet.name, r1, c1, r2, c2, preview, copies))

    def export_pdf(self, sheet, r1, c1, r2, c2, path):
        self.pdfs.append((sheet.name, r1, c1, r2, c2, os.path.basename(path)))

    def get_save_filename(self, initial="", filter="", title=""):
        return "/x/out.pdf"

    def printer_setup(self):
        return True

    def set_clipboard_text(self, t):
        self.clip = t

    def get_clipboard_text(self):
        return self.clip


def setup(mods, host=None):
    book = Workbook()
    eng = FormulaEngine(book)
    host = host or CompatHost()
    rt = VBARuntime(book, eng, host)
    errs = rt.load_modules(mods)
    assert errs == [], [str(e) for e in errs]
    return rt, book, eng, host


def test_udf_in_formulas():
    rt, book, eng, host = setup([{"name": "Mod", "type": "standard", "code": '''
Public Function Wrap(s As String, Optional n As Integer = 0) As String
    Application.Volatile
    Wrap = "<" & s & ">" & n
End Function
Public Function SumTwice(r As Range) As Double
    Dim c As Range
    For Each c In r
        SumTwice = SumTwice + c.Value * 2
    Next
End Function
Public Function Boom(x)
    Boom = 1 / 0
End Function
Private Function Hidden()
    Hidden = 1
End Function
Public Function Writes()
    Range("Z1").Value = 1
End Function'''}])
    sh = book.active
    book.set_cell(sh, 0, 0, '=IF(1,"=%6200","")')
    book.set_cell(sh, 1, 0, "2")
    book.set_cell(sh, 2, 0, "3")
    for i, f in enumerate(["=Wrap(A1)", "=wrap(A1,6)", "=SumTwice(A2:A3)+1", "=Boom(1)",
                           "=Hidden()", "=Nope()", "=Mod.Wrap(\"x\")", "=Writes()"]):
        book.set_cell(sh, i, 3, f)
    rt.attach_udfs(eng)
    vals = [eng.value(sh, i, 3) for i in range(8)]
    assert vals[0] == "<=%6200>0" and vals[1] == "<=%6200>6"
    assert vals[2] == 11.0
    assert [str(v) for v in vals[3:6]] == ["#VALUE!", "#NAME?", "#NAME?"]
    assert str(vals[7]) == "#VALUE!" and eng.value(sh, 0, 25) is None


def test_withevents_form_and_class():
    form = new_form()
    code = '''
Private WithEvents btnCaps As MSForms.CommandButton
Private rows As Collection
Private Sub UserForm_Initialize()
    Set btnCaps = Me.Controls.Add("Forms.CommandButton.1", "btnCaps")
    btnCaps.Caption = "A a": btnCaps.Left = 10: btnCaps.Font.Bold = True
    btnCaps.BackColor = RGB(1, 2, 3): btnCaps.Visible = True
    Set rows = New Collection
    Dim i As Long, r As clsRow, t As MSForms.TextBox
    For i = 1 To 2
        Set t = Me.Controls.Add("Forms.TextBox.1", "txt" & i)
        Set r = New clsRow
        r.Init Me, i
        Set r.TxtA = t
        rows.Add r
    Next
End Sub
Private Sub btnCaps_Click()
    Debug.Print "caps"
End Sub
Public Sub RowChanged(i As Long, s As String)
    Debug.Print "row", i, s
End Sub'''
    cls = '''
Public WithEvents TxtA As MSForms.TextBox
Private mParent As Object, mIdx As Long
Public Sub Init(p As Object, i As Long)
    Set mParent = p: mIdx = i
End Sub
Private Sub TxtA_Change()
    mParent.RowChanged mIdx, TxtA.Text
End Sub'''
    rt, book, eng, host = setup([
        {"name": "M", "type": "standard", "code": "Sub Go()\n UserForm1.Show\nEnd Sub"},
        {"name": "UserForm1", "type": "form", "code": code, "form": form},
        {"name": "clsRow", "type": "class", "code": cls}],
        CompatHost([[click("btnCaps"), click("btnCaps"),
                     lambda i: (i.user_set("txt2", "value", "hi"), i.fire("txt2", "Change"))]]))
    rt.run("Go")
    assert host.prints[:2] == ["caps", "caps"]
    assert host.prints[2].split() == ["row", "2", "hi"]
    c = host.views[0].inst.find_ctrl("btnCaps")
    assert c["bold"] is True and c["back_color"] == "#010203" and c["left"] == 10.0


def test_dataobject_clipboard():
    rt, book, eng, host = setup([{"name": "M", "type": "standard", "code": '''
Sub Go()
    Dim d As Object, e As New MSForms.DataObject
    Set d = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    d.SetText "copied"
    d.PutInClipboard
    e.GetFromClipboard
    Debug.Print e.GetText
End Sub'''}])
    rt.run("Go")
    assert host.clip == "copied" and host.prints == ["copied"]


def test_windows_api_emulation_and_mouse_wheel():
    code = '''
#If VBA7 Then
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal c As String, ByVal n As String) As LongPtr
Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal h As LongPtr, ByVal i As Long) As Long
Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal h As LongPtr, ByVal i As Long, ByVal v As Long) As Long
Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal h As LongPtr) As Long
Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal h As LongPtr, ByVal n As Long) As Long
#End If
Private Sub UserForm_Activate()
    Dim h As LongPtr, st As Long
    h = FindWindow("ThunderDFrame", Me.Caption)
    st = GetWindowLong(h, -16)
    SetWindowLong h, -16, st Or &H40000 Or &H10000
    DrawMenuBar h
    ShowWindow h, 3
    HookWheel Me
End Sub
Public Sub WheelScrollDown()
    Debug.Print "wheel down"
End Sub'''
    hook = '''
Private Declare PtrSafe Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal i As Long, ByVal f As LongPtr, ByVal m As LongPtr, ByVal t As Long) As LongPtr
Private Declare PtrSafe Function CallNextHookEx Lib "user32" (ByVal h As LongPtr, ByVal n As Long, ByVal w As LongPtr, ByVal l As LongPtr) As LongPtr
Private Declare PtrSafe Function GetCursorPos Lib "user32" (ByRef p As POINTAPI) As Long
Private Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal h As LongPtr, ByRef r As RECT) As Long
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal c As String, ByVal n As String) As LongPtr
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef d As Any, ByVal s As LongPtr, ByVal l As Long)
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private gHook As LongPtr, gHwnd As LongPtr, gForm As Object
Public Sub HookWheel(f As Object)
    Set gForm = f
    gHwnd = FindWindow("ThunderDFrame", f.Caption)
    gHook = SetWindowsHookEx(14, AddressOf HookProc, 0, 0)
End Sub
Public Function HookProc(ByVal n As Long, ByVal w As LongPtr, ByVal l As LongPtr) As LongPtr
    Dim pt As POINTAPI, rc As RECT, m As Long, delta As Long
    If w = &H20A Then
        GetCursorPos pt
        GetWindowRect gHwnd, rc
        If pt.X >= rc.Left And pt.X <= rc.Right And pt.Y >= rc.Top And pt.Y <= rc.Bottom Then
            CopyMemory m, l + 8, 4
            delta = m \\ 65536
            If delta < 0 Then CallByName gForm, "WheelScrollDown", VbMethod
            HookProc = 1
            Exit Function
        End If
    End If
    HookProc = CallNextHookEx(gHook, n, w, l)
End Function'''
    rt, book, eng, host = setup([
        {"name": "M", "type": "standard", "code": "Sub Go()\n UserForm1.Show vbModeless\nEnd Sub"},
        {"name": "Hook", "type": "standard", "code": hook},
        {"name": "UserForm1", "type": "form", "code": code, "form": new_form("Wheel test")}])
    rt.run("Go")
    v = host.views[0]
    assert ("flags", {"resizable": True, "minimize": False, "maximize": True, "caption": True}) in v.calls
    assert ("state", "maximized") in v.calls
    assert rt.mouse_wheel(-120) is True
    assert host.prints == ["wheel down"]


def test_printing_pdf_and_dialogs():
    rt, book, eng, host = setup([{"name": "M", "type": "standard", "code": '''
Sub Go()
    Range("A1:C5").Value = 1
    ActiveSheet.PrintOut Copies:=2
    Range("B2:C3").Select
    Selection.PrintOut Preview:=True
    ActiveSheet.PrintPreview
    Dim f
    f = Application.GetSaveAsFilename(InitialFileName:="x.pdf", FileFilter:="PDF (*.pdf), *.pdf")
    Range("A1:B2").ExportAsFixedFormat Type:=xlTypePDF, Filename:=f
    Debug.Print Application.Dialogs(xlDialogPrinterSetup).Show
End Sub'''}])
    rt.run("Go")
    assert host.printed == [("Sheet1", 0, 0, 4, 2, False, 2), ("Sheet1", 1, 1, 2, 2, True, 1),
                            ("Sheet1", 0, 0, 4, 2, True, 1)]
    assert host.pdfs == [("Sheet1", 0, 0, 1, 1, "out.pdf")]
    assert host.prints == ["True"]
