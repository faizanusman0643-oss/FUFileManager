"""Tests for the VBA interpreter (portable_excel.vba)."""
import os
import sys
import time

import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from portable_excel.model import Workbook  # noqa: E402
from portable_excel.engine import FormulaEngine  # noqa: E402
from portable_excel.vba import VBARuntime, VBAHost, VBAError  # noqa: E402


class RecHost(VBAHost):
    def __init__(self, msg_answers=None, input_answers=None):
        super().__init__()
        self.msgs = []
        self.prints = []
        self.inputs = []
        self.status = []
        self.selected = None
        self.activated = []
        self.refreshes = 0
        self.events = 0
        self.msg_answers = list(msg_answers or [])
        self.input_answers = list(input_answers or [])

    def msgbox(self, prompt, buttons=0, title=""):
        self.msgs.append((prompt, buttons, title))
        return self.msg_answers.pop(0) if self.msg_answers else 1

    def inputbox(self, prompt, title="", default=""):
        self.inputs.append((prompt, title, default))
        return self.input_answers.pop(0) if self.input_answers else default

    def debug_print(self, text):
        self.prints.append(text)

    def select(self, sheet, r1, c1, r2, c2):
        self.selected = (sheet, r1, c1, r2, c2)

    def get_selection(self):
        return self.selected

    def activate_sheet(self, sheet):
        self.activated.append(sheet.name)

    def refresh(self):
        self.refreshes += 1

    def status_bar(self, text):
        self.status.append(text)

    def process_events(self):
        self.events += 1


def make(code, modules=None, sheets=1, host=None):
    book = Workbook()
    for i in range(sheets - 1):
        book.add_sheet()
    eng = FormulaEngine(book)
    host = host or RecHost()
    rt = VBARuntime(book, eng, host)
    mods = [{"name": "Module1", "type": "standard", "code": code}] if code is not None else []
    mods += modules or []
    errs = rt.load_modules(mods)
    assert errs == [], [str(e) for e in errs]
    return rt, book, eng, host


def val(eng, sheet, ref):
    from portable_excel.model import parse_a1
    r, c = parse_a1(ref)
    return eng.value(sheet, r, c)


def run_print(code, macro="Test"):
    rt, book, eng, host = make(code)
    rt.run(macro)
    return host.prints


# ---------------------------------------------------------------------------
# Basic language
# ---------------------------------------------------------------------------
def test_sum_column_to_last_row():
    rt, book, eng, host = make('''
Sub SumColumn()
    Dim lastRow As Long, i As Long, total As Double
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    For i = 2 To lastRow
        Cells(i, 2).Value = Cells(i, 1).Value * 2
        total = total + Cells(i, 1).Value
    Next i
    Cells(lastRow + 1, 1).Value = total
    Range("D1").Value = lastRow
End Sub''')
    sh = book.active
    book.set_cell(sh, 0, 0, "Amount")
    for i in range(1, 11):
        book.set_cell(sh, i, 0, str(i * 1.5))
    rt.run("SumColumn")
    assert val(eng, sh, "A12") == sum(i * 1.5 for i in range(1, 11))
    assert val(eng, sh, "B11") == 30.0
    assert val(eng, sh, "D1") == 11.0


def test_with_block_formatting():
    rt, book, eng, host = make('''
Sub Fmt()
    With Worksheets("Sheet1").Range("A1:C1")
        .Value = "Header"
        .Font.Bold = True
        .Font.Color = RGB(255, 0, 0)
        .Interior.Color = vbYellow
        .HorizontalAlignment = xlCenter
        .NumberFormat = "0.00"
        With .Font
            .Size = 14
            .Italic = True
        End With
    End With
    Columns("A").ColumnWidth = 20
End Sub''')
    rt.run("Fmt")
    sh = book.active
    st = sh.get_style(0, 1)
    assert st["bold"] is True and st["italic"] is True
    assert st["color"] == "#FF0000"
    assert st["fill"] == "#FFFF00"
    assert st["align"] == "center"
    assert st["font_size"] == 14
    assert st["number_format"] == "0.00"
    assert sh.col_widths[0] == 145
    assert val(eng, sh, "C1") == "Header"


def test_select_case_grading():
    prints = run_print('''
Function Grade(score As Integer) As String
    Select Case score
        Case Is >= 90: Grade = "A"
        Case 80 To 89: Grade = "B"
        Case 70, 71, 72 To 79: Grade = "C"
        Case Else: Grade = "F"
    End Select
End Function
Sub Test()
    Dim s
    For Each s In Array(95, 85, 72, 10)
        Debug.Print Grade(CInt(s))
    Next
    Select Case "apple"
        Case "banana": Debug.Print "no"
        Case "apple", "pear": Debug.Print "fruit"
    End Select
End Sub''')
    assert prints == ["A", "B", "C", "F", "fruit"]


def test_string_functions():
    prints = run_print('''
Sub Test()
    Dim s As String
    s = "  Hello World  "
    Debug.Print Len(s); "|"; Trim(s); "|"; LTrim(s); "|"; RTrim(s); "|"
    Debug.Print Left("Hello", 2), Right("Hello", 3), Mid("Hello", 2, 3), Mid$("Hello", 3)
    Debug.Print UCase("abc") & LCase("DEF"), InStr("Hello", "l"), InStr(4, "Hello", "l"), InStrRev("Hello", "l")
    Debug.Print Replace("a-b-c", "-", "+"), StrReverse("abc"), Space(3) & "x", String(3, "z")
    Debug.Print Join(Split("a,b,c", ","), "|"), UBound(Split("a,b,c", ","))
    Debug.Print StrComp("abc", "ABC"), StrComp("abc", "ABC", vbTextCompare), StrConv("hello world", vbProperCase)
    Debug.Print Chr(65) & Asc("a"), Val("12abc"), Str(5), CStr(3.5), Hex(255), Oct(8)
    Debug.Print "abc" Like "a*", "abc" Like "a?d", "a1" Like "a#", "b" Like "[a-c]"
    Debug.Print InStr(1, "HELLO", "ll", vbTextCompare), Format("abc", ">")
End Sub''')
    assert prints[0] == " 15 |Hello World|Hello World  |  Hello World|"
    assert prints[1].split() == ["He", "llo", "ell", "llo"]
    assert prints[2].split() == ["ABCdef", "3", "4", "4"]
    assert prints[3].split() == ["a+b+c", "cba", "x", "zzz"]
    assert prints[4].split() == ["a|b|c", "2"]
    assert prints[5].split() == ["1", "0", "Hello", "World"]
    assert prints[6].split() == ["A97", "12", "5", "3.5", "FF", "10"]
    assert prints[7].split() == ["True", "False", "True", "True"]
    assert prints[8].split() == ["3", "ABC"]


def test_arrays_redim_preserve():
    prints = run_print('''
Sub Test()
    Dim arr() As Long, i As Long, n As Long
    For i = 1 To 5
        n = n + 1
        ReDim Preserve arr(1 To n)
        arr(n) = i * i
    Next
    Debug.Print LBound(arr), UBound(arr), arr(3), arr(5)
    Dim m(1 To 3, 0 To 2) As String
    m(2, 1) = "x"
    Debug.Print UBound(m, 1), UBound(m, 2), m(2, 1), "[" & m(1, 1) & "]"
    Dim v As Variant
    v = Array("a", "b", "c")
    Debug.Print IsArray(v), UBound(v), v(1)
    Erase arr
    On Error Resume Next
    Debug.Print UBound(arr)
    Debug.Print Err.Number
    Dim fixedArr(5) As Integer
    Debug.Print UBound(fixedArr), fixedArr(5)
End Sub''')
    assert prints[0].split() == ["1", "5", "9", "25"]
    assert prints[1].split() == ["3", "2", "x", "[]"]
    assert prints[2].split() == ["True", "2", "b"]
    assert prints[-2].strip() == "9"
    assert prints[-1].split() == ["5", "0"]


def test_dictionary_unique_count():
    rt, book, eng, host = make('''
Sub CountUnique()
    Dim d As Object, c As Range, k
    Set d = CreateObject("Scripting.Dictionary")
    For Each c In Range("A1:A8")
        If Not IsEmpty(c.Value) Then
            If d.Exists(c.Value) Then
                d(c.Value) = d(c.Value) + 1
            Else
                d.Add c.Value, 1
            End If
        End If
    Next c
    Dim r As Long: r = 1
    For Each k In d.Keys
        Cells(r, 3).Value = k
        Cells(r, 4).Value = d(k)
        r = r + 1
    Next
    Debug.Print d.Count
    d.Remove "apple"
    Debug.Print d.Count, d.Exists("apple")
    Dim d2 As New Scripting.Dictionary
    d2.CompareMode = vbTextCompare
    d2("Key") = 1
    Debug.Print d2.Exists("KEY"), d2.Item("key")
End Sub''')
    sh = book.active
    for i, v in enumerate(["apple", "pear", "apple", "fig", "pear", "apple", "", "kiwi"]):
        book.set_cell(sh, i, 0, v)
    rt.run("CountUnique")
    assert host.prints[0].strip() == "4"
    assert val(eng, sh, "C1") == "apple" and val(eng, sh, "D1") == 3.0
    assert val(eng, sh, "C2") == "pear" and val(eng, sh, "D2") == 2.0
    assert host.prints[1].split() == ["3", "False"]
    assert host.prints[2].split() == ["True", "1"]


def test_collection():
    prints = run_print('''
Sub Test()
    Dim c As New Collection, x
    c.Add "one"
    c.Add "two", "k2"
    c.Add "zero", Before:=1
    Debug.Print c.Count, c(1), c("k2"), c.Item(3)
    c.Remove 1
    For Each x In c
        Debug.Print x
    Next
    On Error Resume Next
    c.Add "dup", "k2"
    Debug.Print Err.Number
End Sub''')
    assert prints[0].split() == ["3", "zero", "two", "two"]
    assert prints[1:3] == ["one", "two"]
    assert prints[3].strip() == "457"


def test_class_module():
    code = '''
Sub Test()
    Dim a As CAccount
    Set a = New CAccount
    a.Owner = "Ann"
    a.Deposit 100
    a.Deposit 50
    Debug.Print a.Owner, a.Balance, TypeName(a)
    Dim b As CAccount
    Set b = a
    b.Deposit 1
    Debug.Print a.Balance, a Is b, b Is Nothing
    Set a = Nothing
    Set b = Nothing
    Debug.Print "done"
End Sub'''
    cls = '''VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CAccount"
Option Explicit
Private mBalance As Double
Public Owner As String

Private Sub Class_Initialize()
    mBalance = 0
    Debug.Print "init"
End Sub

Private Sub Class_Terminate()
    Debug.Print "terminate"
End Sub

Public Sub Deposit(ByVal amount As Double)
    mBalance = mBalance + amount
End Sub

Public Property Get Balance() As Double
    Balance = mBalance
End Property
'''
    rt, book, eng, host = make(code, [{"name": "CAccount", "type": "class", "code": cls}])
    rt.run("Test")
    p = host.prints
    assert p[0] == "init"
    assert p[1].split() == ["Ann", "150", "CAccount"]
    assert p[2].split() == ["151", "True", "False"]
    assert p[3:] == ["terminate", "done"]


def test_error_handling():
    prints = run_print('''
Sub Test()
    Dim x As Integer
    On Error Resume Next
    x = 1 / 0
    Debug.Print Err.Number, Err.Description
    Err.Clear
    Debug.Print Err.Number
    On Error GoTo Handler
    x = 40000
    Debug.Print "not reached"
    Exit Sub
Handler:
    Debug.Print "handled", Err.Number
    Resume Done
Done:
    Debug.Print "done"
    On Error GoTo 0
End Sub
Sub Test2()
    On Error GoTo EH
    Err.Raise vbObjectError + 513, "MySrc", "Custom failure"
    Exit Sub
EH:
    Debug.Print Err.Number - vbObjectError, Err.Source, Err.Description
    Resume Next
End Sub
Sub Test3()
    Dim i As Integer
    On Error GoTo EH
    i = 0
    Debug.Print 10 / i
    Debug.Print "after"
    Exit Sub
EH:
    i = 2
    Resume
End Sub''')
    assert prints[0].split()[0] == "11" and "Division by zero" in prints[0]
    assert prints[1].strip() == "0"
    assert prints[2].split() == ["handled", "6"]
    assert prints[3] == "done"


def test_error_handlers_resume():
    rt, book, eng, host = make('''
Sub Test2()
    On Error GoTo EH
    Err.Raise vbObjectError + 513, "MySrc", "Custom failure"
    Debug.Print "next"
    Exit Sub
EH:
    Debug.Print Err.Number - vbObjectError, Err.Source, Err.Description
    Resume Next
End Sub
Sub Test3()
    Dim i As Integer
    On Error GoTo EH
    i = 0
    Debug.Print 10 / i
    Debug.Print "after"
    Exit Sub
EH:
    i = 2
    Resume
End Sub''')
    rt.run("Test2")
    assert host.prints[0].split() == ["513", "MySrc", "Custom", "failure"]
    assert host.prints[1] == "next"
    rt.run("Test3")
    assert host.prints[2:] == [" 5 ", "after"]


def test_runtime_error_line_numbers():
    rt, book, eng, host = make('''Sub A()
    Dim x
    x = 1
    B
End Sub

Sub B()
    Dim arr(3)
    arr(5) = 1
End Sub''')
    with pytest.raises(VBAError) as ei:
        rt.run("A")
    e = ei.value
    assert e.vba_number == 9
    assert e.line == 9
    assert e.module == "Module1"
    assert "Subscript out of range" in e.message
    assert e.message.startswith("Run-time error '9'")


def test_type_mismatch_and_object_errors():
    rt, book, eng, host = make('''
Sub TM()
    Dim n As Long
    n = "abc"
End Sub
Sub NotSet()
    Dim r As Range
    r.Value = 1
End Sub
Sub Ovf()
    Dim i As Integer
    i = 32767
    i = i + 1
End Sub''')
    for name, num in (("TM", 13), ("NotSet", 91), ("Ovf", 6)):
        with pytest.raises(VBAError) as ei:
            rt.run(name)
        assert ei.value.vba_number == num


def test_option_explicit():
    rt, book, eng, host = make('''Option Explicit
Sub Good()
    Dim a As Long
    a = 1
End Sub
Sub Bad()
    Dim a As Long
    a = 1
    b = 2
End Sub''')
    rt.run("Good")
    with pytest.raises(VBAError) as ei:
        rt.run("Bad")
    assert "Variable not defined" in ei.value.message
    assert ei.value.line == 9


def test_sub_not_defined():
    rt, book, eng, host = make('''
Sub Bad()
    DoesNotExist 1, 2
End Sub''')
    with pytest.raises(VBAError) as ei:
        rt.run("Bad")
    assert "Sub or Function not defined" in ei.value.message


def test_byref_byval_optional_paramarray_named():
    prints = run_print('''
Sub Inc(ByRef x As Long, Optional ByVal by As Long = 1)
    x = x + by
End Sub
Sub NoChange(ByVal x As Long)
    x = 100
End Sub
Function Total(ParamArray nums()) As Double
    Dim n
    For Each n In nums
        Total = Total + n
    Next
End Function
Function Greet(Optional name As Variant) As String
    If IsMissing(name) Then Greet = "Hi stranger" Else Greet = "Hi " & name
End Function
Sub FillArr(a() As Long)
    a(1) = 42
End Sub
Function Fact(n As Long) As Double
    If n <= 1 Then Fact = 1 Else Fact = n * Fact(n - 1)
End Function
Sub Test()
    Dim v As Long, arr(1 To 2) As Long
    v = 5
    Inc v
    Inc v, by:=10
    NoChange v
    Debug.Print v
    Debug.Print Total(1, 2, 3.5), Greet(), Greet("Bob")
    FillArr arr
    Debug.Print arr(1), Fact(10)
    Call Inc(v)
    Debug.Print v
End Sub''')
    assert prints[0].strip() == "16"
    assert prints[1].split() == ["6.5", "Hi", "stranger", "Hi", "Bob"]
    assert prints[2].split() == ["42", "3628800"]
    assert prints[3].strip() == "17"


def test_loops_and_control_flow():
    prints = run_print('''
Sub Test()
    Dim i As Integer, j As Integer, s As String, n As Long
    For i = 10 To 1 Step -3
        s = s & i & ","
    Next
    Debug.Print s
    i = 0
    Do While i < 5
        i = i + 1
        If i = 3 Then Exit Do
    Loop
    Debug.Print i
    Do
        i = i + 1
    Loop Until i >= 10
    Debug.Print i
    While i > 0
        i = i - 4
    Wend
    Debug.Print i
    For i = 1 To 3
        For j = 1 To 3
            If j = 2 Then Exit For
            n = n + 1
        Next j, i
    Debug.Print n
    i = 0
Again:
    i = i + 1
    If i < 4 Then GoTo Again
    Debug.Print i
    GoSub Sub1
    Debug.Print "back"
    Exit Sub
Sub1:
    Debug.Print "in gosub"
    Return
End Sub''')
    assert prints == ["10,7,4,1,", " 3 ", " 10 ", "-2 ", " 3 ", " 4 ", "in gosub", "back"]


def test_operators_and_semantics():
    prints = run_print('''
Sub Test()
    Debug.Print 7 \\ 2, 7 Mod 3, -7 Mod 3, 2 ^ 10, -2 ^ 2, 10 / 4
    Debug.Print "5" + 3, "a" + "b", 1 & 2, True + 1, Not 0, 5 And 3, 5 Or 3, 5 Xor 3
    Debug.Print CInt(2.5), CInt(3.5), CLng(-2.5), Round(2.5), Round(2.567, 2), Int(-2.5), Fix(-2.5)
    Debug.Print &HFF, &H8000, &O17, 1E3, IsNumeric("12.5"), IsNumeric("abc"), IsDate("1/31/2024")
    Debug.Print TypeName(1), TypeName(1.5), TypeName("s"), TypeName(Empty), TypeName(Null), TypeName(Nothing)
    Dim l As Long, d As Date, c As Currency
    Debug.Print TypeName(l), TypeName(d), TypeName(c), VarType(l), VarType("x")
    Debug.Print IIf(3 > 2, "yes", "no"), Choose(2, "a", "b", "c"), Switch(False, 1, True, 2)
    Debug.Print 0.1 + 0.2, 1 / 3, 1E+20, Abs(-3), Sgn(-2), Sqr(16)
    Debug.Print (1 < 2) = True, Null = 1, IsNull(Null + 1), Empty = 0, "" = Empty
End Sub''')
    assert prints[0].split() == ["3", "1", "-1", "1024", "-4", "2.5"]
    assert prints[1].split() == ["8", "ab", "12", "0", "-1", "1", "7", "6"]
    assert prints[2].split() == ["2", "4", "-2", "2", "2.57", "-3", "-2"]
    assert prints[3].split() == ["255", "-32768", "15", "1000", "True", "False", "True"]
    assert prints[4].split() == ["Integer", "Double", "String", "Empty", "Null", "Nothing"]
    assert prints[5].split() == ["Long", "Date", "Currency", "3", "8"]
    assert prints[6].split() == ["yes", "b", "2"]
    assert prints[7].split() == ["0.3", "0.333333333333333", "1E+20", "3", "-1", "4"]
    assert prints[8].split() == ["True", "Null", "True", "True", "True"]


def test_udt_and_enum_and_const():
    prints = run_print('''
Private Const RATE As Double = 0.5
Public Const GREETING = "Hi"
Enum Colour
    cRed = 1
    cGreen
    cBlue = 10
End Enum
Type Person
    Name As String
    Age As Integer
    Tags(1 To 2) As String
End Type
Sub Test()
    Dim p As Person, q As Person
    p.Name = "Ann": p.Age = 30
    p.Tags(1) = "x"
    q = p
    q.Name = "Bob"
    Debug.Print p.Name, q.Name, q.Age, q.Tags(1), RATE * 4, GREETING
    Debug.Print cGreen, Colour.cBlue
    Dim people(1 To 2) As Person
    people(2).Name = "Cy"
    With people(2)
        .Age = 5
    End With
    Debug.Print people(2).Name, people(2).Age
End Sub''')
    assert prints[0].split() == ["Ann", "Bob", "30", "x", "2", "Hi"]
    assert prints[1].split() == ["2", "10"]
    assert prints[2].split() == ["Cy", "5"]


# ---------------------------------------------------------------------------
# Excel object model
# ---------------------------------------------------------------------------
def test_range_value_2d_array():
    rt, book, eng, host = make('''
Sub Test()
    Dim data As Variant, r As Long, c As Long
    data = Range("A1:B3").Value
    Debug.Print UBound(data, 1), UBound(data, 2), data(2, 1), data(3, 2)
    For r = 1 To 3
        For c = 1 To 2
            If IsNumeric(data(r, c)) Then data(r, c) = data(r, c) * 10
        Next
    Next
    Range("D1:E3").Value = data
    Dim out(1 To 2, 1 To 2)
    out(1, 1) = "a": out(2, 2) = 4
    Range("G1").Resize(2, 2).Value = out
    Range("A5:C5").Value = Array(1, 2, 3)
    Range("A6:C7").Value = 7
    Debug.Print Range("A1:B3").Count, Range("A1:B3").Rows.Count, Range("A1:B3").Columns.Count
    Debug.Print Range("B2").Address, Range("B2").Address(False, False), Range("A1:B3").Address
    Debug.Print Range("A1").Offset(2, 1).Address, Range("B2").Resize(2, 3).Address
    Debug.Print Cells(2, "B").Address, Range("C3").Row, Range("C3").Column
End Sub''')
    sh = book.active
    for r in range(3):
        book.set_cell(sh, r, 0, str(r + 1))
        book.set_cell(sh, r, 1, "t%d" % r)
    rt.run("Test")
    assert host.prints[0].split() == ["3", "2", "2", "t2"]
    assert val(eng, sh, "D3") == 30.0 and val(eng, sh, "E2") == "t1"
    assert val(eng, sh, "G1") == "a" and val(eng, sh, "H2") == 4.0 and val(eng, sh, "H1") is None
    assert val(eng, sh, "C5") == 3.0 and val(eng, sh, "B7") == 7.0
    assert host.prints[1].split() == ["6", "3", "2"]
    assert host.prints[2].split() == ["$B$2", "B2", "$A$1:$B$3"]
    assert host.prints[3].split() == ["$B$3", "$B$2:$D$3"]
    assert host.prints[4].split() == ["$B$2", "3", "3"]


def test_formulas_and_worksheetfunction():
    rt, book, eng, host = make('''
Sub Test()
    Range("A1").Value = 10
    Range("A2").Value = 20
    Range("A3").Formula = "=SUM(A1:A2)"
    Debug.Print Range("A3").Value, Range("A3").Formula, Range("A3").HasFormula
    Debug.Print Application.WorksheetFunction.Sum(Range("A1:A3")), WorksheetFunction.Max(1, 5, 3)
    Range("C1:D3").Value = [{"a",1;"b",2;"c",3}]
    Debug.Print WorksheetFunction.VLookup("b", Range("C1:D3"), 2, False)
    Debug.Print Application.Average(Range("A1:A2")), Evaluate("2*3"), [A1] + 1
    Debug.Print IsError(Application.VLookup("zz", Range("C1:D3"), 2, False))
    Range("B1:B2").FormulaR1C1 = "=RC[-1]*2"
    Debug.Print Range("B2").Formula, Range("B2").Value, Range("B2").FormulaR1C1
    On Error Resume Next
    Dim x
    x = WorksheetFunction.VLookup("zz", Range("C1:D3"), 2, False)
    Debug.Print Err.Number, Err.Description
End Sub''')
    rt.run("Test")
    p = host.prints
    assert p[0].split() == ["30", "=SUM(A1:A2)", "True"]
    assert p[1].split() == ["60", "5"]
    assert p[2].strip() == "2"
    assert p[3].split() == ["15", "6", "11"]
    assert p[4].strip() == "True"
    assert p[5].split() == ["=A2*2", "40", "=RC[-1]*2"]
    assert p[6].split()[0] == "1004" and "Unable to get the VLookup property" in p[6]


def test_worksheets_add_rename_loop():
    rt, book, eng, host = make('''
Sub Test()
    Dim ws As Worksheet, names As String
    Set ws = Worksheets.Add(After:=Worksheets(Worksheets.Count))
    ws.Name = "Report"
    ws.Range("A1").Value = "hello"
    Sheets.Add.Name = "Second"
    For Each ws In ThisWorkbook.Worksheets
        names = names & ws.Name & ";"
    Next
    Debug.Print names, Worksheets.Count, Sheets("Report").Index
    Debug.Print Worksheets("Report").Range("A1").Value, ActiveSheet.Name
    Worksheets("Sheet1").Activate
    Debug.Print ActiveSheet.Name
    Application.DisplayAlerts = False
    Worksheets("Second").Delete
    Debug.Print Worksheets.Count
    On Error Resume Next
    Worksheets("Report").Name = "Sheet1"
    Debug.Print Err.Number
    Set ws = Worksheets("Nope")
    Debug.Print Err.Number
End Sub''')
    rt.run("Test")
    p = host.prints
    assert p[0].split() == ["Sheet1;Second;Report;", "3", "3"]
    assert p[1].split() == ["hello", "Second"]
    assert p[2] == "Sheet1"
    assert p[3].strip() == "2"
    assert p[4].strip() == "1004"
    assert p[5].strip() == "9"
    assert [s.name for s in book.sheets] == ["Sheet1", "Report"]


def test_worksheet_change_event_via_notify_change():
    sheet_code = '''
Private Sub Worksheet_Change(ByVal Target As Range)
    If Not Intersect(Target, Me.Range("A1:A10")) Is Nothing Then
        Application.EnableEvents = False
        Target.Offset(0, 1).Value = "changed " & Target.Address
        Application.EnableEvents = True
    End If
End Sub'''
    rt, book, eng, host = make(None, [{"name": "Sheet1", "type": "document", "code": sheet_code}])
    sh = book.active
    book.set_cell(sh, 2, 0, "x")
    rt.notify_change(sh, 2, 0, 2, 0)
    assert val(eng, sh, "B3") == "changed $A$3"
    rt.notify_change(sh, 20, 0, 20, 0)
    assert val(eng, sh, "B21") is None
    assert rt.list_macros() == []


def test_vba_write_fires_change_event():
    sheet_code = '''
Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.Address = "$Z$1" Then Exit Sub
    Me.Range("Z1").Value = Me.Range("Z1").Value + 1
End Sub'''
    rt, book, eng, host = make('''
Sub W()
    Range("A1").Value = 5
    Application.EnableEvents = False
    Range("A2").Value = 6
    Application.EnableEvents = True
End Sub''', [{"name": "Sheet1", "type": "document", "code": sheet_code}])
    rt.run("W")
    assert val(eng, book.active, "Z1") == 1.0


def test_workbook_open_fire_event():
    wb = '''
Private Sub Workbook_Open()
    MsgBox "Welcome", vbInformation, "Hi"
    Sheets(1).Range("A1").Value = "opened"
End Sub
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Cancel = True
End Sub'''
    rt, book, eng, host = make(None, [{"name": "ThisWorkbook", "type": "document", "code": wb}])
    rt.fire_event("ThisWorkbook", "Workbook_Open")
    assert host.msgs == [("Welcome", 64, "Hi")]
    assert val(eng, book.active, "A1") == "opened"
    rt.fire_event("ThisWorkbook", "Workbook_NotThere")
    res = rt.fire_event("ThisWorkbook", "Workbook_BeforeClose", [False])
    assert res == [True]
    rt.app.enableevents = False
    rt.fire_event("ThisWorkbook", "Workbook_Open")
    assert len(host.msgs) == 1


def test_msgbox_inputbox_return_values():
    host = RecHost(msg_answers=[6, 7], input_answers=["42", None])
    rt, book, eng, _ = make('''
Sub Test()
    If MsgBox("Continue?", vbYesNo + vbQuestion, "Q") = vbYes Then Debug.Print "yes"
    Dim ans As VbMsgBoxResult
    ans = MsgBox("Again?", vbYesNo)
    If ans = vbNo Then Debug.Print "no"
    Dim s As String
    s = InputBox("Number?", "Ask", "1")
    Debug.Print CLng(s) * 2
    s = InputBox("Cancel me")
    Debug.Print "[" & s & "]"
End Sub''', host=host)
    rt.run("Test")
    assert host.prints == ["yes", "no", " 84 ", "[]"]
    assert host.msgs[0] == ("Continue?", 36, "Q")


def test_format_and_dates():
    prints = run_print('''
Sub Test()
    Dim d As Date
    d = DateSerial(2024, 1, 31) + TimeSerial(13, 5, 9)
    Debug.Print Format(d, "dd/mm/yyyy"), Format(d, "yyyy-mm-dd"), Format(d, "hh:mm:ss")
    Debug.Print Format(d, "mmmm"), Format(d, "dddd"), Format(d, "Short Date"), Format(d, "Long Date")
    Debug.Print Format(1234.567, "0.00"), Format(1234.567, "#,##0"), Format(0.256, "0%")
    Debug.Print Format(1234.5, "Currency"), Format(1234.5, "Fixed"), Format(1234.5, "Standard"), Format(0.5, "Percent")
    Debug.Print Year(d), Month(d), Day(d), Hour(d), Minute(d), Second(d), Weekday(d)
    Debug.Print DateAdd("m", 1, #1/31/2024#), DateDiff("d", #1/1/2024#, #3/1/2024#), DatePart("q", d)
    Debug.Print MonthName(2), WeekdayName(1), CDate("2024-03-05"), DateValue("March 5, 2024")
    Debug.Print FormatNumber(1234.5, 1), FormatCurrency(-5), FormatPercent(0.125, 1)
    Debug.Print Format(#1/2/2024#, "General Date"), #10:30:00 PM#
    Debug.Print Format(5, "000"), Format(-3.14159, "0.00"), Format(0, "#,##0.00")
End Sub''')
    assert prints[0].split() == ["31/01/2024", "2024-01-31", "13:05:09"]
    assert prints[1].split() == ["January", "Wednesday", "1/31/2024", "Wednesday,", "January", "31,", "2024"]
    assert prints[2].split() == ["1234.57", "1,235", "26%"]
    assert prints[3].split() == ["$1,234.50", "1234.50", "1,234.50", "50.00%"]
    assert prints[4].split() == ["2024", "1", "31", "13", "5", "9", "4"]
    assert prints[5].split() == ["2/29/2024", "60", "1"]
    assert prints[6].split() == ["February", "Sunday", "3/5/2024", "3/5/2024"]
    assert prints[7].split() == ["1,234.5", "($5.00)", "12.5%"]
    assert prints[8].split() == ["1/2/2024", "10:30:00", "PM"]
    assert prints[9].split() == ["005", "-3.14", "0.00"]


def test_dates_written_to_cells():
    rt, book, eng, host = make('''
Sub Test()
    Range("A1").Value = DateSerial(2024, 1, 31)
    Debug.Print TypeName(Range("A1").Value), Range("A1").Value2, Range("A1").Text
End Sub''')
    rt.run("Test")
    assert host.prints[0].split() == ["Date", "45322", "1/31/2024"]


def test_file_io(tmp_path):
    rt, book, eng, host = make('''
Sub Test()
    Dim f As Integer, line As String, a, b, n As Long
    f = FreeFile
    Open "out.txt" For Output As #f
    Print #f, "hello"; " world"
    Print #f, 42
    Write #f, "x,y", 3.5, True
    Close #f
    Open "out.txt" For Append As #1
    Print #1, "tail"
    Close #1
    Open "out.txt" For Input As #1
    Do While Not EOF(1)
        Line Input #1, line
        n = n + 1
        Debug.Print line
    Loop
    Close #1
    Debug.Print n, Dir("out.txt"), FileLen("out.txt") > 0
    Kill "out.txt"
    Debug.Print "[" & Dir("out.txt") & "]"
    Open "nums.csv" For Output As #2
    Write #2, "a", 1
    Close #2
    Open "nums.csv" For Input As #2
    Input #2, a, b
    Close #2
    Debug.Print a, b + 1
    Dim fso As Object, ts As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.CreateTextFile("t.txt", True)
    ts.WriteLine "line1"
    ts.Close
    Debug.Print fso.FileExists("t.txt"), fso.OpenTextFile("t.txt").ReadLine
End Sub''')
    rt.cur_dir = str(tmp_path)
    rt.run("Test")
    p = host.prints
    assert p[0] == "hello world"
    assert p[1] == " 42 "
    assert p[2] == '"x,y",3.5,#TRUE#'
    assert p[3] == "tail"
    assert p[4].split() == ["4", "out.txt", "True"]
    assert p[5] == "[]"
    assert p[6].split() == ["a", "2"]
    assert p[7].split() == ["True", "line1"]


def test_find_end_currentregion_sort_copy():
    rt, book, eng, host = make('''
Sub Test()
    Dim f As Range
    Set f = Range("A1:A10").Find(What:="pear", LookAt:=xlWhole)
    Debug.Print f.Address, Range("A1").End(xlDown).Row, Range("A1").CurrentRegion.Address
    Debug.Print Range("A1:A10").Find("nothing here") Is Nothing
    Range("A1:B4").Sort Key1:=Range("B1"), Order1:=xlDescending, Header:=xlYes
    Range("A1:B4").Copy Destination:=Range("D1")
    Range("C2").Formula = "=B2*2"
    Range("C2").Copy Range("C3:C4")
    Range("A1:B1").Copy
    Range("A7").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    Rows(1).Insert
    Debug.Print Range("A1").Value = "", Range("A2").Value
    Rows(1).Delete
    Range("A1").Value = "Fruit"
End Sub''')
    sh = book.active
    rows = [("Fruit", "Qty"), ("apple", "3"), ("pear", "10"), ("fig", "7")]
    for r, (a, b) in enumerate(rows):
        book.set_cell(sh, r, 0, a)
        book.set_cell(sh, r, 1, b)
    rt.run("Test")
    p = host.prints
    assert p[0].split() == ["$A$3", "4", "$A$1:$B$4"]
    assert p[1].strip() == "True"
    assert [val(eng, sh, "A%d" % i) for i in range(2, 5)] == ["pear", "fig", "apple"]
    assert val(eng, sh, "E2") == 10.0 and val(eng, sh, "D1") == "Qty" or val(eng, sh, "D1") == "Fruit"
    assert sh.get_raw(3, 2) == "=B4*2"
    assert val(eng, sh, "C4") == 6.0
    assert val(eng, sh, "A7") == "Fruit" and val(eng, sh, "B7") == "Qty"
    assert p[2].split() == ["True", "Fruit"]


def test_selection_and_activecell():
    rt, book, eng, host = make('''
Sub Test()
    Range("B2:C3").Select
    Selection.Value = 1
    ActiveCell.Value = 9
    Debug.Print Selection.Address, ActiveCell.Address
End Sub''')
    rt.run("Test")
    assert host.prints == ["$B$2:$C$3 $B$2"] or host.prints[0].split() == ["$B$2:$C$3", "$B$2"]
    assert val(eng, book.active, "B2") == 9.0 and val(eng, book.active, "C3") == 1.0


def test_application_properties():
    rt, book, eng, host = make('''
Sub Test()
    Application.ScreenUpdating = False
    Application.StatusBar = "Working..."
    Application.Calculation = xlCalculationManual
    Debug.Print Application.Version, Application.ScreenUpdating, Application.Name
    Application.StatusBar = False
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Debug.Print ThisWorkbook.Name, Application.Run("Twice", 21)
End Sub
Function Twice(x)
    Twice = x * 2
End Function''')
    rt.run("Test")
    assert host.status == ["Working...", ""]
    assert host.prints[0].split() == ["16.0", "False", "Microsoft", "Excel"]
    assert host.prints[1].split()[-1] == "42"


def test_list_macros_and_run_variants():
    rt, book, eng, host = make('''
Sub Public1()
End Sub
Private Sub Hidden()
End Sub
Sub WithArg(x)
End Sub
Function F()
    F = 3
End Function''', [{"name": "Sheet1", "type": "document", "code": "Sub SheetMacro()\nEnd Sub\nPrivate Sub Worksheet_Activate()\nEnd Sub"},
                  {"name": "UserForm1", "type": "form", "code": "Private Sub CommandButton1_Click()\n Me.Hide\nEnd Sub"}])
    assert rt.list_macros() == ["Module1.Public1", "Sheet1.SheetMacro"]
    assert rt.run("Module1.F") == 3
    assert rt.run("f") == 3
    with pytest.raises(VBAError):
        rt.run("Nope")


def test_userform_not_supported_and_api_declare():
    rt, book, eng, host = make('''
#If VBA7 Then
Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
#Else
Private Declare Function GetTickCount Lib "kernel32" () As Long
#End If
Sub ShowForm()
    UserForm1.Show
End Sub
Sub CallApi()
    Debug.Print GetTickCount()
End Sub
Sub UseShell()
    Shell "calc.exe"
End Sub''', [{"name": "UserForm1", "type": "form", "code": "Private Sub UserForm_Initialize()\nEnd Sub"}])
    with pytest.raises(VBAError) as ei:
        rt.run("ShowForm")
    assert "UserForms are not supported" in ei.value.message
    with pytest.raises(VBAError) as ei:
        rt.run("CallApi")
    assert "Windows API calls are not supported" in ei.value.message
    with pytest.raises(VBAError) as ei:
        rt.run("UseShell")
    assert ei.value.vba_number == 5 and "Shell" in ei.value.message


def test_parse_errors_returned_not_raised():
    book = Workbook()
    rt = VBARuntime(book, FormulaEngine(book), RecHost())
    errs = rt.load_modules([{"name": "M", "type": "standard",
                             "code": "Sub A()\n  If x Then\nEnd Sub\nSub B()\n  Debug.Print 1\nEnd Sub"}])
    assert errs and errs[0].module == "M"
    rt.run("B")
    assert rt.host.prints == [" 1 "]


def test_end_statement_and_stop():
    rt, book, eng, host = make('''
Dim counter As Long
Sub Test()
    counter = counter + 1
    Debug.Print counter
    End
    Debug.Print "never"
End Sub
Sub S()
    Stop
End Sub''')
    rt.run("Test")
    rt.run("Test")
    assert host.prints == [" 1 ", " 1 "]
    with pytest.raises(VBAError) as ei:
        rt.run("S")
    assert "Stop" in ei.value.message


def test_request_stop():
    class StopHost(RecHost):
        def process_events(self):
            self.events += 1
            if self.events == 3:
                rt.request_stop()
    host = StopHost()
    rt, book, eng, _ = make('''
Sub Forever()
    Do
    Loop
End Sub''', host=host)
    with pytest.raises(VBAError) as ei:
        rt.run("Forever")
    assert "interrupted" in ei.value.message


def test_immediate_window():
    rt, book, eng, host = make('''
Function Sq(x)
    Sq = x * x
End Function''')
    assert rt.evaluate_immediate("? Sq(7) + 1").strip() == "50"
    assert rt.evaluate_immediate("Range(\"A1\").Value = 3") == ""
    assert val(eng, book.active, "A1") == 3.0
    assert rt.evaluate_immediate("?Range(\"A1\").Value * 2").strip() == "6"


def test_static_and_module_vars():
    prints = run_print('''
Private total As Long
Function Counter() As Long
    Static n As Long
    n = n + 1
    Counter = n
End Function
Sub Test()
    Counter
    Counter
    Debug.Print Counter()
    total = total + 5
    Debug.Print total
End Sub''')
    assert prints == [" 3 ", " 5 "]


def test_regexp_and_misc_objects():
    prints = run_print('''
Sub Test()
    Dim re As Object, m
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "(\\d+)"
    re.Global = True
    For Each m In re.Execute("a1 b22 c333")
        Debug.Print m.Value, m.SubMatches(0)
    Next
    Debug.Print re.Replace("a1 b22", "<$1>"), re.Test("xyz")
End Sub''')
    assert prints[0].split() == ["1", "1"]
    assert prints[2].split() == ["333", "333"]
    assert prints[3].split() == ["a<1>", "b<22>", "False"]


def test_performance_loop_and_cell_writes():
    rt, book, eng, host = make('''
Sub Arith()
    Dim i As Long, s As Double
    For i = 1 To 100000
        s = s + i Mod 7 * 2
    Next
    Range("A1").Value = s
End Sub
Sub Cellz()
    Dim i As Long
    For i = 1 To 10000
        Cells(i, 2).Value = i
    Next
    Range("C1").Formula = "=SUM(B1:B10000)"
End Sub''')
    t = time.time()
    rt.run("Arith")
    t1 = time.time() - t
    t = time.time()
    rt.run("Cellz")
    t2 = time.time() - t
    assert t1 < 5 and t2 < 10, (t1, t2)
    assert val(eng, book.active, "C1") == 50005000.0


def test_line_continuation_comments_colon():
    prints = run_print('''
Sub Test() ' comment
    Dim s As String: s = "a" & _
        "b" & _
        "c"   ' trailing
    Rem a remark
    Debug.Print s: Debug.Print Len(s)
10  Debug.Print "numbered"
End Sub''')
    assert prints == ["abc", " 3 ", "numbered"]


def test_unqualified_range_in_sheet_module():
    sheet2 = '''
Sub FillMe()
    Range("A1").Value = "on sheet2"
End Sub'''
    book = Workbook()
    book.add_sheet()
    eng = FormulaEngine(book)
    rt = VBARuntime(book, eng, RecHost())
    rt.load_modules([{"name": "Sheet2", "type": "document", "code": sheet2}])
    rt.run("Sheet2.FillMe")
    assert eng.value(book.sheets[1], 0, 0) == "on sheet2"
    assert eng.value(book.sheets[0], 0, 0) is None


def test_common_idioms_delete_rows_and_codenames():
    rt, book, eng, host = make('''
Sub Clean()
    Dim ws As Worksheet, lastRow As Long, i As Long
    Set ws = ThisWorkbook.Sheets("Sheet1")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    For i = lastRow To 1 Step -1
        If ws.Cells(i, 1).Value = "" Or ws.Cells(i, 1).Value = "x" Then
            ws.Rows(i).Delete
        ElseIf IsNumeric(ws.Cells(i, 1).Value) Then
            ws.Cells(i, 2).Value = ws.Cells(i, 1).Value ^ 2
        Else
            ws.Cells(i, 2).Value = UCase(ws.Cells(i, 1).Value)
        End If
    Next i
    With Sheet1
        .Range("C1:C" & .Cells(.Rows.Count, 1).End(xlUp).Row).Interior.ColorIndex = 6
        Debug.Print WorksheetFunction.CountIf(.Range("A:A"), ">1"), .UsedRange.Address
    End With
End Sub''')
    sh = book.active
    for i, v in enumerate(["1", "x", "", "abc", "3", "x"]):
        if v:
            book.set_cell(sh, i, 0, v)
    rt.run("Clean")
    assert [sh.get_raw(r, 0) for r in range(3)] == ["1", "abc", "3"]
    assert val(eng, sh, "B3") == 9.0 and val(eng, sh, "B2") == "ABC"
    assert sh.get_style(2, 2)["fill"] == "#FFFF00"
    assert host.prints[0].split() == ["1", "$A$1:$B$3"]


def test_selection_change_and_before_save_events():
    code = '''
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Debug.Print "sel " & Target.Address
End Sub
Private Sub Worksheet_Activate()
    Debug.Print "activated"
End Sub'''
    wb = '''
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    Debug.Print "saving", SaveAsUI
End Sub'''
    rt, book, eng, host = make(None, [{"name": "Sheet1", "type": "document", "code": code},
                                      {"name": "ThisWorkbook", "type": "document", "code": wb}])
    sh = book.active
    rt.fire_event(sh.code_name, "Worksheet_SelectionChange", (rt.make_range(sh, 1, 1, 2, 2),))
    rt.fire_event("Sheet1", "Worksheet_Activate", ())
    rt.fire_event("ThisWorkbook", "Workbook_BeforeSave", ())
    assert host.prints[0] == "sel $B$2:$C$3"
    assert host.prints[1] == "activated"
    assert host.prints[2].split() == ["saving", "False"]
