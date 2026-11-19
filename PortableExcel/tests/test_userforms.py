"""UserForm support in the VBA runtime, tested with a headless fake FormView."""
import os
import sys

import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from portable_excel.model import Workbook, parse_a1  # noqa: E402
from portable_excel.engine import FormulaEngine  # noqa: E402
from portable_excel.forms import new_form, new_control  # noqa: E402
from portable_excel.vba import VBARuntime, VBAHost, VBAError  # noqa: E402


class FakeView:
    def __init__(self, host, inst):
        self.host = host
        self.inst = inst
        self.calls = []
        self.closed = False
        self.hidden = False

    def _active(self):
        return not self.closed and not self.hidden

    def show(self, modal):
        self.calls.append(("show", modal))
        self.hidden = False
        if not modal:
            return
        script = self.host.scripts.pop(0) if self.host.scripts else []
        for action in script:
            if not self._active():
                break
            action(self.inst)
        if self._active():
            self.inst.request_close(0)
            if self._active():          # QueryClose cancelled: user then clicks a Hide button
                self.inst.hide()

    def hide(self):
        self.calls.append(("hide",))
        self.hidden = True

    def close(self):
        self.calls.append(("close",))
        self.closed = True

    def set_prop(self, ctrl, prop, value):
        self.calls.append(("set_prop", ctrl, prop, value))

    def add_control(self, ctrl):
        self.calls.append(("add_control", ctrl["name"], ctrl["type"]))

    def remove_control(self, name):
        self.calls.append(("remove_control", name))

    def set_focus(self, name):
        self.calls.append(("set_focus", name))

    def repaint(self):
        self.calls.append(("repaint",))


class FormHost(VBAHost):
    def __init__(self, scripts=None):
        super().__init__()
        self.scripts = list(scripts or [])
        self.views = []
        self.prints = []
        self.msgs = []

    def create_form_view(self, inst):
        v = FakeView(self, inst)
        self.views.append(v)
        return v

    def debug_print(self, text):
        self.prints.append(text)

    def msgbox(self, prompt, buttons=0, title=""):
        self.msgs.append(prompt)
        return 1


# user actions -------------------------------------------------------------
def type_text(ctrl, text):
    def act(inst):
        inst.user_set(ctrl, "value", text)
        inst.fire(ctrl, "Change")
    return act


def click(ctrl):
    def act(inst):
        inst.fire(ctrl, "Click")
    return act


def choose(ctrl, index):
    def act(inst):
        inst.user_set(ctrl, "list_index", index)
        inst.fire(ctrl, "Change")
        inst.fire(ctrl, "Click")
    return act


def entry_form():
    f = new_form("Data Entry")
    f["controls"] = [
        new_control("Label", "lblName", caption="Name"),
        new_control("TextBox", "txtName"),
        new_control("TextBox", "txtAge"),
        new_control("ComboBox", "cboDept"),
        new_control("CheckBox", "chkActive", caption="Active"),
        new_control("OptionButton", "optA", group_name="g"),
        new_control("OptionButton", "optB", group_name="g"),
        new_control("CommandButton", "cmdOK", caption="OK", default=True),
        new_control("CommandButton", "cmdCancel", caption="Cancel", cancel=True),
    ]
    return f


FORM_CODE = '''
Option Explicit
Public Cancelled As Boolean

Private Sub UserForm_Initialize()
    cboDept.AddItem "Sales"
    cboDept.AddItem "IT"
    cboDept.AddItem "HR"
    cboDept.ListIndex = 0
    Me.Caption = "New employee"
    Debug.Print "init"
End Sub

Private Sub UserForm_Activate()
    Debug.Print "activate"
End Sub

Private Sub cmdOK_Click()
    Dim r As Long
    If Trim(txtName.Text) = "" Then
        MsgBox "Name required"
        Exit Sub
    End If
    r = Sheets(1).Cells(Rows.Count, 1).End(xlUp).Row + 1
    With Sheets(1)
        .Cells(r, 1).Value = Me.txtName.Value
        .Cells(r, 2).Value = CLng(txtAge.Value)
        .Cells(r, 3).Value = cboDept.Value
        .Cells(r, 4).Value = chkActive.Value
    End With
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Cancelled = True
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Debug.Print "queryclose", CloseMode
End Sub

Private Sub UserForm_Terminate()
    Debug.Print "terminate"
End Sub
'''


def make(code, form_code=FORM_CODE, form=None, scripts=None, extra=None):
    book = Workbook()
    eng = FormulaEngine(book)
    host = FormHost(scripts)
    rt = VBARuntime(book, eng, host)
    mods = [{"name": "Module1", "type": "standard", "code": code},
            {"name": "UserForm1", "type": "form", "code": form_code, "form": form or entry_form()}]
    mods += extra or []
    errs = rt.load_modules(mods)
    assert errs == [], [str(e) for e in errs]
    return rt, book, eng, host


def val(eng, book, ref, sheet=0):
    r, c = parse_a1(ref)
    return eng.value(book.sheets[sheet], r, c)


def test_data_entry_ok_writes_cells_and_unloads():
    rt, book, eng, host = make('''
Sub AddEmployee()
    Sheets(1).Range("A1").Value = "Name"
    UserForm1.Show
    Debug.Print "after show"
End Sub''', scripts=[[type_text("txtName", "Ann"), type_text("txtAge", "34"),
                      choose("cboDept", 1),
                      lambda i: i.user_set("chkActive", "value", True),
                      click("cmdOK")]])
    rt.run("AddEmployee")
    assert val(eng, book, "A2") == "Ann"
    assert val(eng, book, "B2") == 34.0
    assert val(eng, book, "C2") == "IT"
    assert val(eng, book, "D2") is True
    assert host.prints == ["init", "activate", "queryclose    1 ", "terminate", "after show"] or \
        [p.split() for p in host.prints] == [["init"], ["activate"], ["queryclose", "1"],
                                             ["terminate"], ["after", "show"]]
    v = host.views[0]
    assert ("close",) in v.calls
    assert v.inst.form["caption"] == "New employee"
    assert rt.form_defaults == {}


def test_validation_msgbox_then_cancel_hides_and_keeps_values():
    rt, book, eng, host = make('''
Sub Go()
    UserForm1.Show
    Debug.Print UserForm1.Cancelled, UserForm1.txtName.Text, UserForm1.cboDept.Text
    UserForm1.txtName.Text = "kept"
    Unload UserForm1
    Debug.Print "[" & UserForm1.txtName.Text & "]", UserForm1.Cancelled
End Sub''', scripts=[[click("cmdOK"), type_text("txtName", "Bob"), click("cmdCancel")]])
    rt.run("Go")
    assert host.msgs == ["Name required"]
    p = [x.split() for x in host.prints]
    assert p[2] == ["True", "Bob", "Sales"]
    # after Unload the default instance is re-created and re-initialised
    assert p[-1] == ["[]", "False"]
    assert p.count(["init"]) == 2


def test_hide_show_persists_and_modeless():
    rt, book, eng, host = make('''
Sub Go()
    Load UserForm1
    UserForm1.txtName.Value = "first"
    UserForm1.Show vbModeless
    UserForm1.Hide
    UserForm1.Show 0
    Debug.Print UserForm1.txtName.Value, UserForm1.Visible
End Sub''')
    rt.run("Go")
    assert host.prints[-1].split() == ["first", "True"]
    v = host.views[0]
    assert [c for c in v.calls if c[0] in ("show", "hide")] == [("show", False), ("hide",), ("show", False)]
    assert len(host.views) == 1


def test_query_close_cancel():
    code = '''
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then Cancel = True
    Debug.Print "qc", CloseMode
End Sub'''
    rt, book, eng, host = make('''
Sub Go()
    UserForm1.Show
    Debug.Print "loaded", UserForm1.Visible
End Sub''', form_code=code, scripts=[[]])
    rt.run("Go")
    assert host.prints[0].split() == ["qc", "0"]
    assert host.prints[1].split() == ["loaded", "False"]
    inst = rt.form_defaults["userform1"]
    assert inst.request_close(0) is False
    assert inst.loaded


def test_new_instances_are_separate():
    rt, book, eng, host = make('''
Sub Go()
    Dim a As New UserForm1, b As UserForm1
    Set b = New UserForm1
    a.txtName.Text = "A"
    b.txtName.Text = "B"
    Debug.Print a.txtName.Text, b.txtName.Text, TypeName(a), a Is b
    Unload a
    Unload b
End Sub''', form_code="Private Sub UserForm_Terminate()\n Debug.Print \"bye\"\nEnd Sub")
    rt.run("Go")
    assert host.prints[0].split() == ["A", "B", "UserForm1", "False"]
    assert host.prints[1:] == ["bye", "bye"]


def test_controls_collection_add_remove():
    code = '''
Private Sub UserForm_Initialize()
    Dim c As Control, t As MSForms.TextBox, n As Long
    Set t = Me.Controls.Add("Forms.TextBox.1", "txtDyn", True)
    t.Left = 10: t.Top = 20: t.Text = "dyn"
    For Each c In Me.Controls
        n = n + 1
    Next
    Debug.Print n, Me.Controls.Count, Controls("txtDyn").Text, TypeName(t), Me.Controls(0).Name
    Me.Controls.Remove "txtDyn"
    Debug.Print Controls.Count
End Sub'''
    rt, book, eng, host = make('''
Sub Go()
    Load UserForm1
End Sub''', form_code=code)
    rt.run("Go")
    assert host.prints[0].split() == ["10", "10", "dyn", "TextBox", "lblName"]
    assert host.prints[1].strip() == "9"


def test_listbox_multiselect_and_columns():
    f = new_form()
    f["controls"] = [new_control("ListBox", "lst", multi_select=1, column_count=2)]
    code = '''
Private Sub UserForm_Initialize()
    Dim arr(0 To 2, 0 To 1)
    arr(0, 0) = "a": arr(0, 1) = 1
    arr(1, 0) = "b": arr(1, 1) = 2
    arr(2, 0) = "c": arr(2, 1) = 3
    lst.List = arr
    lst.Selected(0) = True
    lst.Selected(2) = True
End Sub
Function Picked() As String
    Dim i As Long
    For i = 0 To lst.ListCount - 1
        If lst.Selected(i) Then Picked = Picked & lst.List(i, 0) & lst.Column(1, i) & ";"
    Next
End Function'''
    rt, book, eng, host = make('''
Sub Go()
    Debug.Print UserForm1.Picked(), UserForm1.lst.ListCount, IsNull(UserForm1.lst.Value)
    UserForm1.lst.Selected(0) = False
    Debug.Print UserForm1.Picked()
    UserForm1.lst.AddItem "d"
    UserForm1.lst.List(3, 1) = 4
    Debug.Print UserForm1.lst.List(3, 0) & UserForm1.lst.List(3, 1), UBound(UserForm1.lst.List, 2)
End Sub''', form_code=code, form=f)
    rt.run("Go")
    assert host.prints[0].split() == ["a1;c3;", "3", "True"]
    assert host.prints[1].split() == ["c3;"]
    assert host.prints[2].split() == ["d4", "1"]


def test_rowsource_and_controlsource():
    f = new_form()
    f["controls"] = [new_control("ComboBox", "cbo", row_source="Sheet1!A1:A3"),
                     new_control("TextBox", "txt", control_source="Sheet1!C1")]
    code = '''
Private Sub txt_AfterUpdate()
    Debug.Print "after"
End Sub'''
    book = Workbook()
    eng = FormulaEngine(book)
    host = FormHost([[type_text("txt", "new"), lambda i: i.fire("txt", "AfterUpdate")]])
    rt = VBARuntime(book, eng, host)
    for i, v in enumerate(["red", "green", "blue"]):
        book.set_cell(book.sheets[0], i, 0, v)
    book.set_cell(book.sheets[0], 0, 2, "old")
    rt.load_modules([{"name": "Module1", "type": "standard", "code": '''
Sub Go()
    Debug.Print UserForm1.cbo.ListCount, UserForm1.cbo.List(2), UserForm1.txt.Text
    UserForm1.cbo.Value = "green"
    Debug.Print UserForm1.cbo.ListIndex
    UserForm1.Show
End Sub'''}, {"name": "UserForm1", "type": "form", "code": code, "form": f}])
    rt.run("Go")
    assert host.prints[0].split() == ["3", "blue", "old"]
    assert host.prints[1].strip() == "1"
    assert host.prints[2] == "after"
    assert eng.value(book.sheets[0], 0, 2) == "new"


def test_option_buttons_and_change_events_from_code():
    f = new_form()
    f["controls"] = [new_control("OptionButton", "o1", value=True),
                     new_control("OptionButton", "o2"),
                     new_control("Frame", "fr"),
                     new_control("OptionButton", "o3", parent="fr", value=True),
                     new_control("TextBox", "t"),
                     new_control("SpinButton", "sp", min=0, max=10)]
    code = '''
Private Sub o2_Click()
    Debug.Print "o2 click"
End Sub
Private Sub t_Change()
    Debug.Print "t change " & t.Text
End Sub
Private Sub sp_Change()
    t.Text = sp.Value
End Sub'''
    rt, book, eng, host = make('''
Sub Go()
    With UserForm1
        .o2.Value = True
        Debug.Print .o1.Value, .o2.Value, .o3.Value
        .sp.Value = 50
        Debug.Print .sp.Value
        .t.Text = "10"
    End With
End Sub''', form_code=code, form=f)
    rt.run("Go")
    assert host.prints[0] == "o2 click"
    assert host.prints[1].split() == ["False", "True", "True"]
    assert host.prints[2] == "t change 10"
    assert host.prints[3].strip() == "10"
    assert len(host.prints) == 4      # setting the same text again fires no Change


def test_event_error_propagates_to_gui_without_killing_macro():
    code = '''
Private Sub cmdOK_Click()
    Dim x As Integer
    x = 1 / 0
End Sub
Private Sub cmdCancel_Click()
    Unload Me
End Sub'''
    errors = []

    def click_ok(inst):
        try:
            inst.fire("cmdOK", "Click")
        except VBAError as e:
            errors.append(e)
    rt, book, eng, host = make('''
Sub Go()
    UserForm1.Show
    Debug.Print "macro continues"
End Sub''', form_code=code, scripts=[[click_ok, click("cmdCancel")]])
    rt.run("Go")
    assert len(errors) == 1 and errors[0].vba_number == 11
    assert errors[0].module == "UserForm1" and errors[0].line == 4
    assert host.prints == ["macro continues"]


def test_end_in_event_closes_forms_and_stops_macro():
    code = '''
Private Sub cmdOK_Click()
    End
End Sub'''
    rt, book, eng, host = make('''
Sub Go()
    UserForm1.Show
    Debug.Print "never"
End Sub''', form_code=code, scripts=[[click("cmdOK")]])
    rt.run("Go")
    assert host.prints == []
    assert ("close",) in host.views[0].calls
    assert rt.live_forms == []


def test_keypress_byref_and_form_properties():
    code = '''
Private Sub txtAge_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
End Sub
Private Sub txtName_KeyPress(KeyAscii As Integer)
    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub
Private Sub txtName_Exit(Cancel As Boolean)
    Cancel = (txtName.Text = "")
End Sub'''
    rt, book, eng, host = make('''
Sub Go()
    With UserForm1
        .Width = 300
        .BackColor = RGB(255, 0, 0)
        .Font.Bold = True
        .txtName.ForeColor = vbBlue
        .txtName.Font.Size = 12
        .txtName.SetFocus
        Debug.Print .Width, .BackColor = vbRed, .ActiveControl.Name, .txtName.ForeColor
        .Tag = "t"
        Debug.Print .Tag, .Controls.Count, .lblName.Caption
    End With
End Sub''', form_code=code)
    rt.run("Go")
    inst = rt.form_defaults["userform1"]
    assert inst.fire("txtName", "KeyPress", [65]) == [0]
    assert inst.fire("txtName", "KeyPress", [50]) == [50]
    assert inst.fire("txtName", "Exit", [False]) == [True]
    assert inst.fire("cmdOK", "Click") == []
    assert host.prints[0].split() == ["300", "True", "txtName", "16711680"]
    assert host.prints[1].split() == ["t", "9", "Name"]
    assert inst.form["width"] == 300.0 and inst.form["back_color"] == "#FF0000"
    assert inst.find_ctrl("txtName")["fore_color"] == "#0000FF"
    assert inst.find_ctrl("txtName")["font_size"] == 12.0


def test_multipage_and_user_set_no_echo():
    f = new_form()
    f["controls"] = [new_control("MultiPage", "mp"), new_control("TextBox", "t")]
    rt, book, eng, host = make('''
Sub Go()
    UserForm1.mp.Value = 1
    UserForm1.mp.Pages(0).Caption = "General"
    Debug.Print UserForm1.mp.Pages.Count, UserForm1.mp.Pages(0).Caption, UserForm1.mp.Value
    UserForm1.Show vbModeless
End Sub''', form_code="", form=f)
    rt.run("Go")
    assert host.prints[0].split() == ["2", "General", "1"]
    inst = rt.form_defaults["userform1"]
    v = host.views[0]
    n = len(v.calls)
    inst.user_set("t", "value", "typed")
    assert len(v.calls) == n
    assert inst.find_ctrl("t")["value"] == "typed"


def test_no_gui_host_gives_clear_error():
    book = Workbook()
    rt = VBARuntime(book, FormulaEngine(book), VBAHost())
    rt.load_modules([{"name": "Module1", "type": "standard", "code": "Sub Go()\n UserForm1.Show\nEnd Sub"},
                     {"name": "UserForm1", "type": "form", "code": "", "form": new_form()}])
    with pytest.raises(VBAError) as ei:
        rt.run("Go")
    assert "UserForms need the GUI" in ei.value.message
