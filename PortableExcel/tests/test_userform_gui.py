"""UserForm rendering + designer tests (offscreen Qt)."""
import os
import sys

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
os.environ["PORTABLEEXCEL_NO_LOCK"] = "1"
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

import pytest  # noqa: E402
from PySide6.QtCore import QPoint, Qt, QTimer  # noqa: E402
from PySide6.QtTest import QTest  # noqa: E402
from PySide6.QtWidgets import QApplication, QMessageBox  # noqa: E402

from portable_excel import forms  # noqa: E402
from portable_excel.gui import config  # noqa: E402
from portable_excel.gui.app import create_app  # noqa: E402
from portable_excel.gui.userform_designer import FormDesigner, rename_event_procs  # noqa: E402
from portable_excel.gui.userform_view import QtFormView, norm_prop  # noqa: E402


@pytest.fixture(scope="module")
def app():
    return create_app(["test"])


@pytest.fixture(autouse=True)
def temp_ini(tmp_path):
    config.set_ini_path(str(tmp_path / "PortableExcel.ini"))
    yield
    config.set_ini_path(None)


class RecordingInstance:
    """Fake runtime instance recording the GUI -> runtime calls."""

    def __init__(self, form, handlers=None):
        self.name = "UserForm1"
        self.form = forms.normalize_form(form)
        self.calls = []
        self.handlers = handlers or {}
        self.view = None

    def fire(self, ctrl, event, args=()):
        self.calls.append(("fire", ctrl, event, tuple(args)))
        h = self.handlers.get((ctrl, event))
        if h:
            return h(list(args))
        return list(args)

    def user_set(self, ctrl, prop, value):
        self.calls.append(("set", ctrl, prop, value))
        for c in self.form["controls"]:
            if c["name"] == ctrl:
                c[prop] = value

    def request_close(self, close_mode=0):
        self.calls.append(("close", close_mode))
        self.view.close()


def sample_form():
    N = forms.new_control
    f = forms.new_form("Test")
    f["controls"] = [
        N("TextBox", "TextBox1", left=6, top=6, width=100, tab_index=0),
        N("CheckBox", "CheckBox1", left=6, top=30, tab_index=1),
        N("Frame", "Frame1", left=6, top=54, width=150, height=60),
        N("OptionButton", "Opt1", parent="Frame1", left=6, top=0),
        N("OptionButton", "Opt2", parent="Frame1", left=6, top=20),
        N("ComboBox", "Combo1", left=120, top=6, list=["a", "b"], style=2),
        N("ListBox", "List1", left=120, top=30, height=40, list=[["x", "1"], ["y", "2"]],
          column_count=2),
        N("CommandButton", "OK", left=150, top=120, caption="OK", default=True),
        N("CommandButton", "Cancel", left=80, top=120, caption="Cancel", cancel=True),
    ]
    return f


def make_view(handlers=None):
    inst = RecordingInstance(sample_form(), handlers)
    v = QtFormView(inst)
    inst.view = v
    return inst, v


def test_norm_prop():
    assert norm_prop("BackColor") == "back_color"
    assert norm_prop("ListIndex") == "list_index"
    assert norm_prop("caption") == "caption"


def test_render_and_interaction_order(app):
    inst, v = make_view()
    v.show(False)
    QApplication.processEvents()
    r = v.renderer
    assert set(r.widgets) >= {"TextBox1", "CheckBox1", "Frame1", "Opt1", "List1"}
    # option buttons live inside the frame
    assert r.widgets["Opt1"].parent() is r.containers["Frame1"]
    tb = r.widgets["TextBox1"]
    tb.setFocus()
    QApplication.processEvents()
    inst.calls.clear()
    QTest.keyClicks(tb, "x")
    assert inst.calls == [("fire", "TextBox1", "KeyPress", (ord("x"),)),
                          ("set", "TextBox1", "value", "x"), ("set", "TextBox1", "text", "x"),
                          ("fire", "TextBox1", "Change", ())]
    inst.calls.clear()
    r.widgets["CheckBox1"].click()
    assert inst.calls[0] == ("set", "CheckBox1", "value", True)
    assert [c[2] for c in inst.calls[1:]] == ["Change", "Click", "AfterUpdate"]
    inst.calls.clear()
    r.widgets["Opt2"].click()
    r.widgets["Opt1"].click()
    assert ("set", "Opt2", "value", False) in inst.calls
    assert r.widgets["Opt1"].isChecked() and not r.widgets["Opt2"].isChecked()
    inst.calls.clear()
    r.widgets["List1"].setCurrentItem(r.widgets["List1"].topLevelItem(1))
    assert ("set", "List1", "list_index", 1) in inst.calls
    events = [c[2] for c in inst.calls if c[0] == "fire"]
    assert events[:2] == ["Change", "Click"]
    inst.calls.clear()
    r.widgets["OK"].click()
    assert inst.calls == [("fire", "OK", "Click", ())]
    v.close()


def test_keypress_can_be_swallowed_or_replaced(app):
    def upper(args):
        k = args[0]
        if chr(k) == "!":
            return [0]
        return [ord(chr(k).upper())]
    inst, v = make_view({("TextBox1", "KeyPress"): upper})
    v.show(False)
    tb = v.renderer.widgets["TextBox1"]
    tb.setFocus()
    QTest.keyClicks(tb, "ab!")
    assert tb.text() == "AB"
    v.close()


def test_exit_cancel_keeps_focus(app):
    inst, v = make_view({("TextBox1", "Exit"): lambda a: [True]})
    v.show(False)
    v.activateWindow()
    tb = v.renderer.widgets["TextBox1"]
    tb.setFocus()
    QApplication.processEvents()
    v.renderer.widgets["CheckBox1"].setFocus(Qt.TabFocusReason)
    QApplication.processEvents()
    QApplication.processEvents()
    assert ("fire", "TextBox1", "Exit", (False,)) in inst.calls
    assert QApplication.focusWidget() is tb
    v.close()


def test_set_prop_and_modal_show_returns(app):
    inst, v = make_view()
    v.set_prop(None, "Caption", "Changed")
    assert v.windowTitle() == "Changed"
    v.set_prop("TextBox1", "Value", "hello")
    assert v.renderer.widgets["TextBox1"].text() == "hello"
    assert not any(c[0] == "set" for c in inst.calls)  # no echo back to the runtime
    v.set_prop("CheckBox1", "Visible", False)
    v.add_control(forms.new_control("Label", "Late", caption="late"))
    assert "Late" in v.renderer.widgets
    QTimer.singleShot(50, lambda: v.hide())
    v.show(True)   # must return once hide() is called
    assert not v.isVisible()
    # X button -> request_close -> runtime closes the view
    from PySide6.QtGui import QCloseEvent
    QTimer.singleShot(50, lambda: QApplication.sendEvent(v, QCloseEvent()))
    v.show(True)
    assert ("close", 0) in inst.calls


def test_escape_clicks_cancel_button(app):
    inst, v = make_view()
    v.show(False)
    QTest.keyClick(v, Qt.Key_Escape)
    assert ("fire", "Cancel", "Click", ()) in inst.calls
    v.close()


# ------------------------------------------------------------------ designer
def test_designer_place_rename_code(app, monkeypatch):
    monkeypatch.setattr(QMessageBox, "warning", staticmethod(lambda *a, **k: QMessageBox.Ok))
    m = {"name": "UserForm1", "type": "form", "code": "", "form": forms.new_form("UserForm1")}
    d = FormDesigner(None, m)
    d.resize(900, 600)
    d.show()
    tb = d.place_control("TextBox", 13, 13)
    assert tb["name"] == "TextBox1" and tb["left"] == 12.0  # snapped to the 6pt grid
    fr = d.place_control("Frame", 6, 40, 150, 80)
    # place a control by mouse inside the frame -> parent is the frame
    d.set_tool("OptionButton")
    r = d.ctrl_rect(fr["name"])
    pos = r.center()
    QTest.mouseClick(d.surface.client, Qt.LeftButton, Qt.NoModifier, pos)
    opt = d.ctrl("OptionButton1")
    assert opt is not None and opt["parent"] == "Frame1"
    # selection + nudge + undo
    d.select(["TextBox1"])
    d.nudge(6, 0)
    assert d.ctrl("TextBox1")["left"] == 18.0
    d.undo()
    assert d.ctrl("TextBox1")["left"] == 12.0
    # view code creates the default event procedure
    assert d.view_code("TextBox1") == "TextBox1_Change"
    assert "Private Sub TextBox1_Change()" in m["code"]
    # rename renames the event procs and keeps parents in sync
    d.select(["Frame1"])
    assert d.rename("fraType")
    assert d.ctrl("OptionButton1")["parent"] == "fraType"
    d.select(["TextBox1"])
    assert d.rename("txtName")
    assert "Private Sub txtName_Change()" in m["code"]
    assert not d.rename("bad name")
    # copy/paste and delete
    d.copy_selected()
    names = d.paste()
    assert names == ["TextBox1"]
    d.select(["fraType"])
    d.delete_selected()
    assert d.ctrl("OptionButton1") is None  # children deleted with the frame
    # properties
    d.select(["txtName"])
    d.set_property("multi_line", True)
    assert d.ctrl("txtName")["multi_line"] is True
    d.close()


def test_rename_event_procs():
    code = "Private Sub Cmd_Click()\nCall Cmd_Click\nEnd Sub\nSub Cmd_Helper()\nEnd Sub\n"
    out = rename_event_procs(code, "Cmd", "btnOK")
    assert "btnOK_Click()" in out and "Call btnOK_Click" in out and "Cmd_Helper" in out


def test_designer_drag_move_and_resize(app):
    m = {"name": "UserForm1", "type": "form", "code": "", "form": forms.new_form()}
    d = FormDesigner(None, m)
    d.resize(900, 600)
    d.show()
    c = d.place_control("CommandButton", 12, 12)
    r = d.ctrl_rect(c["name"])
    client = d.surface.client
    start = r.center()
    QTest.mousePress(client, Qt.LeftButton, Qt.NoModifier, start)
    from PySide6.QtCore import QEvent, QPointF
    from PySide6.QtGui import QMouseEvent
    end = start + QPoint(forms.points_to_px(12), forms.points_to_px(6))
    ev = QMouseEvent(QEvent.MouseMove, QPointF(end), QPointF(client.mapToGlobal(end)),
                     Qt.NoButton, Qt.LeftButton, Qt.NoModifier)
    QApplication.sendEvent(client, ev)
    QTest.mouseRelease(client, Qt.LeftButton, Qt.NoModifier, end)
    assert (c["left"], c["top"]) == (24.0, 18.0)
    # resize with the bottom-right handle
    r = d.ctrl_rect(c["name"])
    h = d.handle_rects(r)["br"].center()
    QTest.mousePress(client, Qt.LeftButton, Qt.NoModifier, h)
    end = h + QPoint(forms.points_to_px(12), 0)
    ev = QMouseEvent(QEvent.MouseMove, QPointF(end), QPointF(client.mapToGlobal(end)),
                     Qt.NoButton, Qt.LeftButton, Qt.NoModifier)
    QApplication.sendEvent(client, ev)
    QTest.mouseRelease(client, Qt.LeftButton, Qt.NoModifier, end)
    assert c["width"] == 84.0
    d.close()


def test_editor_insert_userform_and_run(app, monkeypatch):
    from portable_excel.gui import main_window as mw
    monkeypatch.setattr(QMessageBox, "question", staticmethod(lambda *a, **k: QMessageBox.Yes))
    win = mw.MainWindow()
    win.show()
    ed = win.show_vba_editor()
    m = ed.insert_userform()
    assert m["type"] == "form" and m["name"] == "UserForm1"
    d = ed.designers["userform1"]
    d.place_control("TextBox", 6, 6)
    d.place_control("CommandButton", 6, 30, caption="OK")
    d.view_code("CommandButton1")
    assert ed.tabs.currentWidget() is ed.editors["userform1"]
    m["code"] = m["code"].replace(
        "Private Sub CommandButton1_Click()\n    \n",
        "Private Sub CommandButton1_Click()\n    Range(\"A1\").Value = TextBox1.Text\n"
        "    Unload Me\n")
    ed.code_replaced("UserForm1")
    assert win.book.vba_modified

    def interact():
        v = win.form_views[-1]
        t = v.renderer.widgets["TextBox1"]
        t.setFocus()
        QTest.keyClicks(t, "Bob")
        v.renderer.widgets["CommandButton1"].click()
        # a preview (no runtime form support) is closed explicitly
        if v.isVisible():
            v.close()
    QTimer.singleShot(200, interact)
    ed.open_designer("UserForm1")
    ed.run_current()  # F5 on a form shows it
    QApplication.processEvents()
    if type(win.form_views[-1].instance).__name__ != "PreviewInstance":
        assert win.book.active.get_raw(0, 0) == "Bob"
    win.modified = False
    win.close()
    ed.close()
