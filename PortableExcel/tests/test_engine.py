import math
import time

import pytest

from portable_excel import model
from portable_excel.model import Workbook
from portable_excel.engine import (
    FormulaEngine, FormulaError, RangeRef, FUNCTIONS, format_value, shift_formula,
    insert_delete, rename_sheet_refs, delete_sheet_refs,
)


@pytest.fixture
def wb():
    b = Workbook()
    return b


def setup(b, cells, sheet=None):
    sh = sheet or b.active
    for ref, raw in cells.items():
        r, c = model.parse_a1(ref)
        sh.set_raw(r, c, raw)
    return FormulaEngine(b)


def val(e, sh, ref):
    r, c = model.parse_a1(ref)
    return e.value(sh, r, c)


def ev(formula, cells=None):
    b = Workbook()
    e = setup(b, cells or {})
    return e.evaluate(formula, b.active)


# ---------------------------------------------------------------- basics
@pytest.mark.parametrize("f,expected", [
    ("=1+2*3", 7), ("=(1+2)*3", 9), ("=-2^2", 4), ("=2^3^2", 64), ("=10-2-3", 5),
    ("=8/2/2", 2), ("=50%", 0.5), ("=-50%*2", -1), ("=2*-3", -6), ("=1+2&3", "33"),
    ("=1<2", True), ("=\"a\"=\"A\"", True), ("=\"b\">\"A\"", True), ("=1<\"a\"", True),
    ("=\"z\"<TRUE", True), ("=TRUE+1", 2), ("=\"3\"+4", 7), ("=\"a\"\"b\"", 'a"b'),
    ("=1&\"\"", "1"), ("=1/4&\"\"", "0.25"), ("=0.1+0.2=0.3", True), ("=1e3", 1000),
    ("= 1 + 1 ", 2), ("=+5", 5), ("=--5", 5), ("=3-1=2", True),
])
def test_precedence_and_ops(f, expected):
    assert ev(f) == expected


@pytest.mark.parametrize("f,err", [
    ("=1/0", model.ERR_DIV0), ("=\"a\"+1", model.ERR_VALUE), ("=NOSUCH(1)", model.ERR_NAME),
    ("=1+", model.ERR_SYNTAX), ("=(1", model.ERR_SYNTAX), ("=\"abc", model.ERR_SYNTAX),
    ("=#N/A+1", model.ERR_NA), ("=#REF!", model.ERR_REF), ("=SQRT(-1)", model.ERR_NUM),
    ("=undefinedname", model.ERR_NAME), ("=1/0&\"x\"", model.ERR_DIV0), ("=0^0", model.ERR_NUM),
])
def test_errors(f, err):
    assert ev(f) == err


def test_empty_cells(wb):
    sh = wb.active
    e = setup(wb, {"B1": "=A1", "B2": "=A1&\"x\"", "B3": "=A1+5", "B4": "=A1=\"\"",
                   "B5": "=A1=0", "B6": "=ISBLANK(A1)"})
    assert val(e, sh, "B1") == 0
    assert e.display_text(sh, 0, 1) == "0"
    assert val(e, sh, "B2") == "x"
    assert val(e, sh, "B3") == 5
    assert val(e, sh, "B4") is True
    assert val(e, sh, "B5") is True
    assert val(e, sh, "B6") is True
    assert val(e, sh, "A1") is None


def test_literals_and_cache(wb):
    sh = wb.active
    e = setup(wb, {"A1": "5", "A2": "abc", "A3": "TRUE", "A4": "=A1*2"})
    assert val(e, sh, "A1") == 5.0
    assert val(e, sh, "A2") == "abc"
    assert val(e, sh, "A3") is True
    assert val(e, sh, "A4") == 10
    sh.set_raw(0, 0, "7")
    assert val(e, sh, "A4") == 10  # cached
    e.invalidate()
    assert val(e, sh, "A4") == 14


def test_error_propagation(wb):
    sh = wb.active
    e = setup(wb, {"A1": "=1/0", "A2": "=A1+1", "A3": "=SUM(A1:A2)", "A4": "=IFERROR(A2,\"ok\")"})
    assert val(e, sh, "A2") == model.ERR_DIV0
    assert val(e, sh, "A3") == model.ERR_DIV0
    assert val(e, sh, "A4") == "ok"


def test_cycles(wb):
    sh = wb.active
    e = setup(wb, {"A1": "=B1+1", "B1": "=A1+1", "C1": "=C1", "D1": "=A1*2", "E1": "=5"})
    assert val(e, sh, "A1") == model.ERR_CYCLE
    assert val(e, sh, "B1") == model.ERR_CYCLE
    assert val(e, sh, "C1") == model.ERR_CYCLE
    assert val(e, sh, "D1") == model.ERR_CYCLE
    assert val(e, sh, "E1") == 5
    e.recalc()
    assert val(e, sh, "A1") == model.ERR_CYCLE


def test_long_cycle_no_crash(wb):
    sh = wb.active
    n = 500
    for i in range(n):
        sh.set_raw(i, 0, f"=A{i + 2}+1" if i < n - 1 else "=A1+1")
    e = FormulaEngine(wb)
    assert val(e, sh, "A1") == model.ERR_CYCLE


def test_cross_sheet(wb):
    s1 = wb.active
    s2 = wb.add_sheet("Data")
    s3 = wb.add_sheet("My Sheet")
    s2.set_raw(0, 0, "10")
    s3.set_raw(1, 1, "5")
    e = setup(wb, {"A1": "=Data!A1*2", "A2": "='My Sheet'!B2+Data!A1",
                   "A3": "=SUM('My Sheet'!A1:B2)", "A4": "=Nope!A1", "A5": "=data!a1"})
    assert val(e, s1, "A1") == 20
    assert val(e, s1, "A2") == 15
    assert val(e, s1, "A3") == 5
    assert val(e, s1, "A4") == model.ERR_REF
    assert val(e, s1, "A5") == 10


def test_defined_names(wb):
    sh = wb.active
    wb.defined_names["Rate"] = "Sheet1!$B$1"
    wb.defined_names["Items"] = "Sheet1!$A$1:$A$3"
    wb.defined_names["Konst"] = "=0.5"
    e = setup(wb, {"A1": "1", "A2": "2", "A3": "3", "B1": "10",
                   "C1": "=rate*2", "C2": "=SUM(Items)", "C3": "=Konst*4"})
    assert val(e, sh, "C1") == 20
    assert val(e, sh, "C2") == 6
    assert val(e, sh, "C3") == 2


def test_absolute_and_whole_ranges(wb):
    sh = wb.active
    e = setup(wb, {"A1": "1", "A2": "2", "B1": "3", "C1": "=SUM($A$1:$B$2)",
                   "D1": "=SUM(A:A)", "D2": "=SUM(1:1)", "D3": "=SUM(A:B)"})
    assert val(e, sh, "C1") == 6
    assert val(e, sh, "D1") == 3
    assert val(e, sh, "D2") == 13  # A1 + B1 + C1 + D1
    assert val(e, sh, "D3") == 6


def test_array_ops_and_constants():
    assert ev("=SUM({1,2;3,4})") == 10
    assert ev("=SUMPRODUCT((A1:A4>1)*B1:B4)",
              {"A1": "1", "A2": "2", "A3": "3", "A4": "0",
               "B1": "10", "B2": "20", "B3": "30", "B4": "40"}) == 50
    assert ev("=SUM(A1:A3*2)", {"A1": "1", "A2": "2", "A3": "3"}) == 12


def test_evaluate_ranges(wb):
    sh = wb.active
    e = setup(wb, {"A1": "1", "B1": "x", "A2": "=A1+1"})
    assert e.evaluate("A1:B2", sh) == [[1.0, "x"], [2.0, None]]
    assert e.evaluate("=A2", sh) == 2
    assert e.evaluate("B2", sh) is None


# ---------------------------------------------------------------- aggregates
AGG = {"A1": "1", "A2": "2", "A3": "3", "A4": "4", "A5": "text", "A6": "TRUE",
       "B1": "2", "B2": "2", "B3": "5"}


@pytest.mark.parametrize("f,expected", [
    ("=SUM(A1:A6)", 10), ("=SUM(A1:A3,10,\"5\",TRUE)", 22), ("=AVERAGE(A1:A6)", 2.5),
    ("=MIN(A1:A6)", 1), ("=MAX(A1:A6)", 4), ("=MAX(C1:C3)", 0), ("=COUNT(A1:A7)", 4),
    ("=COUNTA(A1:A7)", 6), ("=COUNTBLANK(A1:A8)", 2), ("=PRODUCT(A1:A4)", 24),
    ("=MEDIAN(A1:A4)", 2.5), ("=MEDIAN(1,3,2)", 2), ("=MODE(B1:B3)", 2),
    ("=LARGE(A1:A4,1)", 4), ("=SMALL(A1:A4,2)", 2), ("=RANK(3,A1:A4)", 2),
    ("=RANK(3,A1:A4,1)", 3), ("=SUMPRODUCT(A1:A3,B1:B3)", 21),
    ("=ROUND(STDEV(A1:A4),6)", round(math.sqrt(5 / 3), 6)), ("=VAR(A1:A4)", 5 / 3),
    ("=VARP(A1:A4)", 1.25), ("=STDEVP(1,1)", 0),
])
def test_aggregates(f, expected):
    assert ev(f, AGG) == pytest.approx(expected)


def test_aggregate_errors():
    assert ev("=AVERAGE(C1:C3)") == model.ERR_DIV0
    assert ev("=LARGE(A1:A2,5)", AGG) == model.ERR_NUM
    assert ev("=SUM(\"abc\")") == model.ERR_VALUE
    assert ev("=MODE(1,2,3)") == model.ERR_NA


# ---------------------------------------------------------------- conditional
IFS = {"A1": "apple", "A2": "banana", "A3": "apricot", "A4": "cherry", "A5": "",
       "B1": "10", "B2": "20", "B3": "30", "B4": "40", "B5": "50",
       "C1": "1", "C2": "6", "C3": "7", "C4": "3", "C5": "9"}


@pytest.mark.parametrize("f,expected", [
    ("=SUMIF(C1:C5,\">5\")", 22), ("=SUMIF(C1:C5,\">5\",B1:B5)", 100),
    ("=SUMIF(A1:A5,\"ap*\",B1:B5)", 40), ("=SUMIF(A1:A5,\"?anana\",B1:B5)", 20),
    ("=SUMIF(A1:A5,\"<>apple\",B1:B5)", 140), ("=SUMIF(A1:A5,\"=APPLE\",B1:B5)", 10),
    ("=SUMIF(C1:C5,7,B1:B5)", 30), ("=COUNTIF(C1:C5,\">=6\")", 3),
    ("=COUNTIF(A1:A5,\"*an*\")", 1), ("=COUNTIF(A1:A6,\"\")", 2),
    ("=COUNTIF(A1:A5,\"<>\")", 4), ("=COUNTIF(C1:C5,\"<5\")", 2),
    ("=SUMIFS(B1:B5,A1:A5,\"a*\",C1:C5,\">5\")", 30),
    ("=COUNTIFS(A1:A5,\"a*\",C1:C5,\"<5\")", 1),
    ("=AVERAGEIF(C1:C5,\">5\",B1:B5)", 100 / 3), ("=AVERAGEIFS(B1:B5,C1:C5,\">5\")", 100 / 3),
    ("=MAXIFS(B1:B5,C1:C5,\"<7\")", 40), ("=MINIFS(B1:B5,C1:C5,\">5\")", 20),
    ("=COUNTIF(A:A,\"a*\")", 2), ("=SUMIF(A:A,\"cherry\",B:B)", 40),
])
def test_conditional(f, expected):
    assert ev(f, IFS) == pytest.approx(expected)


def test_countif_blank_whole_column():
    assert ev("=COUNTIF(A:A,\"\")", {"A1": "x", "A3": "y"}) == model.MAX_ROWS - 2


def test_averageif_div0():
    assert ev("=AVERAGEIF(C1:C5,\">100\")", IFS) == model.ERR_DIV0


# ---------------------------------------------------------------- logical
@pytest.mark.parametrize("f,expected", [
    ("=IF(1>0,\"y\",\"n\")", "y"), ("=IF(0,\"y\",\"n\")", "n"), ("=IF(FALSE,1)", False),
    ("=IF(TRUE,1/1,1/0)", 1), ("=IF(FALSE,1/0,2)", 2), ("=IFS(1>2,\"a\",2>1,\"b\")", "b"),
    ("=IFS(FALSE,1)", model.ERR_NA), ("=IFERROR(1/0,\"e\")", "e"), ("=IFERROR(5,\"e\")", 5),
    ("=IFNA(NA(),\"na\")", "na"), ("=IFNA(1/0,\"na\")", model.ERR_DIV0),
    ("=AND(TRUE,1,2>1)", True), ("=AND(TRUE,FALSE)", False), ("=OR(FALSE,0,1)", True),
    ("=NOT(TRUE)", False), ("=XOR(TRUE,TRUE,TRUE)", True), ("=XOR(TRUE,TRUE)", False),
    ("=SWITCH(2,1,\"a\",2,\"b\",\"z\")", "b"), ("=SWITCH(9,1,\"a\",\"z\")", "z"),
    ("=SWITCH(9,1,\"a\")", model.ERR_NA), ("=CHOOSE(2,\"a\",\"b\",\"c\")", "b"),
    ("=CHOOSE(5,\"a\")", model.ERR_VALUE), ("=AND(\"x\")", model.ERR_VALUE),
    ("=if(true,\"low\",\"x\")", "low"),
])
def test_logical(f, expected):
    assert ev(f) == expected


def test_if_lazy_cycle_branch(wb):
    sh = wb.active
    e = setup(wb, {"A1": "=IF(TRUE,5,A1)"})
    assert val(e, sh, "A1") == 5


# ---------------------------------------------------------------- math
@pytest.mark.parametrize("f,expected", [
    ("=ROUND(2.5,0)", 3), ("=ROUND(-2.5,0)", -3), ("=ROUND(2.675,2)", 2.68),
    ("=ROUND(1234.567,-2)", 1200), ("=ROUNDUP(1.21,1)", 1.3), ("=ROUNDUP(-1.21,1)", -1.3),
    ("=ROUNDDOWN(1.29,1)", 1.2), ("=INT(-1.5)", -2), ("=TRUNC(-1.59,1)", -1.5),
    ("=MOD(10,3)", 1), ("=MOD(-10,3)", 2), ("=MOD(10,-3)", -2), ("=ABS(-3)", 3),
    ("=SQRT(16)", 4), ("=POWER(2,10)", 1024), ("=EXP(0)", 1), ("=LN(EXP(2))", 2),
    ("=LOG(100)", 2), ("=LOG(8,2)", 3), ("=LOG10(1000)", 3), ("=PI()", math.pi),
    ("=CEILING(4.2,1)", 5), ("=CEILING(4.2,0.5)", 4.5), ("=CEILING(-2.5,-2)", -4),
    ("=CEILING(-2.5,2)", -2), ("=FLOOR(4.8,1)", 4), ("=FLOOR(-2.5,-2)", -2),
    ("=SIGN(-4)", -1), ("=SIGN(0)", 0), ("=FACT(5)", 120), ("=EVEN(3)", 4), ("=ODD(2)", 3),
    ("=QUOTIENT(7,2)", 3), ("=GCD(12,18)", 6), ("=LCM(4,6)", 12), ("=MROUND(10,3)", 9),
])
def test_math(f, expected):
    assert ev(f) == pytest.approx(expected)


def test_math_errors():
    assert ev("=MOD(1,0)") == model.ERR_DIV0
    assert ev("=LN(0)") == model.ERR_NUM
    assert ev("=CEILING(2,-1)") == model.ERR_NUM
    assert ev("=FLOOR(2,0)") == model.ERR_DIV0


def test_rand():
    for _ in range(20):
        x = ev("=RAND()")
        assert 0 <= x < 1
        y = ev("=RANDBETWEEN(1,3)")
        assert y in (1, 2, 3)


# ---------------------------------------------------------------- text
@pytest.mark.parametrize("f,expected", [
    ("=LEN(\"hello\")", 5), ("=LEFT(\"hello\",2)", "he"), ("=LEFT(\"hello\")", "h"),
    ("=RIGHT(\"hello\",3)", "llo"), ("=RIGHT(\"hi\",0)", ""), ("=MID(\"hello\",2,3)", "ell"),
    ("=UPPER(\"ab\")", "AB"), ("=LOWER(\"AB\")", "ab"), ("=PROPER(\"hello wORLD-x\")", "Hello World-X"),
    ("=TRIM(\"  a   b  \")", "a b"), ("=CONCAT(\"a\",1,TRUE)", "a1TRUE"),
    ("=CONCATENATE(\"a\",\"b\")", "ab"), ("=CONCAT(A1:A3)", "xz"),
    ("=TEXTJOIN(\",\",TRUE,A1:A3)", "x,z"), ("=TEXTJOIN(\",\",FALSE,A1:A3)", "x,,z"),
    ("=SUBSTITUTE(\"aaa\",\"a\",\"b\")", "bbb"), ("=SUBSTITUTE(\"aaa\",\"a\",\"b\",2)", "aba"),
    ("=REPLACE(\"abcdef\",2,3,\"X\")", "aXef"), ("=FIND(\"l\",\"hello\")", 3),
    ("=FIND(\"l\",\"hello\",4)", 4), ("=FIND(\"L\",\"hello\")", model.ERR_VALUE),
    ("=SEARCH(\"L\",\"hello\")", 3), ("=SEARCH(\"e?l\",\"hello\")", 2),
    ("=REPT(\"ab\",3)", "ababab"), ("=EXACT(\"a\",\"A\")", False), ("=EXACT(\"a\",\"a\")", True),
    ("=VALUE(\"12.5\")", 12.5), ("=VALUE(\"abc\")", model.ERR_VALUE), ("=CHAR(65)", "A"),
    ("=CODE(\"A\")", 65), ("=TEXT(1234.567,\"#,##0.00\")", "1,234.57"),
    ("=TEXT(0.256,\"0%\")", "26%"), ("=TEXT(45292,\"yyyy-mm-dd\")", "2024-01-01"),
    ("=TEXT(5,\"0.00\")", "5.00"), ("=TEXT(0.5,\"hh:mm\")", "12:00"),
    ("=LEN(123)", 3), ("=\"x\"&TEXT(3,\"000\")", "x003"),
])
def test_text(f, expected):
    assert ev(f, {"A1": "x", "A3": "z"}) == expected


# ---------------------------------------------------------------- lookup
LK = {"A1": "1", "B1": "one", "C1": "uno",
      "A2": "2", "B2": "two", "C2": "dos",
      "A3": "5", "B3": "five", "C3": "cinco",
      "E1": "Name", "F1": "Ann", "G1": "Bob",
      "E2": "Age", "F2": "30", "G2": "40"}


@pytest.mark.parametrize("f,expected", [
    ("=VLOOKUP(2,A1:C3,2,FALSE)", "two"), ("=VLOOKUP(3,A1:C3,3)", "dos"),
    ("=VLOOKUP(3,A1:C3,3,TRUE)", "dos"), ("=VLOOKUP(9,A1:C3,2)", "five"),
    ("=VLOOKUP(0,A1:C3,2)", model.ERR_NA), ("=VLOOKUP(3,A1:C3,2,FALSE)", model.ERR_NA),
    ("=VLOOKUP(2,A1:C3,4,FALSE)", model.ERR_REF), ("=VLOOKUP(\"t*\",B1:C3,2,FALSE)", "dos"),
    ("=VLOOKUP(1,A:C,2,FALSE)", "one"),
    ("=HLOOKUP(\"Bob\",F1:G2,2,FALSE)", 40), ("=HLOOKUP(\"age\",E1:G2,1,FALSE)", model.ERR_NA),
    ("=MATCH(5,A1:A3,0)", 3), ("=MATCH(4,A1:A3)", 2), ("=MATCH(4,A1:A3,1)", 2),
    ("=MATCH(\"FIVE\",B1:B3,0)", 3), ("=MATCH(7,A1:A3,0)", model.ERR_NA),
    ("=MATCH(\"Bob\",E1:G1,0)", 3), ("=MATCH(3,{5,4,3,1},-1)", 3),
    ("=INDEX(A1:C3,2,3)", "dos"), ("=INDEX(A1:A3,3)", 5), ("=INDEX(E1:G1,2)", "Ann"),
    ("=INDEX(A1:C3,4,1)", model.ERR_REF), ("=SUM(INDEX(A1:C3,0,1))", 8),
    ("=INDEX(B1:B3,MATCH(5,A1:A3,0))", "five"),
    ("=XLOOKUP(5,A1:A3,B1:B3)", "five"), ("=XLOOKUP(7,A1:A3,B1:B3,\"none\")", "none"),
    ("=XLOOKUP(7,A1:A3,B1:B3)", model.ERR_NA), ("=XLOOKUP(4,A1:A3,B1:B3,,-1)", "two"),
    ("=XLOOKUP(4,A1:A3,B1:B3,,1)", "five"), ("=XLOOKUP(\"Bob\",F1:G1,F2:G2)", 40),
    ("=ROW(B7)", 7), ("=COLUMN(C2)", 3), ("=ROWS(A1:C3)", 3), ("=COLUMNS(A1:C3)", 3),
    ("=ROWS({1;2})", 2), ("=SUM(OFFSET(A1,1,0,2,1))", 7), ("=INDIRECT(\"B\"&2)", "two"),
])
def test_lookup(f, expected):
    assert ev(f, LK) == expected


def test_row_column_no_args(wb):
    sh = wb.active
    e = setup(wb, {"C5": "=ROW()*100+COLUMN()"})
    assert val(e, sh, "C5") == 503


# ---------------------------------------------------------------- info
@pytest.mark.parametrize("f,expected", [
    ("=ISBLANK(Z9)", True), ("=ISBLANK(A1)", False), ("=ISNUMBER(A1)", True),
    ("=ISNUMBER(B1)", False), ("=ISTEXT(B1)", True), ("=ISTEXT(A1)", False),
    ("=ISERROR(1/0)", True), ("=ISERROR(1)", False), ("=ISNA(NA())", True),
    ("=ISNA(1/0)", False), ("=ISERR(NA())", False), ("=ISLOGICAL(TRUE)", True),
    ("=ISLOGICAL(1)", False), ("=ISERROR(C1)", True),
])
def test_info(f, expected):
    assert ev(f, {"A1": "1", "B1": "x", "C1": "=1/0"}) == expected


# ---------------------------------------------------------------- dates
@pytest.mark.parametrize("f,expected", [
    ("=DATE(2024,1,1)", 45292), ("=DATE(1900,1,1)", 1), ("=DATE(1900,3,1)", 61),
    ("=DATE(2024,13,1)", 45658), ("=DATE(2024,3,0)", 45351),
    ("=YEAR(45292)", 2024), ("=MONTH(45351)", 2), ("=DAY(45351)", 29),
    ("=DAY(60)", 29), ("=MONTH(60)", 2), ("=HOUR(0.75)", 18), ("=MINUTE(0.5+1/1440*5)", 5),
    ("=WEEKDAY(45292)", 2), ("=WEEKDAY(45292,2)", 1), ("=WEEKDAY(45292,3)", 0),
    ("=EDATE(DATE(2024,1,31),1)", 45351), ("=EOMONTH(DATE(2024,1,15),0)", 45322),
    ("=DATEDIF(DATE(2020,5,15),DATE(2024,5,14),\"Y\")", 3),
    ("=DATEDIF(DATE(2020,5,15),DATE(2024,5,15),\"Y\")", 4),
    ("=DATEDIF(DATE(2024,1,15),DATE(2024,3,10),\"M\")", 1),
    ("=DATEDIF(DATE(2024,1,1),DATE(2024,3,1),\"D\")", 60),
    ("=DATEDIF(DATE(2024,1,15),DATE(2024,3,10),\"MD\")", 24),
    ("=DATEDIF(DATE(2020,5,15),DATE(2024,8,1),\"YM\")", 2),
    ("=DATEDIF(DATE(2024,2,1),DATE(2024,1,1),\"D\")", model.ERR_NUM),
    ("=TIME(12,30,0)", 0.520833333333333), ("=DATE(2024,1,1)+1", 45293),
    ("=YEAR(\"2024-06-30\")", 2024),
])
def test_dates(f, expected):
    r = ev(f)
    if isinstance(expected, model.XLError):
        assert r == expected
    else:
        assert r == pytest.approx(expected)


def test_today_now():
    import datetime as dt
    t = ev("=TODAY()")
    assert t == (dt.date.today() - dt.date(1899, 12, 30)).days
    n = ev("=NOW()")
    assert t <= n < t + 1


# ---------------------------------------------------------------- formatting
@pytest.mark.parametrize("v,fmt,expected", [
    (1234.5, None, "1234.5"), (1234.5, "General", "1234.5"), (1234.5, "0", "1235"),
    (1234.5, "0.00", "1234.50"), (1234.5, "#,##0", "1,235"), (1234567.891, "#,##0.00", "1,234,567.89"),
    (-1234.5, "#,##0.00", "-1,234.50"), (0.256, "0%", "26%"), (0.25678, "0.00%", "25.68%"),
    (1234.5, "$#,##0.00", "$1,234.50"), (-5, "$#,##0.00", "-$5.00"),
    (1234567, "Rs #,##0", "Rs 1,234,567"), (1234.5, "\"Rs\" #,##0.00", "Rs 1,234.50"),
    (45292, "dd/mm/yyyy", "01/01/2024"), (45292, "yyyy-mm-dd", "2024-01-01"),
    (45292, "mm/dd/yyyy", "01/01/2024"), (45351, "d-mmm-yy", "29-Feb-24"),
    (45351, "mmmm d, yyyy", "February 29, 2024"), (45292, "dddd", "Monday"),
    (45292.75, "hh:mm", "18:00"), (45292.75, "hh:mm:ss", "18:00:00"),
    (45292.75, "h:mm AM/PM", "6:00 PM"), (45292.5, "dd/mm/yyyy hh:mm", "01/01/2024 12:00"),
    (1.5, "[h]:mm", "36:00"), ("abc", "0.00", "abc"), ("abc", "@", "abc"),
    (True, "0.00", "TRUE"), (model.ERR_NA, "0.00", "#N/A"), (None, "0.00", ""),
    (5, "0.00;(0.00)", "5.00"), (-5, "0.00;(0.00)", "(5.00)"), (0, "0;-0;\"zero\"", "zero"),
    (12345678, "0.00E+00", "1.23E+07"), (0.5, "#.##", ".5"), (5, "#.##", "5."),
    (1234.5, "", "1234.5"), (3, "000", "003"), (1500000, "#,##0,", "1,500"),
])
def test_format_value(v, fmt, expected):
    assert format_value(v, fmt) == expected


def test_display_text(wb):
    sh = wb.active
    e = setup(wb, {"A1": "0.5", "A2": "=DATE(2024,1,1)"})
    sh.update_style(0, 0, number_format="0.0%")
    sh.update_style(1, 0, number_format="dd/mm/yyyy")
    assert e.display_text(sh, 0, 0) == "50.0%"
    assert e.display_text(sh, 1, 0) == "01/01/2024"
    assert e.display_text(sh, 5, 5) == ""


# ---------------------------------------------------------------- call_function
def test_call_function(wb):
    sh = wb.active
    e = setup(wb, {"A1": "1", "A2": "2", "A3": "3", "B1": "a", "B2": "b", "B3": "c"})
    rr = RangeRef(sh, 0, 0, 2, 0)
    assert rr.nrows == 3 and rr.ncols == 1
    assert e.call_function("Sum", [rr, 4]) == 10
    assert e.call_function("MAX", [[1, 7, 3]]) == 7
    assert e.call_function("sum", [[[1, 2], [3, 4]]]) == 10
    assert e.call_function("VLOOKUP", [2, RangeRef(sh, 0, 0, 2, 1), 2, False]) == "b"
    assert e.call_function("Match", ["c", RangeRef(sh, 0, 1, 2, 1), 0]) == 3
    assert e.call_function("IF", [True, "y", "n"]) == "y"
    assert e.call_function("Round", [2.345, 2]) == 2.35
    with pytest.raises(FormulaError) as ei:
        e.call_function("VLOOKUP", [9, RangeRef(sh, 0, 0, 2, 1), 2, False])
    assert ei.value.error == model.ERR_NA
    with pytest.raises(FormulaError) as ei:
        e.call_function("NoSuchFn", [1])
    assert ei.value.error == model.ERR_NAME


def test_functions_registry():
    for name in ["SUM", "VLOOKUP", "XLOOKUP", "TEXTJOIN", "DATEDIF", "SUMIFS", "IFERROR"]:
        assert name in FUNCTIONS
    assert FormulaEngine.FUNCTIONS is FUNCTIONS


# ---------------------------------------------------------------- shift_formula
@pytest.mark.parametrize("raw,dr,dc,expected", [
    ("=A1+B2", 1, 1, "=B2+C3"), ("=$A$1+A1", 2, 2, "=$A$1+C3"), ("=$A1+A$1", 1, 1, "=$A2+B$1"),
    ("=SUM(A1:B2)", 3, 0, "=SUM(A4:B5)"), ("=SUM(A:A)", 5, 1, "=SUM(B:B)"),
    ("=SUM(1:2)", 1, 3, "=SUM(2:3)"), ("=Sheet2!A1*2", 1, 0, "=Sheet2!A2*2"),
    ("='My Sheet'!A1", 0, 1, "='My Sheet'!B1"), ("=A1-1", -1, 0, "=#REF!-1"),
    ("=A1", 0, -1, "=#REF!"), ("=\"A1\"&A1", 1, 0, "=\"A1\"&A2"), ("hello A1", 1, 1, "hello A1"),
    ("=LOG10(A1)", 1, 0, "=LOG10(A2)"), ("=SUM( A1 , B1 )", 1, 0, "=SUM( A2 , B2 )"),
    ("=$A$1:$B$2", 9, 9, "=$A$1:$B$2"), ("=5", 3, 3, "=5"),
])
def test_shift_formula(raw, dr, dc, expected):
    assert shift_formula(raw, dr, dc) == expected


# ---------------------------------------------------------------- insert/delete
def test_insert_rows(wb):
    sh = wb.active
    other = wb.add_sheet("Other")
    setup(wb, {"A1": "1", "A2": "2", "A3": "3", "A4": "=SUM(A1:A3)", "B1": "=A3*2",
               "C1": "=$A$3", "D1": "=SUM(A:A)"})
    sh.update_style(2, 0, bold=True)
    sh.row_heights[2] = 40
    other.set_raw(0, 0, "=Sheet1!A3+Sheet1!A1")
    insert_delete(wb, sh, "row", 1, 2)  # insert 2 rows before row 2
    assert sh.get_raw(0, 0) == "1"
    assert sh.get_raw(1, 0) == ""
    assert sh.get_raw(3, 0) == "2"
    assert sh.get_raw(4, 0) == "3"
    assert sh.get_raw(5, 0) == "=SUM(A1:A5)"  # expanded
    assert sh.get_raw(0, 1) == "=A5*2"
    assert sh.get_raw(0, 2) == "=$A$5"
    assert sh.get_raw(0, 3) == "=SUM(A:A)"
    assert sh.get_style(4, 0).get("bold")
    assert sh.row_heights == {4: 40}
    assert other.get_raw(0, 0) == "=Sheet1!A5+Sheet1!A1"
    e = FormulaEngine(wb)
    assert e.value(sh, 5, 0) == 6


def test_delete_rows(wb):
    sh = wb.active
    setup(wb, {"A1": "1", "A2": "2", "A3": "3", "A4": "4", "A6": "=SUM(A1:A4)",
               "B1": "=A2", "B2": "=A4", "B3": "=SUM(A2:A3)"})
    insert_delete(wb, sh, "row", 1, -2)  # delete rows 2..3
    assert sh.get_raw(1, 0) == "4"
    assert sh.get_raw(3, 0) == "=SUM(A1:A2)"  # shrunk
    assert sh.get_raw(0, 1) == "=#REF!"
    # B2 and B3 were deleted; B4 not existing
    assert (1, 1) not in sh.cells
    e = FormulaEngine(wb)
    assert e.value(sh, 3, 0) == 5
    assert e.value(sh, 0, 1) == model.ERR_REF


def test_delete_rows_range_fully_and_partially(wb):
    sh = wb.active
    setup(wb, {"A10": "=SUM(A2:A3)", "B10": "=SUM(A3:A6)", "C10": "=SUM(A1:A3)"})
    insert_delete(wb, sh, "row", 1, -3)  # delete rows 2..4
    assert sh.get_raw(6, 0) == "=SUM(#REF!)"
    assert sh.get_raw(6, 1) == "=SUM(A2:A3)"
    assert sh.get_raw(6, 2) == "=SUM(A1:A1)"


def test_insert_delete_cols(wb):
    sh = wb.active
    setup(wb, {"A1": "1", "B1": "2", "C1": "3", "D1": "=SUM(A1:C1)", "E1": "=B1",
               "A2": "=SUM(1:1)"})
    sh.col_widths[2] = 120
    insert_delete(wb, sh, "col", 1, 1)
    assert sh.get_raw(0, 2) == "2"
    assert sh.get_raw(0, 4) == "=SUM(A1:D1)"
    assert sh.get_raw(0, 5) == "=C1"
    assert sh.get_raw(1, 0) == "=SUM(1:1)"
    assert sh.col_widths == {3: 120}
    insert_delete(wb, sh, "col", 2, -1)  # delete column C (the moved B)
    assert sh.get_raw(0, 3) == "=SUM(A1:C1)"
    assert sh.get_raw(0, 4) == "=#REF!"


def test_insert_delete_other_sheet_untouched(wb):
    sh = wb.active
    other = wb.add_sheet("Other")
    other.set_raw(0, 0, "=A5")  # refers to Other!A5, not Sheet1
    other.set_raw(0, 1, "=Sheet1!A5")
    insert_delete(wb, sh, "row", 0, 1)
    assert other.get_raw(0, 0) == "=A5"
    assert other.get_raw(0, 1) == "=Sheet1!A6"


def test_insert_updates_defined_names(wb):
    sh = wb.active
    wb.defined_names["Data"] = "Sheet1!$A$1:$A$5"
    insert_delete(wb, sh, "row", 2, 3)
    assert wb.defined_names["Data"] == "Sheet1!$A$1:$A$8"


# ---------------------------------------------------------------- sheet rename/delete
def test_rename_sheet_refs(wb):
    s1 = wb.active
    s2 = wb.add_sheet("Data")
    s2.set_raw(0, 0, "4")
    s1.set_raw(0, 0, "=Data!A1+data!A1+SUM(Data!A1:B2)+A1")
    wb.defined_names["X"] = "Data!$A$1"
    rename_sheet_refs(wb, "Data", "My Data")
    s2.name = "My Data"
    assert s1.get_raw(0, 0) == "='My Data'!A1+'My Data'!A1+SUM('My Data'!A1:B2)+A1"
    assert wb.defined_names["X"] == "'My Data'!$A$1"
    rename_sheet_refs(wb, "My Data", "Plain")
    assert s1.get_raw(0, 0) == "=Plain!A1+Plain!A1+SUM(Plain!A1:B2)+A1"
    rename_sheet_refs(wb, "Plain", "O'Neil")
    assert s1.get_raw(0, 0).startswith("='O''Neil'!A1")
    s2.name = "O'Neil"
    s1.set_raw(0, 0, s1.get_raw(0, 0).replace("+A1", "+0"))
    e = FormulaEngine(wb)
    assert e.value(s1, 0, 0) == 12


def test_delete_sheet_refs(wb):
    s1 = wb.active
    wb.add_sheet("Gone")
    s1.set_raw(0, 0, "=Gone!A1+1")
    s1.set_raw(0, 1, "=SUM(Gone!A1:B5)")
    delete_sheet_refs(wb, "Gone")
    wb.sheets.pop()
    assert s1.get_raw(0, 0) == "=#REF!+1"
    assert s1.get_raw(0, 1) == "=SUM(#REF!)"
    e = FormulaEngine(wb)
    assert e.value(s1, 0, 0) == model.ERR_REF


# ---------------------------------------------------------------- scale
def test_running_total_chain_recalc(wb):
    sh = wb.active
    n = 5000
    for i in range(n):
        sh.set_raw(i, 0, "1")
        sh.set_raw(i, 1, "=A1" if i == 0 else f"=B{i}+A{i + 1}")
    e = FormulaEngine(wb)
    e.recalc()
    assert e.value(sh, n - 1, 1) == n


def test_running_total_chain_without_recalc(wb):
    sh = wb.active
    n = 5000
    for i in range(n):
        sh.set_raw(i, 0, "1")
        sh.set_raw(i, 1, "=A1" if i == 0 else f"=B{i}+A{i + 1}")
    e = FormulaEngine(wb)
    assert e.value(sh, n - 1, 1) == n  # bottom-up request: must not overflow the stack


def test_whole_column_sum_speed(wb):
    sh = wb.active
    n = 20000
    for i in range(n):
        sh.set_raw(i, 0, str(i))
    for i in range(200):
        sh.set_raw(i, 2, "=SUM(A:A)")
    e = FormulaEngine(wb)
    t = time.perf_counter()
    e.recalc()
    assert e.value(sh, 199, 2) == n * (n - 1) / 2
    assert time.perf_counter() - t < 5.0
