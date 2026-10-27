"""Unit tests for the formula evaluator, independent of any workbook."""

from __future__ import annotations

import pytest

from officecheck.formula import DIV0, NA, Evaluator, ExcelError, UnsupportedFormula

GRID = {
    ("S", 1, 1): 1.0, ("S", 2, 1): 2.0, ("S", 3, 1): 3.0, ("S", 4, 1): "text", ("S", 5, 1): None,
    ("S", 1, 2): "a", ("S", 2, 2): "b", ("S", 3, 2): "c",
    ("S", 1, 3): 0.0,
    ("Other Sheet", 1, 1): 10.0,
}


def ev(formula: str):
    e = Evaluator(lambda s, r, c: GRID.get((s, r, c)), ["S", "Other Sheet"], lambda s: 5)
    return e.eval_formula(formula, "S")


@pytest.mark.parametrize("formula, expected", [
    ("=1+2*3", 7.0), ("=(1+2)*3", 9.0), ("=-2^2", 4.0), ("=2^3^2", 64.0), ("=50%", 0.5),
    ("=10/4", 2.5), ('="a"&1&TRUE', "a1TRUE"), ("=1=1", True), ('="A"="a"', True),
    ('=1<"a"', True), ("=TRUE>\"z\"", True),
    ("=SUM(A1:A5)", 6.0), ("=SUM(A1:A3,10,TRUE)", 17.0), ("=AVERAGE(A1:A4)", 2.0),
    ("=COUNT(A1:A5)", 3.0), ('=COUNT(1,"2","x",TRUE)', 3.0), ("=COUNTA(A1:A5)", 4.0),
    ("=MIN(A1:A3)", 1.0), ("=MAX(A1:A3)", 3.0), ("=SUM(A:A)", 6.0),
    ("=IF(A1>0,\"pos\",\"neg\")", "pos"), ("=IF(A1>5,1)", False), ("=IF(C1=0,0,1/C1)", 0.0),
    ("=IFERROR(1/C1,-1)", -1.0),
    ("=VLOOKUP(2,A1:B3,2,FALSE)", "b"), ("=VLOOKUP(2.5,A1:B3,2)", "b"), ("=VLOOKUP(2,A1:B3,2,0)", "b"),
    ('=VLOOKUP("c*",B1:B3,1,FALSE)', "c"),
    ("=ROUND(2.5,0)", 3.0), ("=ROUND(-2.5,0)", -3.0), ("=ROUND(1.005,2)", 1.01), ("=ABS(-3)", 3.0),
    ("=AND(TRUE,A1=1)", True), ("=OR(FALSE,0)", False), ("=NOT(0)", True),
    ("='Other Sheet'!A1*2", 20.0), ("=$A$1+A$2", 3.0), ('=CONCATENATE("x",A1,"y")', "x1y"),
    ("=A5+1", 1.0), ('=A5&"!"', "!"),
])
def test_values(formula, expected):
    assert ev(formula) == expected


@pytest.mark.parametrize("formula, err", [
    ("=1/0", DIV0), ("=1/C1", DIV0), ("=A4+1", ExcelError("#VALUE!")), ("=SUM(1/0,2)", DIV0),
    ("=AVERAGE(A4)", DIV0), ("=VLOOKUP(9,A1:B3,2,FALSE)", NA), ("=VLOOKUP(1,A1:B3,3,FALSE)", ExcelError("#REF!")),
    ("=#N/A+1", NA), ("=Missing!A1", ExcelError("#REF!")), ("=IF(1/0,1,2)", DIV0),
])
def test_errors(formula, err):
    assert ev(formula) == err


@pytest.mark.parametrize("formula", ["=XLOOKUP(1,A1:A2,B1:B2)", "=MyName+1", "=1+", "=SUM(1"])
def test_unsupported(formula):
    with pytest.raises(UnsupportedFormula):
        ev(formula)
