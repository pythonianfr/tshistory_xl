from pyxll import xl_macro, xl_app, xl_func
from tshistory_xl.excel_connect import (
    macro_pull_all,
    macro_push_all,
    macro_pull_tab,
    macro_push_tab,
)


@xl_macro
def pull_all_pyxll():
    wd = xl_app().Application.ActiveWorkbook.FullName
    macro_pull_all(wd)


@xl_macro
def pull_tab_pyxll():
    xl = xl_app()
    tab = xl.ActiveSheet.Name
    wd = xl.Application.ActiveWorkbook.FullName
    macro_pull_tab(wd, tab)


@xl_macro
def push_all_pyxll():
    wd = xl_app().Application.ActiveWorkbook.FullName
    macro_push_all(wd)


@xl_macro
def push_tab_pyxll():
    xl = xl_app()
    wd = xl.Application.ActiveWorkbook.FullName
    tab = xl.ActiveSheet.Name
    macro_push_tab(wd, tab)
