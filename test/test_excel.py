from functools import partial
from datetime import datetime, timedelta
from pathlib import Path
import logging

import pytest

import numpy as np
import pandas as pd

try:
    import xlwings as xw
except ImportError:
    xw = None

from tshistory_xl import excel_connect

from tshistory_xl.excel_connect import (
    compress_series,
    find_zone_by_name,
    find_zones_having_keyword,
    int_to_str,
    PULL,
    pull_rolling,
    push_parse_ts as parse_ts,
    PUSH,
    push_rolling,
    xlbook
)


def genserie(start, freq, repeat, initval=None, tz=None, name=None):
    if initval is None:
        values = range(repeat)
    else:
        values = initval * repeat
    return pd.Series(values,
                     name=name,
                     index=pd.date_range(start=start,
                                         freq=freq,
                                         periods=repeat,
                                         tz=tz))


def ingest_formulas(tsh, engine, formula_file):
    df = pd.read_csv(formula_file)
    with engine.begin() as cn:
        for row in df.itertuples():
            tsh.register_formula(
                cn,
                row.name,
                row.text,
                reject_unknown=False
            )


# track the threadpool activity when running with pytest -s
L = logging.getLogger('parallel')
L.addHandler(logging.StreamHandler())
L.setLevel(logging.INFO)


excel_connect.DIE_ON_ERROR = False
excel_connect._TESTS = True

DATADIR = Path(__file__).parent / 'data'
XLPATH = DATADIR / 'proto_validator.xlsx'
XLPATH_DEMAND = DATADIR / 'proto_demand.xlsm'
XLPATH_IMPORTER = DATADIR / 'proto_importer.xlsm'
XLPATH_PRIORITY = DATADIR / 'proto_priority.xlsx'
XLPATH_SELECT = DATADIR / 'proto_select_name.xlsx'
XLPATH_DELTA = DATADIR / 'proto_delta.xlsx'
XLPATH_XLWINGS_BUG = DATADIR / 'xlwings_bug.xlsx'

HERE = Path(__file__).parent

skipnoxlwing = partial(
    pytest.mark.skipif, xw is None,
    reason='xwlings not supported on linux'
)()


recentxlsversion = partial(
    pytest.mark.skipif, xw is not None and str(xw.App().version) == '16.0',
    reason='Excel version is too recent: problem with datetime parsing'
)


find_zones = partial(
    find_zones_having_keyword,
    predicate=lambda name, part: part in name and 'Mercure' not in name
)

# don't let all these things dangling open after the tests
xlbook = partial(xlbook, close=True)


def compare_wb(wb1, tab1, wb2, tab2, from_row, to_row, from_col, to_col):
    sheet1 = wb1.sheets[tab1]
    sheet2 = wb2.sheets[tab2]
    extract1 = sheet1.range((from_row, from_col), (to_row, to_col)).value
    extract2 = sheet2.range((from_row, from_col), (to_row, to_col)).value
    for idx in range(len(extract1)):
        assert extract1[idx] == extract2[idx]


def init_df(wb, sheetname, name, first_date, header):
    sheet = wb.sheets[sheetname]
    coord = find_zones(wb, [name])[name]
    top, down, left, right = coord['top'], coord['down'], coord['left'], coord['right']
    height = down - top + 1
    width = right - left + 1

    df = np.array([range(idx, idx + width)
                   for idx in range(height)])

    sheet.range(coord['range']).value = df

    dates = [first_date + timedelta(days=idx) for idx in range(height)]
    dates = [[elt] for elt in dates]

    sheet.range((top, left - 1), (down, left - 1)).value = dates
    columnsrange = sheet.range((top - 1, left), (top - 1, right))
    columnsrange.value = header

    return columnsrange


def parse_df(wb, zones, sheetname=None, index=False):
    dico_output = {}


    for name, value in zones.items():
        sheet_from_dico = value['sheetname']
        if sheetname and sheetname != sheet_from_dico:
            continue
        sheet = wb.sheets[sheet_from_dico]
        coord = zones[name]
        if index:
            extract = sheet.range((coord['top'], coord['left'] - 1),
                                 (coord['down'], coord['right'])).value
            data_extract = pd.DataFrame(extract) if isinstance(extract, list) else pd.DataFrame([extract])
            cell_header = sheet.range(
                (coord['top'] - 1, coord['left']),
                (coord['top'] - 1, coord['right'])
            ).options(ndim=1).value
            cell_header = int_to_str(cell_header)
            data_extract.set_index(0, inplace=True)
            dico_col = {data_extract.columns[idx]: str(cell_header[idx])
                        for idx in range(data_extract.shape[1])}
            data_extract.rename(columns=dico_col, inplace=True)
        else:
            extract = sheet.range(coord['range']).value
            data_extract = pd.DataFrame(extract) if isinstance(extract, list) else pd.DataFrame([extract])
        dico_output[name] = data_extract

    return dico_output


def change_date(wb, zones, name, nb):
    dico_rolling_df = parse_df(wb, zones, index=True)
    df = dico_rolling_df[name]
    new_date = df.index
    coord = zones[name]
    position = (coord['top'], coord['left'] - 1)
    for date in new_date:
        wb.sheets[coord['sheetname']].range(position).value = date + timedelta(days=nb)
        position = position[0] + 1, position[1]


@skipnoxlwing
def test_rolling(engine, tsh, excel):
    with xlbook(XLPATH) as wb:
        # test when one sheet name is included into another,
        # here, there is the tab 'gap' and 'ga'
        found_zone = find_zones_having_keyword(wb, [PULL], 'ga')
        assert 1 == len(found_zone)

        init_df(wb, 'rolling', 'rw_test', datetime(2012, 4, 5),
                ('bidule', 'chose', 'truc', 'machin'))
        zones = find_zones_having_keyword(wb, [PUSH, PULL])
        # test of proper writing in excel
        name = 'rw_test'
        df = parse_df(wb, {name: zones[name]}, sheetname=None, index=False)[name]
        assert """
     0     1     2     3
0  0.0   1.0   2.0   3.0
1  1.0   2.0   3.0   4.0
2  2.0   3.0   4.0   5.0
3  3.0   4.0   5.0   6.0
4  4.0   5.0   6.0   7.0
5  5.0   6.0   7.0   8.0
6  6.0   7.0   8.0   9.0
7  7.0   8.0   9.0  10.0
8  8.0   9.0  10.0  11.0
9  9.0  10.0  11.0  12.0""".strip() == df.to_string().strip()

        parse_ts(wb, {name: zones[name]})

        push_rolling(wb, {'rw_test': zones['rw_test']})
        change_date(wb, zones, 'rw_test', nb=3)
        pull_rolling(wb, {'rw_test': zones['rw_test']})
        name = 'rw_test'

        df = parse_df(wb, {name: zones[name]}, sheetname=None, index=False)[name]

        assert"""
     0     1     2     3
0  3.0   4.0   5.0   6.0
1  4.0   5.0   6.0   7.0
2  5.0   6.0   7.0   8.0
3  6.0   7.0   8.0   9.0
4  7.0   8.0   9.0  10.0
5  8.0   9.0  10.0  11.0
6  9.0  10.0  11.0  12.0
7  NaN   NaN   NaN   NaN
8  NaN   NaN   NaN   NaN
9  NaN   NaN   NaN   NaN""".strip() == df.to_string().strip()

        ts_1 = pd.Series(
            range(31),
            index=pd.date_range(
                start=('2012-4-1'),
                end=('2012-5-1'),
                freq='D'
            )
        )
        ts_2 = pd.Series(
            range(5, 36),
            index=pd.date_range(
                start=('2012-4-1'),
                end=('2012-5-1'),
                freq='D'
            )
        )
        ts_3 = pd.Series(
            range(10, 41),
            index=pd.date_range(
                start=('2012-4-1'),
                end=('2012-5-1'),
                freq='D'
            )
        )
        ts_4 = pd.Series(
            range(15, 46),
            index=pd.date_range(
                start=('2012-4-1'),
                end=('2012-5-1'),
                freq='D'
            )
        )

        tsh.update(engine, ts_1, 'ts_1', 'test')
        tsh.update(engine, ts_2, 'ts_2', 'test')
        tsh.update(engine, ts_3, 'ts_3', 'test')
        tsh.update(engine, ts_4, 'ts_4', 'test')

        ts_begin = pd.Series([2] * 5)
        ts_begin.index = pd.date_range(start=datetime(2010, 1, 1), freq='D', periods=5)
        ts_begin.loc['2010-01-04'] = -1
        tsh.update(engine, ts_begin, 'ts_mixte', 'test')

        name = 'rw_screwed'
        df = parse_df(wb, {name: zones[name]}, sheetname=None, index=False)[name]
        assert """
       0     1     2     3     4     5
0    NaN   NaN   NaN   NaN  None  None
1    NaN   NaN   NaN   NaN  None  None
2    NaN   NaN   NaN   NaN  None  None
3    NaN   NaN   NaN   NaN  None  None
4    NaN   NaN   NaN   NaN  None  None
5    7.0  12.0  17.0  22.0  None  None
6    8.0  13.0  18.0  23.0  None  None
7    9.0  14.0  19.0  24.0  None  None
8   10.0  15.0  20.0  25.0  None  None
9   11.0  16.0  21.0  26.0  None  None
10  12.0  17.0  22.0  27.0  None  None
11  13.0  18.0  23.0  28.0  None  None
12  14.0  19.0  24.0  29.0  None  None
13  15.0  20.0  25.0  30.0  None  None
14  16.0  21.0  26.0  31.0  None  None
15   NaN   NaN   NaN   NaN  None  None
16   NaN   NaN   NaN   NaN  None  None
17   NaN   NaN   NaN   NaN  None  None
18   NaN   NaN   NaN   NaN  None  None
19   NaN   NaN   NaN   NaN  None  None
20   NaN   NaN   NaN   NaN  None  None
            """.strip() == df.to_string().strip()

        # Test with unusual case:
        # 1) missing column name
        init_df(wb, 'rolling', 'rw_test', datetime(2012, 4, 5),
                ('bidule', '', 'truc', 'machin'))
        zones = find_zones(wb, ['rolling'])
        push_rolling(wb, zones)
        pull_rolling(wb, zones)
        # 2) pull with column name absent from database
        init_df(wb, 'rolling', 'rw_test', datetime(2012, 4, 5),
                ('bidule', 'chose', 'truc', 'machin'))
        zones = find_zones(wb, ['rolling'])
        push_rolling(wb, zones)

        init_df(wb, 'rolling', 'rw_test', datetime(2012, 4, 5),
                ('bidule', 'intruder', 'truc', 'machin'))
        zones = find_zones(wb, ['rolling'])
        pull_rolling(wb, zones)

        init_df(wb, 'rolling', 'rw_test', datetime(2012, 4, 5),
                ('bidule', 'chose', 'truc'))
        zones = find_zones_having_keyword(wb, [PULL, PUSH])

        init_df(wb, 'rolling', 'rw_test', datetime(2012, 4, 5),
                ('bidule', 'chose', 'truc', 'machin'))

        change_date(wb, zones, 'rw_test', nb=3)

        wb.sheets['other_sheet'].range((7, 12), (27, 17)).value = None

        pull_rolling(wb, {'rw_screwed': zones['rw_screwed']})
        df = parse_df(wb,  {'rw_screwed': zones['rw_screwed']},
                      sheetname=None, index=False)['rw_screwed']

        assert """
       0     1     2     3
5    7.0  12.0  17.0  22.0
6    8.0  13.0  18.0  23.0
7    9.0  14.0  19.0  24.0
8   10.0  15.0  20.0  25.0
9   11.0  16.0  21.0  26.0
10  12.0  17.0  22.0  27.0
11  13.0  18.0  23.0  28.0
12  14.0  19.0  24.0  29.0
13  15.0  20.0  25.0  30.0
14  16.0  21.0  26.0  31.0
""".strip() == df.iloc[5:15, 0:4].to_string().strip()

        #    3) duplication of column: => raise an error
        undo = init_df(wb, 'rolling', 'rw_test', datetime(2012, 4, 5),
                       ('bidule', 'chose', 'truc', 'chose'))

        with pytest.raises(Exception):
            zones = find_zones_having_keyword(wb, [PULL, PUSH])

        undo.value = None


@skipnoxlwing
def test_revision_date_excel(engine, tsh, excel):
    ts = pd.Series([4] * 7,
                   index=pd.date_range(start=datetime(2010, 1, 4),
                                       freq='D', periods=7), name='truc')
    tsh.update(
        engine, ts, 'ts_constant', 'test',
        insertion_date=pd.Timestamp(datetime(2015, 1, 1, 15, 43, 23), tz='UTC')
    )

    ts = pd.Series([1] * 4,
                   index=pd.date_range(start=datetime(2010, 1, 4),
                                       freq='D', periods=4), name='truc')
    tsh.update(
        engine, ts, 'ts_through_time', 'test',
        insertion_date=pd.Timestamp(datetime(2015, 1, 1, 15, 43, 23), tz='UTC')
    )

    ts = pd.Series([2] * 4,
                   index=pd.date_range(start=datetime(2010, 1, 4),
                                       freq='D', periods=4), name='truc')
    tsh.update(
        engine, ts, 'ts_through_time', 'test',
        insertion_date=pd.Timestamp(datetime(2015, 1, 2, 15, 43, 23), tz='UTC')
    )


    ts = pd.Series([3] * 4,
                   index=pd.date_range(start=datetime(2010, 1, 4),
                                       freq='D', periods=4), name='truc')
    tsh.update(
        engine, ts, 'ts_through_time', 'test',
        insertion_date=pd.Timestamp(datetime(2015, 1, 3, 15, 43, 23), tz='UTC')
    )

    ingest_formulas(tsh, engine,  DATADIR / 'formula_definitions.csv')

    with xlbook(XLPATH) as wb:
        zones = find_zone_by_name(wb, ['rc_past'])
        name = 'rc_past'
        corner = zones[name]['top'] - 1, zones[name]['left'] - 1
        sheetname = zones[name]['sheetname']

        wb.sheets[sheetname].range(corner).value = datetime(2015, 1, 4)
        zones = find_zone_by_name(wb, ['rc_past'])
        pull_rolling(wb, zones)
        df = parse_df(wb, zones, sheetname=None, index=False)[name]
        assert """
     0    1
0  3.0  3.0
1  3.0  3.0
2  3.0  3.0
3  3.0  3.0
4  NaN  4.0
5  NaN  4.0
6  NaN  4.0
""".strip() == df.to_string().strip()

        wb.sheets[sheetname].range(corner).value = datetime(2015, 1, 3, 1, 12, 35)
        zones = find_zone_by_name(wb, ['rc_past'])
        pull_rolling(wb, zones)
        df = parse_df(wb, zones, sheetname=None, index=False)[name]
        assert """
     0    1
0  2.0  2.0
1  2.0  2.0
2  2.0  2.0
3  2.0  2.0
4  NaN  4.0
5  NaN  4.0
6  NaN  4.0
""".strip() == df.to_string().strip()

        wb.sheets[sheetname].range(corner).value = datetime(2015, 1, 2)
        zones = find_zone_by_name(wb, ['rc_past'])
        pull_rolling(wb, zones)
        df = parse_df(wb, zones, sheetname=None, index=False)[name]
        assert """
     0    1
0  1.0  1.0
1  1.0  1.0
2  1.0  1.0
3  1.0  1.0
4  NaN  4.0
5  NaN  4.0
6  NaN  4.0
""".strip() == df.to_string().strip()

        #here we pull the same series in two different names, one with the rev_date,
        # the other without
        zones = find_zones_having_keyword(wb, [PULL], sheetname = 'through_time')
        pull_rolling(wb, zones)

        df = parse_df(wb, zones, sheetname=None, index=False)['rc_past']
        assert """
     0    1
0  1.0  1.0
1  1.0  1.0
2  1.0  1.0
3  1.0  1.0
4  NaN  4.0
5  NaN  4.0
6  NaN  4.0
        """.strip() == df.to_string().strip()

        df = parse_df(wb, zones, sheetname=None, index=False)['rc_now']
        assert """
     0
0  3.0
1  3.0
2  3.0
3  3.0
4  NaN
5  NaN
6  NaN
        """.strip() == df.to_string().strip()

        df = parse_df(wb, zones, sheetname=None, index=False)['rwc_multiple']
        assert """
     0    1    2    3
0  3.0  1.0  3.0  1.0
1  3.0  1.0  3.0  1.0
2  3.0  1.0  3.0  1.0
3  3.0  1.0  3.0  1.0
4  NaN  NaN  4.0  4.0
5  NaN  NaN  4.0  4.0
6  NaN  NaN  4.0  4.0
        """.strip() == df.to_string().strip()

        # before the first insertion
        wb.sheets[sheetname].range(corner).value = datetime(2015, 1, 1)
        zones = find_zone_by_name(wb, ['rc_past'])
        pull_rolling(wb, zones)
        df = parse_df(wb, zones, sheetname=None, index=False)[name]
        assert """
      0     1
0  None  None
1  None  None
2  None  None
3  None  None
4  None  None
5  None  None
6  None  None
        """.strip() == df.to_string().strip()

        wb.sheets[sheetname].range(corner).value = None


@skipnoxlwing
def test_name():
    with xlbook(XLPATH) as wb:
        zones = find_zones_having_keyword(wb, [PULL])
        assert set(['rw_test', 'rw_digit', 'rw_fill_the_gap', 'rwc_inside', 'rc_now',
                    'rw_manual', 'rc_past', 'rw_screwed', 'rwc_multiple']) == set(zones.keys())
        zones_pull = find_zones_having_keyword(wb, [PULL])
        zones_push = find_zones_having_keyword(wb, [PUSH])
        assert set(['rc_past', 'rc_now']) == set(zones_pull.keys()) - set(zones_push.keys())

        sheetname = 'rolling'
        zones = find_zones_having_keyword(wb, [PULL], sheetname=sheetname)
        assert set(['rw_test']) == set(zones.keys())


@skipnoxlwing
def test_manual(engine, tsh, excel):
    with xlbook(XLPATH) as wb:
        init_df(wb, 'manual', 'rw_manual', datetime(2010, 1, 1),
                ('ts_edited', 'ts_1'))

        ts_begin = pd.Series([2] * 5)
        ts_begin.index = pd.date_range(start=datetime(2010, 1, 1), freq='D', periods=5)
        ts_begin.loc['2010-01-04'] = -1
        tsh.update(engine, ts_begin, 'ts_edited', 'test')

        ts_more = pd.Series([2] * 5)
        ts_more.index = pd.date_range(start=datetime(2010, 1, 2), freq='D', periods=5)
        ts_more.loc['2010-01-04'] = -1
        tsh.update(engine, ts_more, 'ts_edited', 'test')

        ts_1 = pd.Series(
            range(31),
            index=pd.date_range(start=('2010-1-1'), end=('2010-1-31'), freq='D')
        )
        tsh.update(engine, ts_1, 'ts_1', 'test')

        ts_manual = pd.Series([4] * 2)
        ts_manual.index = pd.date_range(start=datetime(2010, 1, 3), freq='D', periods=2)
        tsh.update(engine, ts_manual, 'ts_1', 'test', manual=True)

        ts_one_more = pd.Series([2])  # with no intersection with the previous ts
        ts_one_more.index = pd.date_range(start=datetime(2010, 1, 7), freq='D', periods=1)
        tsh.update(engine, ts_one_more, 'ts_edited', 'test')

        print('INSERT MANUAL STUFF')
        ts_manual = pd.Series([3])
        ts_manual.index = pd.date_range(start=datetime(2010, 1, 4), freq='D', periods=1)
        tsh.update(engine, ts_manual, 'ts_edited', 'test', manual=True)

        zones = find_zones_having_keyword(wb, [PULL], 'manual')
        pull_rolling(wb, zones)

        assert """
2010-01-01    2.0
2010-01-02    2.0
2010-01-03    2.0
2010-01-04    3.0
2010-01-05    2.0
2010-01-06    2.0
2010-01-07    2.0
""".strip() == tsh.get(engine, 'ts_edited').to_string().strip()

        zones = find_zones_having_keyword(wb, [PULL], 'manual')
        df = parse_df(wb, zones, sheetname=None, index=False)['rw_manual']
        assert """
     0    1
0  2.0  0.0
1  2.0  1.0
2  2.0  4.0
3  3.0  4.0
4  2.0  4.0
5  2.0  5.0
6  2.0  6.0""".strip() == df.to_string().strip()


def closebook(xlpath):
    try:
        xw.Book(str(xlpath)).close()
    except:
        pass


@skipnoxlwing
def test_compress():
    ts = pd.Series(['a'] * 4 + ['b'] * 1 + ['c'] * 3 + ['d'] * 2 ,
                   index=pd.date_range(start=('2015-1-1'), end=('2015-1-10'), freq='D'))
    result_list = compress_series(ts)
    assert result_list == [
        ('a', (0, 3)),
        ('b', (4, 4)),
        ('c', (5, 7)),
        ('d', (8, 9))
    ]

    ts = pd.Series(['a'] * 10 ,
                   index=pd.date_range(start=('2015-1-1'), end=('2015-1-10'), freq='D'))
    result_list = compress_series(ts)
    assert [('a', (0, 9))] == result_list

    ts = pd.Series(['a'] * 1 + ['b'] * 2 + ['c'] * 1 ,
                   index=pd.date_range(start=('2015-1-1'), end=('2015-1-4'), freq='D'))
    result_list = compress_series(ts)
    assert [
        ('a', (0, 0)),
        ('b', (1, 2)),
        ('c', (3, 3))
    ] == result_list

    ts = pd.Series([None] * 2 + ['b'] * 2 + [None] * 2 +  ['c'] * 3 + [None],
                   index=pd.date_range(start=('2015-1-1'), end=('2015-1-10'), freq='D'))
    result_list = compress_series(ts)
    assert [
        ('b', (2, 3)),
        ('c', (6, 8))
    ] == result_list

    ts = pd.Series([np.nan] * 2 + ['b'] * 2 + [np.nan] * 2 + ['c'] * 3 + [np.nan],
                   index=pd.date_range(start=('2015-1-1'), end=('2015-1-10'), freq='D'))
    result_list = compress_series(ts)
    assert [
               ('b', (2, 3)),
               ('c', (6, 8))
           ] == result_list


@skipnoxlwing
def test_combination(excel, tsh, engine):
    ingest_formulas(tsh, engine,  DATADIR / 'formula_definitions.csv')
    with xlbook(XLPATH_PRIORITY) as wb:
        zones = find_zones(wb, ['combined'], sheetname='Sheet2')
        zones = find_zones(wb, ['creation'], sheetname='Sheet1')
        push_rolling(wb, zones)

        zones = find_zones(wb, ['combined'], sheetname='Sheet2')
        pull_rolling(wb, zones)

        bar = parse_df(wb, zones)['rwc_combined']
        assert """
      0     1    2
0   1.0  None  4.0
1   1.0  None  4.0
2   1.0  None  4.0
3   1.0  None  4.0
4   1.0  None  4.0
5   1.0  None  4.0
6   1.0  None  4.0
7   1.0  None  4.0
8   1.0  None  6.0
9   1.0  None  6.0
10  2.0  None  6.0
11  2.0  None  6.0
12  3.0  None  6.0
13  3.0  None  6.0
14  3.0  None  6.0
15  3.0  None  6.0
16  3.0  None  6.0
17  3.0  None  6.0
18  3.0  None  6.0
""".strip() == bar.to_string().strip()


@skipnoxlwing
def test_coef(engine, tsh, excel):
    ingest_formulas(tsh, engine,  DATADIR / 'formula_definitions.csv')
    with xlbook(XLPATH_PRIORITY) as wb:
        prems = pd.Series(
            [1] * 14,
            index=pd.date_range(start=('2015-1-1'), end=('2015-1-14'), freq='D')
        )
        deuz = pd.Series(
            [2000] * 16,
            index=pd.date_range(start=('2015-1-1'), end=('2015-1-16'), freq='D')
        )
        troiz = pd.Series(
            [3000000] * 18,
            index=pd.date_range(start=('2015-1-1'), end=('2015-1-18'), freq='D')
        )
        vaifoif = pd.Series(
            [1] * 18,
            index=pd.date_range(start=('2015-1-1'), end=('2015-1-18'), freq='D')
        )

        tsh.update(engine, prems, 'prems', 'test')
        tsh.update(engine, deuz, 'deuz', 'test')
        tsh.update(engine, troiz, 'troiz', 'test')
        tsh.update(engine, vaifoif, 'vaifoif', 'test')

        zones = find_zones(wb, ['rwc_compo'], sheetname='compo full')
        pull_rolling(wb, zones)
        bar = parse_df(wb, zones)['rwc_compo']
        assert """
      0     1     2       3
0   NaN  None  None     NaN
1   1.0  None  None  1000.0
2   1.0  None  None  1000.0
3   1.0  None  None  1000.0
4   1.0  None  None  1000.0
5   1.0  None  None  1000.0
6   1.0  None  None  1000.0
7   1.0  None  None  1000.0
8   1.0  None  None  1000.0
9   1.0  None  None  1000.0
10  NaN  None  None     NaN
11  1.0  None  None  1000.0
12  1.0  None  None  1000.0
13  1.0  None  None  1000.0
14  1.0  None  None  1000.0
15  1.0  None  None  1000.0
16  1.0  None  None  2000.0
17  1.0  None  None  2000.0
18  1.0  None  None  3000.0
19  1.0  None  None  3000.0
    """.strip() == bar.to_string().strip()


@skipnoxlwing
def test_na_col(engine, tsh, excel):
    without_na = pd.Series(
        [1] * 5,
        index=pd.date_range(start=('2015-1-1'), end=('2015-1-5'), freq='D')
    )
    with_na = pd.Series(
        [2] * 5,
        index=pd.date_range(start=('2015-1-1'), end=('2015-1-5'), freq='D')
    )
    total = pd.Series(
        [3] * 5,
        index=pd.date_range(start=('2015-1-1'), end=('2015-1-5'), freq='D')
    )
    with_zero = pd.Series(
        [4] * 5,
        index=pd.date_range(start=('2015-1-1'), end=('2015-1-5'), freq='D')
    )
    with_prev = pd.Series(
        [5] * 5,
        index=pd.date_range(start=('2015-1-1'), end=('2015-1-5'), freq='D')
    )

    tsh.update(engine, without_na, 'without_na', 'test')
    tsh.update(engine, with_na, 'with_na', 'test')
    tsh.update(engine, total, 'total', 'test')
    tsh.update(engine, with_zero, 'with_zero', 'test')
    tsh.update(engine, with_prev, 'with_prev', 'test')

    with xlbook(XLPATH_PRIORITY) as wb:
        zones = find_zones(wb, ['rwc_test_na'], sheetname='na')
        pull_rolling(wb, zones)
        push_rolling(wb, zones)
        pull_rolling(wb, zones)

        result = parse_df(wb, zones)['rwc_test_na']

    assert """
       0    1    2     3    4    5    6
0   None  NaN  NaN  None  NaN  NaN  NaN
1   None  1.0  2.0  None  3.0  4.0  5.0
2   None  1.0  2.0  None  3.0  4.0  5.0
3   None  NaN  NaN  None  NaN  NaN  NaN
4   None  1.0  2.0  None  3.0  4.0  5.0
5   None  1.0  2.0  None  3.0  4.0  5.0
6   None  1.0  2.0  None  3.0  4.0  5.0
7   None  NaN  NaN  None  NaN  0.0  5.0
8   None  NaN  NaN  None  NaN  0.0  5.0
9   None  NaN  NaN  None  NaN  0.0  5.0
10  None  NaN  NaN  None  NaN  NaN  NaN
11  None  NaN  NaN  None  NaN  0.0  NaN
12  None  NaN  NaN  None  NaN  0.0  NaN
""".strip() == result.to_string().strip()
    # the #N/A are interpreted as -2.146826e+09


@skipnoxlwing
def test_monthly(engine, excel):
    with xlbook(XLPATH_PRIORITY) as wb:
        zones = find_zones(wb, ['_month_'])
        push_rolling(wb, zones)

        zones = find_zones(wb, ['daily'])
        pull_rolling(wb, zones)

        #let's simulate an data "erasing"
        #before
        zones = find_zones(wb, ['rw_daily_beast'], sheetname='daily_test')
        df_result = parse_df(wb, zones)['rw_daily_beast']
        assert 153 == sum(~df_result[0].isnull())

        # add an Nan and push monthly, pull daily
        wb.sheets['monthly_test'].range((10,4)).value = np.nan

        zones = find_zones(wb, ['_month_'])
        push_rolling(wb, zones)
        zones = find_zones(wb, ['daily'])
        pull_rolling(wb, zones)

        #after
        zones = find_zones(wb, ['rw_daily_beast'], sheetname='daily_test')
        df_result = parse_df(wb, zones)['rw_daily_beast']
        assert 123 == sum(~df_result[0].isnull())
        # we succesuly remove 30 points



@skipnoxlwing
def test_extrapole_fill(engine, tsh, excel):
    with xlbook(XLPATH_PRIORITY) as wb:
        ts_num = pd.Series(range(5),
                           index=pd.date_range(start=('2015-1-1'),
                                               end=('2015-1-5'),
                                               freq='D'))

        ts_str = pd.Series(['toto'] * 5,
                           index=pd.date_range(start=('2015-1-1'),
                                               end=('2015-1-5'),
                                               freq='D'))

        ts_num_a = ts_num.copy()
        ts_num_a.iloc[3] = np.nan
        ts_str.iloc[4] = np.nan
        tsh.update(engine, ts_num_a, 'test_fill_a', 'test')
        tsh.update(engine, ts_str, 'test_fill_b', 'test')
        tsh.update(engine, ts_num, 'test_fill_c', 'test')
        zones = find_zones(wb, ['testfill'])

        # a warning is produced by this line, but we don't know why yet
        pull_rolling(wb, zones)

        push_rolling(wb, zones)
        df = parse_df(wb, zones)['rwcf_testfill']

        assert """
     0     1     2    3
0  NaN  None  None  NaN
1  0.0  toto  None  0.0
2  1.0  toto  None  1.0
3  2.0  toto  None  2.0
4  NaN  None  None  NaN
5  NaN  toto  None  3.0
6  4.0  None  None  4.0
7  4.0  None  None  4.0
8  4.0  None  None  4.0
9  4.0  None  None  4.0 """.strip() == df.to_string().strip()


@skipnoxlwing
def test_gap_filling(excel):
    with xlbook(XLPATH) as wb:
        zones = find_zones(wb, ['fill_the_gap'])
        push_rolling(wb, zones)
        pull_rolling(wb, zones)

        df = parse_df(wb, zones)['rw_fill_the_gap']

        assert """
       0     1     2     3     4     5
0   blob   1.0  blob   3.0   4.0   5.0
1   blob   2.0  blob   4.0   5.0   6.0
2   truc   7.0   7.0  truc  truc  truc
3   blob   7.0   7.0   6.0   7.0   8.0
4   blob   5.0  blob   7.0   8.0   9.0
5   blob   6.0  blob   8.0   9.0  10.0
6   blob   7.0  blob   9.0  10.0  11.0
7   blob   8.0  blob  10.0  11.0  12.0
8   blob   9.0  blob  11.0  12.0  13.0
9   blob  10.0  blob  12.0  13.0  14.0
10  blob  11.0  blob  13.0  14.0  15.0
11  truc  truc  truc  truc  truc  truc""".strip() == df.to_string().strip()


@skipnoxlwing
def test_exotic_excel(excel):
    with xlbook(XLPATH) as wb:
        zones = find_zones(wb, ['exotic1'])
        push_rolling(wb, zones)
        zones = find_zones(wb, ['exotic2'])
        pull_rolling(wb, zones)


@skipnoxlwing
def test_data_type(engine, tsh, excel):
    with xlbook(XLPATH_PRIORITY) as wb:
        zones = find_zones(wb, ['rwc_test_type'], sheetname='type')

        push_rolling(wb, zones)
        ts1 = tsh.get(engine, 'ts_type1')
        ts2 = tsh.get(engine, 'ts_type2')
        ts3 = tsh.get(engine, 'ts_type3')
        ts4 = tsh.get(engine, 'ts_type4')

        zones = find_zones(wb, ['rwc_type_arriva'], sheetname='type')
        pull_rolling(wb, zones)

        assert ts1.dtype == 'float64'
        assert ts2.dtype == 'float64'
        assert ts3.dtype == 'object'  # there are legit string in the serie
        assert ts4.dtype == 'float64'


@skipnoxlwing
def test_fill_blank(engine, tsh, excel):
    # NOTE: this test should be redone
    with xlbook(XLPATH_PRIORITY) as wb:
        ts = pd.Series([1] * 6, index=pd.date_range(start=('2015-1-1'),
                                                       end=('2015-1-6'), freq='D'))
        ts = ts[[0, 1, 2, 4, 5]]
        tsh.update(engine, ts, 'rom1', 'test')
        tsh.update(engine, ts, 'rom2', 'test')
        tsh.update(engine, ts, 'rom3', 'test')
        tsh.update(engine, ts, 'rom4', 'test')
        zones = find_zones(wb, ['rwc_ro_na_departure'], sheetname='ro_mercure')
        pull_rolling(wb, zones)
        wb.sheets['ro_mercure'].range((11, 3), (11, 6)).value = -2
        push_rolling(wb, zones)

        zones = find_zones(wb, ['rwc_ro_na_arrival'], sheetname='ro_mercure')
        pull_rolling(wb, zones)
        df_result = parse_df(wb, zones)['rwc_ro_na_arrival']
        assert """
        0       1       2       3
0  filler  filler  filler  filler
1     1.0     1.0     1.0     1.0
2     1.0     1.0     1.0     1.0
3    None    None    None    None
4     1.0     1.0     1.0     1.0
5    None    None    None    None
6    -2.0    -2.0     1.0     1.0
7     1.0     1.0     1.0     1.0
""".strip() == df_result.to_string().strip()


@skipnoxlwing
def test_find_zone():
    with xlbook(XLPATH_SELECT) as wb:
        # we didnt't touch the excel test files
        # from the time the default exclusion term was 'request'
        zones = find_zones_having_keyword(wb, [PULL], sheetname='tab1')
        assert {
            'rwc_custom1', 'rwc_request_tab1', 'rwc_custom2'
        } == set(zones.keys())
        zones = find_zones_having_keyword(wb, [PULL])
        assert {
            'rwc_custom1', 'rwc_cutsom3', 'rwc_request_tab1',
            'rwc_request_tab3', 'rwc_custom2'
        } == set(zones.keys())


@skipnoxlwing
def test_empty_ts(engine, tsh, excel):
    with xlbook(XLPATH_PRIORITY) as wb:
        ts_full = pd.Series([1] * 11, index=pd.date_range(start=('2015-1-1'),
                                                           end=('2015-1-11'), freq='D'))

        tsh.update(engine, ts_full, 'ts_full', 'test')
        zone = find_zones(wb, ['rwc_empty'], sheetname= 'empty')
        pull_rolling(wb, zone)

        zone = find_zones(wb, ['rwc_semi_empty'], sheetname= 'empty')
        pull_rolling(wb, zone)


@skipnoxlwing
def test_xlwings_bug(excel):
    import datetime
    with xlbook(XLPATH_XLWINGS_BUG) as xl_bug:
        sheet = xl_bug.sheets[0]
        timestamps = sheet["A2:A33"].options(ndim=1).value
        assert timestamps[1] == datetime.datetime(2015, 1, 1, 0, 0)
        assert timestamps[3] == datetime.datetime(2015, 1, 1, 1, 59, 59, 990000)


@skipnoxlwing
@recentxlsversion
def test_delta_excel(engine, tsh, excel):
    ingest_formulas(tsh, engine,  DATADIR / 'formula_definitions.csv')

    with xlbook(XLPATH_DELTA) as wb:
        for insertion_date in pd.date_range(
                start=datetime(2015, 1, 1), end=datetime(2015, 1, 2), freq='H'
        ):
            ts1 = genserie(start=insertion_date, freq='H', repeat=6)
            ts2 = ts1 + 1
            tsh.update(
                engine, ts1, 'rep1', 'test',
                insertion_date=pd.Timestamp(insertion_date,tz='UTC')
            )
            tsh.update(
                engine, ts2, 'rep2', 'test',
                insertion_date=pd.Timestamp(insertion_date,tz='UTC')
            )

        zone = find_zones(wb, ['rwc_delta'], sheetname= 'delta')

        pull_rolling(wb, zone)
        df = parse_df(wb, zone, 'delta')['rwc_delta'].iloc[-11:-1,:]

        assert """
      0    1     2     3     4
21  3.0  4.0  None  24.0  47.0
22  3.0  4.0  None  24.0  47.0
23  3.0  4.0  None  24.0  47.0
24  3.0  4.0  None  24.0  47.0
25  3.0  4.0  None  24.0  47.0
26  3.0  4.0  None  24.0  47.0
27  3.0  4.0  None  24.0  47.0
28  3.0  4.0  None  24.0  47.0
29  4.0  5.0  None  32.0  59.0
30  5.0  6.0  None  40.0  71.0""".strip() == df.to_string().strip()


@skipnoxlwing
def test_aggregation_excel(engine, tsh, excel):
    with xlbook(XLPATH_DELTA) as wb:
        tsh.update(engine, genserie(start=datetime(2015,1,1), freq='H', repeat=2880),
                   'hourly_series_1', 'test')
        tsh.update(engine, genserie(start=datetime(2015,1,1), freq='H', repeat=3500),
                   'hourly_series_2', 'test')
        tsh.update(engine, genserie(start=datetime(2015,1,1), freq='D', repeat=220),
                   'daily_series', 'test')
        tsh.update(engine, genserie(start=datetime(2015,1,1), freq='D', repeat=300),
                   'monthly_series', 'test')

        zone = find_zones(wb, 'rwc_aggregation_daily', sheetname='aggregation')
        pull_rolling(wb, zone)
        df1 = parse_df(wb, zone, 'aggregation')['rwc_aggregation_daily']

        zone = find_zones(wb, 'rwc_aggregation_monthly', sheetname='aggregation')
        pull_rolling(wb, zone)
        df2 = parse_df(wb, zone, 'aggregation')['rwc_aggregation_monthly']

        assert """
       0       1     2      3    4
0   None     NaN  None    NaN  NaN
1   None     NaN  None    NaN  NaN
2   None   276.0  None   11.5  0.0
3   None   852.0  None   35.5  1.0
4   None  1428.0  None   59.5  2.0
5   None  2004.0  None   83.5  3.0
6   None  2580.0  None  107.5  4.0
7   None  3156.0  None  131.5  5.0
8   None  3732.0  None  155.5  6.0
9   None  4308.0  None  179.5  7.0
10  None     NaN  None    NaN  8.0""".strip() == df1.to_string().strip()

        assert """
         0       1       2      3
0    743.0     0.0   465.0    0.0
1   1415.0   744.0  1246.0   31.0
2   2159.0  1416.0  2294.0   59.0
3   2879.0  2160.0  3135.0   90.0
4      NaN  2880.0  4185.0  120.0
5      NaN     NaN  4965.0  151.0
6      NaN     NaN     NaN  181.0
7      NaN     NaN     NaN    NaN
8      NaN     NaN     NaN    NaN
9      NaN     NaN     NaN    NaN
10     NaN     NaN     NaN    NaN""".strip() == df2.to_string().strip()
