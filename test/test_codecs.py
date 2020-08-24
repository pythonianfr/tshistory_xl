from datetime import datetime, timedelta

import pandas as pd
from tshistory.testutil import utcdt, assert_df

from tshistory_xl import codecs


def test_pack_unpack_insert():
    series = pd.Series(
        [1., 2., 3.],
        index=pd.date_range(utcdt(2020, 1, 1), freq='H', periods=3)
    )

    packed = codecs.pack_insert_series(
        'Babar', [('a', series), ('b', series)]
    )
    author, allseries = codecs.unpack_insert_series(packed)
    assert_df("""
2020-01-01 00:00:00    1.0
2020-01-01 01:00:00    2.0
2020-01-01 02:00:00    3.0
""", allseries[0])
    assert_df("""
2020-01-01 00:00:00    1.0
2020-01-01 01:00:00    2.0
2020-01-01 02:00:00    3.0
""", allseries[1])

    assert allseries[0].name == 'a'
    assert allseries[1].name == 'b'
    assert author == 'Babar'


def test_pack_unpack_getmany():
    series = pd.Series(
        [1., 2., 3.],
        index=pd.date_range(utcdt(2020, 1, 1), freq='H', periods=3)
    )
    markers = pd.Series(
        [True, False, True],
        index=pd.date_range(utcdt(2020, 1, 1), freq='H', periods=3)
    )
    output = [
        ('zone', 'forecast', pd.Timestamp('2020-1-1'),
         (series, None, markers))
    ]
    packed = codecs.pack_getmany(output)
    unpacked = codecs.unpack_getmany(packed)
    item = unpacked[0]
    assert item[:3] == (
        'zone', 'forecast', pd.Timestamp('2020-1-1')
    )
    base, origin, marker = item[3]
    assert_df("""
2020-01-01 00:00:00    1.0
2020-01-01 01:00:00    2.0
2020-01-01 02:00:00    3.0
""", base)
    assert len(origin) == 0
    assert_df("""
2020-01-01 00:00:00     True
2020-01-01 01:00:00    False
2020-01-01 02:00:00     True
""", marker)


def test_pack_unpack_getmany_request():
    queries = [
        ('rw_test', 'bidule',
         {'delta': timedelta(days=1),
          'from_value_date': datetime(2012, 4, 5, 0, 0),
          'name': 'bidule',
          'revision_date': None,
          'to_value_date': datetime(2012, 4, 14, 0, 0)
         }
        )]

    packed = codecs.pack_getmany_request(queries)
    unpacked = codecs.unpack_getmany_request(packed)
    assert unpacked == [
        ['rw_test', 'bidule', {
            'delta': timedelta(days=1),
            'from_value_date': pd.Timestamp('2012-04-05 00:00:00'),
            'name': 'bidule',
            'revision_date': None,
            'to_value_date': pd.Timestamp('2012-04-14 00:00:00')
        }]
    ]
