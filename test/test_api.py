from datetime import datetime

import pytest
import pandas as pd

from tshistory import api
from tshistory.testutil import (
    assert_df,
    genserie,
    make_tsx
)
from tshistory.schema import tsschema
from tshistory_formula.schema import formula_schema

from tshistory_xl.testutil import with_http_bridge
from tshistory_xl import tsio, http


def _initschema(engine, ns='tsh'):
    tsschema(ns).create(engine)
    tsschema(ns + '-upstream').create(engine)
    formula_schema(ns).create(engine)


def make_api(engine, ns, sources=()):
    _initschema(engine, ns)
    for _uri, sns in sources.values():
        _initschema(engine, sns)

    return api.timeseries(
        str(engine.url),
        namespace=ns,
        handler=tsio.timeseries,
        sources=sources
    )


@pytest.fixture(scope='session')
def tsa1(engine):
    tsa = make_api(
        engine,
        'test-api',
        {'remote': (str(engine.url), 'test-remote')}
    )

    return tsa


@pytest.fixture(scope='session')
def tsa2(engine):
    ns = 'test-remote'
    tsschema(ns).create(engine)
    tsschema(ns + '-upstream').create(engine)
    formula_schema(ns).create(engine)
    dburi = str(engine.url)

    return api.timeseries(
        dburi,
        namespace=ns,
        handler=tsio.timeseries,
        sources={}
    )


tsx = make_tsx(
    'http://test.me',
    _initschema,
    tsio.timeseries,
    http.xl_httpapi,
    http.xl_httpclient,
    with_http_bridge=with_http_bridge
)


def test_get_many(tsx):
    for name in ('scalarprod', 'base'):
        tsx.delete(name)

    ts_base = genserie(datetime(2010, 1, 1), 'D', 3, [1])
    tsx.update('base', ts_base, 'test')

    tsx.register_formula(
        'scalarprod',
        '(* 2 (series "base"))'
    )

    v, m, o = tsx.values_markers_origins('scalarprod')
    assert m is None
    assert o is None
    assert_df("""
2010-01-01    2.0
2010-01-02    2.0
2010-01-03    2.0
""", v)

    # get_many, republications & revision date
    for idx, idate in enumerate(pd.date_range(datetime(2015, 1, 1),
                                              datetime(2015, 1, 3),
                                              freq='D',
                                              tz='utc')):
        tsx.update('comp1', ts_base * idx, 'test',
                   insertion_date=idate)
        tsx.update('comp2', ts_base * idx, 'test',
                   insertion_date=idate)

    tsx.register_formula(
        'repusum',
        '(add (series "comp1") (series "comp2"))'
    )

    tsx.register_formula(
        'repuprio',
        '(priority (series "comp1") (series "comp2"))'
    )

    lastsum, _, _ = tsx.values_markers_origins('repusum')
    assert_df("""
2010-01-01    4.0
2010-01-02    4.0
2010-01-03    4.0
""", lastsum)

    pastsum, _, _ = tsx.values_markers_origins(
        'repusum',
        revision_date=datetime(2015, 1, 2, 18)
    )
    assert_df("""
2010-01-01    2.0
2010-01-02    2.0
2010-01-03    2.0
""", pastsum)

    lastprio, _, _ = tsx.values_markers_origins('repuprio')
    assert_df("""
2010-01-01    2.0
2010-01-02    2.0
2010-01-03    2.0
""", lastprio)

    print('*' * 50)
    pastprio, _, _ = tsx.values_markers_origins(
        'repuprio',
        revision_date=datetime(2015, 1, 2, 18)
    )
    assert_df("""
2010-01-01    1.0
2010-01-02    1.0
2010-01-03    1.0
""", pastprio)


def test_get_many_federated(tsa1, tsa2):
    # same test as above
    # tsa1: local with remote source
    # tsa2: remote source
    ts_base = genserie(datetime(2010, 1, 1), 'D', 3, [1])
    tsa2.update('base', ts_base, 'test')

    tsa2.register_formula(
        'scalarprod',
        '(* 2 (series "base"))'
    )

    v, m, o = tsa1.values_markers_origins('scalarprod')
    assert m is None
    assert o is None
    assert_df("""
2010-01-01    2.0
2010-01-02    2.0
2010-01-03    2.0
""", v)

    # get_many, republications & revision date
    for idx, idate in enumerate(pd.date_range(datetime(2015, 1, 1),
                                              datetime(2015, 1, 3),
                                              freq='D',
                                              tz='utc')):
        tsa2.update('comp1', ts_base * idx, 'test',
                   insertion_date=idate)
        tsa2.update('comp2', ts_base * idx, 'test',
                   insertion_date=idate)

    tsa2.register_formula(
        'repusum',
        '(add (series "comp1") (series "comp2"))'
    )

    tsa2.register_formula(
        'repuprio',
        '(priority (series "comp1") (series "comp2"))'
    )

    lastsum, _, _ = tsa1.values_markers_origins('repusum')
    pastsum, _, _ = tsa1.values_markers_origins(
        'repusum',
        revision_date=datetime(2015, 1, 2, 18)
    )

    lastprio, _, _ = tsa1.values_markers_origins('repuprio')
    pastprio, _, _ = tsa1.values_markers_origins(
        'repuprio',
        revision_date=datetime(2015, 1, 2, 18)
    )

    assert_df("""
2010-01-01    4.0
2010-01-02    4.0
2010-01-03    4.0
""", lastsum)

    assert_df("""
2010-01-01    2.0
2010-01-02    2.0
2010-01-03    2.0
""", pastsum)

    assert_df("""
2010-01-01    2.0
2010-01-02    2.0
2010-01-03    2.0
""", lastprio)

    assert_df("""
2010-01-01    1.0
2010-01-02    1.0
2010-01-03    1.0
""", pastprio)


def test_origin(tsx):
    ts_real = genserie(datetime(2010, 1, 1), 'D', 10, [1])
    ts_nomination = genserie(datetime(2010, 1, 1), 'D', 12, [2])
    ts_forecast = genserie(datetime(2010, 1, 1), 'D', 20, [3])

    tsx.update('realised', ts_real, 'test')
    tsx.update('nominated', ts_nomination, 'test')
    tsx.update('forecasted', ts_forecast, 'test')

    tsx.register_formula(
        'serie5',
        '(priority (series "realised") (series "nominated") (series "forecasted"))'
    )

    values, _, origin = tsx.values_markers_origins('serie5')

    assert_df("""
2010-01-01    1.0
2010-01-02    1.0
2010-01-03    1.0
2010-01-04    1.0
2010-01-05    1.0
2010-01-06    1.0
2010-01-07    1.0
2010-01-08    1.0
2010-01-09    1.0
2010-01-10    1.0
2010-01-11    2.0
2010-01-12    2.0
2010-01-13    3.0
2010-01-14    3.0
2010-01-15    3.0
2010-01-16    3.0
2010-01-17    3.0
2010-01-18    3.0
2010-01-19    3.0
2010-01-20    3.0
""", values)

    assert_df("""
2010-01-01      realised
2010-01-02      realised
2010-01-03      realised
2010-01-04      realised
2010-01-05      realised
2010-01-06      realised
2010-01-07      realised
2010-01-08      realised
2010-01-09      realised
2010-01-10      realised
2010-01-11     nominated
2010-01-12     nominated
2010-01-13    forecasted
2010-01-14    forecasted
2010-01-15    forecasted
2010-01-16    forecasted
2010-01-17    forecasted
2010-01-18    forecasted
2010-01-19    forecasted
2010-01-20    forecasted
""", origin)

    # we remove the last value of the 2 first series which are considered as bogus

    tsx.register_formula(
        'serie6',
        '(priority '
        ' (slice (series "realised") #:todate (date "2010-1-9"))'
        ' (slice (series "nominated") #:todate (date "2010-1-11"))'
        ' (series "forecasted"))'
    )

    values, _, origin = tsx.values_markers_origins('serie6')

    assert_df("""
2010-01-01    1.0
2010-01-02    1.0
2010-01-03    1.0
2010-01-04    1.0
2010-01-05    1.0
2010-01-06    1.0
2010-01-07    1.0
2010-01-08    1.0
2010-01-09    1.0
2010-01-10    2.0
2010-01-11    2.0
2010-01-12    3.0
2010-01-13    3.0
2010-01-14    3.0
2010-01-15    3.0
2010-01-16    3.0
2010-01-17    3.0
2010-01-18    3.0
2010-01-19    3.0
2010-01-20    3.0
""", values)

    assert_df("""
2010-01-01      realised
2010-01-02      realised
2010-01-03      realised
2010-01-04      realised
2010-01-05      realised
2010-01-06      realised
2010-01-07      realised
2010-01-08      realised
2010-01-09      realised
2010-01-10     nominated
2010-01-11     nominated
2010-01-12    forecasted
2010-01-13    forecasted
2010-01-14    forecasted
2010-01-15    forecasted
2010-01-16    forecasted
2010-01-17    forecasted
2010-01-18    forecasted
2010-01-19    forecasted
2010-01-20    forecasted
""", origin)

    tsx.register_formula(
        'serie7',
        '(priority '
        ' (slice (series "realised") #:todate (date "2010-1-9"))'
        ' (slice (series "nominated") #:todate (date "2010-1-9"))'
        ' (series "forecasted"))'
    )

    values, _, origin = tsx.values_markers_origins('serie7')

    assert_df("""
2010-01-01      realised
2010-01-02      realised
2010-01-03      realised
2010-01-04      realised
2010-01-05      realised
2010-01-06      realised
2010-01-07      realised
2010-01-08      realised
2010-01-09      realised
2010-01-10    forecasted
2010-01-11    forecasted
2010-01-12    forecasted
2010-01-13    forecasted
2010-01-14    forecasted
2010-01-15    forecasted
2010-01-16    forecasted
2010-01-17    forecasted
2010-01-18    forecasted
2010-01-19    forecasted
2010-01-20    forecasted
""", origin)


def test_origin_federated(tsa1, tsa2):
    # same test as above
    # tsa1: local with remote source
    # tsa2: remote source
    ts_real = genserie(datetime(2010, 1, 1), 'D', 10, [1])
    ts_nomination = genserie(datetime(2010, 1, 1), 'D', 12, [2])
    ts_forecast = genserie(datetime(2010, 1, 1), 'D', 20, [3])

    tsa2.update('realised', ts_real, 'test')
    tsa2.update('nominated', ts_nomination, 'test')
    tsa2.update('forecasted', ts_forecast, 'test')

    tsa2.register_formula(
        'serie5',
        '(priority (series "realised") (series "nominated") (series "forecasted"))'
    )

    values, _, origin = tsa1.values_markers_origins('serie5')

    assert_df("""
2010-01-01    1.0
2010-01-02    1.0
2010-01-03    1.0
2010-01-04    1.0
2010-01-05    1.0
2010-01-06    1.0
2010-01-07    1.0
2010-01-08    1.0
2010-01-09    1.0
2010-01-10    1.0
2010-01-11    2.0
2010-01-12    2.0
2010-01-13    3.0
2010-01-14    3.0
2010-01-15    3.0
2010-01-16    3.0
2010-01-17    3.0
2010-01-18    3.0
2010-01-19    3.0
2010-01-20    3.0
""", values)

    assert_df("""
2010-01-01      realised
2010-01-02      realised
2010-01-03      realised
2010-01-04      realised
2010-01-05      realised
2010-01-06      realised
2010-01-07      realised
2010-01-08      realised
2010-01-09      realised
2010-01-10      realised
2010-01-11     nominated
2010-01-12     nominated
2010-01-13    forecasted
2010-01-14    forecasted
2010-01-15    forecasted
2010-01-16    forecasted
2010-01-17    forecasted
2010-01-18    forecasted
2010-01-19    forecasted
2010-01-20    forecasted
""", origin)

    # we remove the last value of the 2 first series which are considered as bogus

    tsa2.register_formula(
        'serie6',
        '(priority '
        ' (slice (series "realised") #:todate (date "2010-1-9"))'
        ' (slice (series "nominated") #:todate (date "2010-1-11"))'
        ' (series "forecasted"))'
    )

    values, _, origin = tsa1.values_markers_origins('serie6')

    assert_df("""
2010-01-01    1.0
2010-01-02    1.0
2010-01-03    1.0
2010-01-04    1.0
2010-01-05    1.0
2010-01-06    1.0
2010-01-07    1.0
2010-01-08    1.0
2010-01-09    1.0
2010-01-10    2.0
2010-01-11    2.0
2010-01-12    3.0
2010-01-13    3.0
2010-01-14    3.0
2010-01-15    3.0
2010-01-16    3.0
2010-01-17    3.0
2010-01-18    3.0
2010-01-19    3.0
2010-01-20    3.0
""", values)

    assert_df("""
2010-01-01      realised
2010-01-02      realised
2010-01-03      realised
2010-01-04      realised
2010-01-05      realised
2010-01-06      realised
2010-01-07      realised
2010-01-08      realised
2010-01-09      realised
2010-01-10     nominated
2010-01-11     nominated
2010-01-12    forecasted
2010-01-13    forecasted
2010-01-14    forecasted
2010-01-15    forecasted
2010-01-16    forecasted
2010-01-17    forecasted
2010-01-18    forecasted
2010-01-19    forecasted
2010-01-20    forecasted
""", origin)

    tsa2.register_formula(
        'serie7',
        '(priority '
        ' (slice (series "realised") #:todate (date "2010-1-9"))'
        ' (slice (series "nominated") #:todate (date "2010-1-9"))'
        ' (series "forecasted"))'
    )

    values, _, origin = tsa1.values_markers_origins('serie7')

    assert_df("""
2010-01-01      realised
2010-01-02      realised
2010-01-03      realised
2010-01-04      realised
2010-01-05      realised
2010-01-06      realised
2010-01-07      realised
2010-01-08      realised
2010-01-09      realised
2010-01-10    forecasted
2010-01-11    forecasted
2010-01-12    forecasted
2010-01-13    forecasted
2010-01-14    forecasted
2010-01-15    forecasted
2010-01-16    forecasted
2010-01-17    forecasted
2010-01-18    forecasted
2010-01-19    forecasted
2010-01-20    forecasted
""", origin)


def test_today_vs_revision_date(tsx):
    tsx.register_formula(
        'constant-1',
        '(constant 1. (date "2020-1-1") (today) "D" (date "2020-2-1"))'
    )

    ts, _, _ = tsx.values_markers_origins(
        'constant-1',
        revision_date=datetime(2020, 2, 1)
    )
    assert len(ts) == 32
