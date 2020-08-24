from pathlib import Path
import io
from functools import partial

import pytest
import responses
from sqlalchemy import create_engine
import webtest
from flask import Flask

from pytest_sa_pg import db
from pymercure.utils import cachedf

from rework.schema import init as rework_init
from rework import api

from tshistory.schema import tsschema
from tshistory_formula.schema import formula_schema

from tshistory_xl import tsio
from tshistory_xl.blueprint import blueprint
from tshistory_xl.http_custom import HTTPClient

DATADIR = Path(__file__).parent / 'data'


@pytest.fixture(scope='session')
def engine(request):
    port = 5433
    db.setup_local_pg_cluster(request, DATADIR, port)
    uri = 'postgresql://localhost:{}/postgres'.format(port)
    e = create_engine(uri)
    tsschema('tsh').create(e)
    tsschema('tsh-upstream').create(e)
    formula_schema().create(e)
    yield e


@pytest.fixture(scope='session')
def tsh(engine):
    return tsio.timeseries()


class NonSuckingWebTester(webtest.TestApp):

    def _check_status(self, status, res):
        try:
            super()._check_status(self, status, res)
        except:
            pass
            # raise <- default behaviour on 4xx is silly

APP = None

def webapp(engine):
    global APP
    if APP is not None:
        return APP
    APP = Flask('test-xl')

    from tshistory.api import timeseries
    tsa = timeseries(
        str(engine.url),
        handler=tsio.timeseries
    )
    APP.register_blueprint(
        blueprint(tsa)
    )

    return APP


@pytest.fixture
def client(engine):
    return NonSuckingWebTester(webapp(engine))


dfcache = cachedf(DATADIR)


# HTTP tests
# Error-displaying web tester

class WebTester(NonSuckingWebTester):

    def _gen_request(self, method, url, params,
                     headers=None,
                     extra_environ=None,
                     status=None,
                     upload_files=None,
                     expect_errors=False,
                     content_type=None):
        """
        Do a generic request.
        PATCH: *bypass* all transformation as params comes
               straight from a prepared (python-requests) request.
        """
        environ = self._make_environ(extra_environ)

        environ['REQUEST_METHOD'] = str(method)
        url = str(url)
        url = self._remove_fragment(url)
        req = self.RequestClass.blank(url, environ)

        req.environ['wsgi.input'] = io.BytesIO(params)
        req.content_length = len(params)
        if headers:
            req.headers.update(headers)
        return self.do_request(req, status=status,
                               expect_errors=expect_errors)

# test uri
HTTPClient._uri = 'http://test-uri'

# responses wrapper for requests
def transmit_things_bridge(method):
    def bridge(request):
        resp = method(request.url,
                      params=request.body,
                      headers=request.headers)
        return (resp.status_code, resp.headers, resp.body)
    return bridge


def get_things_bridge(client, request):
    resp = client.get(request.url,
                      params=request.body,
                      headers=request.headers)
    return (resp.status_code, resp.headers, resp.body)


@pytest.fixture(scope='session')
def excel(engine):
    client = WebTester(webapp(engine))
    with responses.RequestsMock(assert_all_requests_are_fired=False) as resp:

        resp.add_callback(
            responses.GET, 'http://test-uri/api/series/state',
            callback=partial(get_things_bridge, client)
        )

        resp.add_callback(
            responses.GET, 'http://test-uri/api/series/metadata',
            callback=partial(get_things_bridge, client)
        )

        resp.add_callback(
            responses.PATCH, 'http://test-uri/insert_from_many',
            callback=transmit_things_bridge(client.patch)
        )

        resp.add_callback(
            responses.POST, 'http://test-uri/get_many',
            callback=transmit_things_bridge(client.post)
        )

        yield
