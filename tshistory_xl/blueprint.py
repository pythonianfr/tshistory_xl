import time
import traceback as tb
from contextlib import contextmanager

from flask import (
    Blueprint,
    make_response,
    request,
)

from sqlalchemy import create_engine
from tshistory.util import threadpool

from tshistory_xl.codecs import (
    pack_getmany,
    unpack_getmany_request,
    unpack_insert_series,
)


NTHREAD = 16


@contextmanager
def yield_engine(dburi):
    engine = create_engine(dburi, pool_size=NTHREAD)
    yield engine
    engine.dispose()


def blueprint(tsa):
    bp = Blueprint('xlapi', __name__)
    dburi = str(tsa.engine.url)

    @bp.route('/insert_from_many', methods=['PATCH'])
    def insert_from_many():
        author, allseries = unpack_insert_series(request.get_data())
        errors = []

        poolrun = threadpool(NTHREAD)
        def insert(series, author):
            try:
                tsa.update(
                    series.name,
                    series,
                    author=author,
                    manual=True
                )
            except Exception as err:
                errors.append((series.name, err))

        poolrun(
            insert,
            [(series, author) for series in allseries]
        )

        if not len(errors):
            return make_response('got it', 200)
        return make_response(str(errors), 500)


    @bp.route('/get_many', methods=['POST'])
    def get_many():
        start = time.time()
        poolrun = threadpool(NTHREAD)

        inputlist = unpack_getmany_request(request.get_data())
        output = []
        errors = []

        def get(zonename, seriename, items):
            try:
                base, origin, marker = tsa.values_markers_origins(**items)
                output.append(
                    (zonename, seriename, items['revision_date'],
                     (base, origin, marker))
                )
            except Exception:
                errors.append((items['name'], tb.format_exc()))

        poolrun(
            get,
            [(zonename, seriename, items)
             for zonename, seriename, items in inputlist]
        )

        print('get_many done in %s seconds' % (time.time() - start))
        if not len(errors):
            response = make_response(
                pack_getmany(output),
                200
            )
            response.headers['Content-Type'] = 'application/octet-stream'
            return response
        return make_response(str(errors), 500)


    return bp
