import json
import pandas as pd

import requests
import isodate
from flask import make_response

from flask_restx import (
    Resource,
    reqparse
)

from tshistory import util
from tshistory.http.client import (
    Client,
    strft,
    unwraperror
)
from tshistory.http.server import httpapi
from tshistory.http.util import (
    enum,
    onerror,
    utcdt
)

base = reqparse.RequestParser()
base.add_argument(
    'name',
    type=str,
    required=True,
    help='timeseries name'
)

xl = base.copy()
xl.add_argument(
    'revision_date', type=utcdt, default=None,
    help='select a specific version'
)
xl.add_argument(
    'from_value_date', type=utcdt, default=None,
    help='left boundary'
)
xl.add_argument(
    'to_value_date', type=utcdt, default=None,
    help='right boundary'
)
xl.add_argument(
    'delta', type=str, default=None,
    help='optional time delta'
)


class xl_httpapi(httpapi):

    def routes(self):
        super().routes()

        tsa = self.tsa
        api = self.api
        nss = self.nss
        nsg = self.nsg

        @nss.route('/xl')
        class series_xl(Resource):

            @api.expect(xl)
            @onerror
            def get(self):
                args = xl.parse_args()
                if not tsa.exists(args.name):
                    api.abort(404, f'`{args.name}` does not exists')

                v, m, o = tsa.values_markers_origins(
                    args.name,
                    args.revision_date,
                    args.from_value_date,
                    args.to_value_date,
                    args.delta
                )

                if v is not None:
                    v = v.to_json(orient='index', date_format='iso')
                if m is not None:
                    m = m.to_json(orient='index', date_format='iso')
                if o is not None:
                    o = o.to_json(orient='index', date_format='iso')

                resp = make_response(
                    json.dumps((v, m, o))
                )
                resp.headers['Content-Type'] = 'text/json'
                resp.status_code = 200

                return resp


class XLClient(Client):

    def __repr__(self):
        return f"tshistory-xl-http-client(uri='{self.uri}')"

    @unwraperror
    def values_markers_origins(self,
                               name,
                               revision_date=None,
                               from_value_date=None,
                               to_value_date=None,
                               delta=None):
        args = {'name': name}
        if revision_date:
            args['revision_date'] = strft(revision_date)
        if from_value_date:
            args['from_value_date'] = strft(from_value_date)
        if to_value_date:
            args['to_value_date'] = strft(to_value_date)
        if delta:
            args['delta'] = isodate(delta)

        res = requests.get(
            f'{self.uri}/series/xl', params=args
        )
        if res.status_code == 404:
            return None, None, None

        if res.status_code == 200:
            return [
                util.fromjson(item, name, tzaware=self.metadata(name, all=True)['tzaware'])
                if item else item
                for item in res.json()
            ]

        return res
