import json

import isodate
from flask import make_response

from flask_restx import (
    Resource,
    reqparse
)

from tshistory import util
from tshistory.http.client import (
    strft,
    unwraperror
)
from tshistory.http.util import (
    onerror,
    required_roles,
    utcdt
)

from tshistory_supervision.http import (
    supervision_httpapi,
    supervision_httpclient
)

from tshistory_formula.http import (
    formula_httpapi,
    formula_httpclient
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


class xl_httpapi(supervision_httpapi, formula_httpapi):
    __slots__ = 'tsa', 'bp', 'api', 'nss', 'nsg'

    def routes(self):
        super().routes()

        tsa = self.tsa
        api = self.api
        nss = self.nss

        @nss.route('/xl')
        class series_xl(Resource):

            @api.doc(
                responses={200: 'Got content',
                           404: 'Does not exist'},
                description="""Get series with values, markers, and origins

Excel-optimized endpoint that returns all data in a single request: series values, manual edit markers (from supervision), and value origins (for priority formulas).

**Parameters:**
- name: series name
- revision_date: get version at this timestamp (ISO8601, optional, default: latest)
- from_value_date: restrict time range start (ISO8601, optional)
- to_value_date: restrict time range end (ISO8601, optional)
- delta: time delta for relative range (ISO8601 duration, optional)

**Returns:** JSON array with three elements [values, markers, origins]
```json
[
  {"2025-01-01T00:00:00+00:00": 100.5, "2025-01-02T00:00:00+00:00": 105.2},
  {"2025-01-01T00:00:00+00:00": false, "2025-01-02T00:00:00+00:00": true},
  {"2025-01-01T00:00:00+00:00": "source-a", "2025-01-02T00:00:00+00:00": "source-b"}
]
```

Each element can be null if not available.

**Example:**
```
GET /series/xl?name=temperature&from_value_date=2025-01-01
→ [{"2025-01-01T00:00:00+00:00": 12.5, ...}, null, null]
```

**Note:** Origins are only computed for priority formulas.
"""
            )
            @api.expect(xl)
            @onerror
            @required_roles('admin', 'rw', 'ro')
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


class xl_httpclient(supervision_httpclient, formula_httpclient):
    index = 2

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

        res = self.session.get(
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
