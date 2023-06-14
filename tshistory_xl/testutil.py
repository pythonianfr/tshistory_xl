from functools import partial

import responses

from tshistory.testutil import read_request_bridge

from tshistory_supervision.testutil import with_http_bridge as supervisionbridge
from tshistory_formula.testutil import with_http_bridge as formulabridge


class with_http_bridge(supervisionbridge, formulabridge):

    def __init__(self, uri, resp, wsgitester):
        super().__init__(uri, resp, wsgitester)

        resp.add_callback(
            responses.GET, uri + '/series/xl',
            callback=partial(read_request_bridge, wsgitester)
        )
