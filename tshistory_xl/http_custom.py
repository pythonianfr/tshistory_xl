import requests

from tshistory.config import configuration
from tshistory.http.util import get_auth
from tshistory.http.client import (
    oauth2_auth,
    pkce_auth,
)


from tshistory_xl.codecs import (
    pack_getmany_request,
    pack_insert_series,
    unpack_getmany,
)


class HTTPClient:
    _uri = None

    def __init__(self, uri=None):
        if self._uri is None and uri:
            self._uri = uri.strip()
        self.session = requests.Session()
        self.session.trust_env = False
        cfg = configuration()
        auth = get_auth(self._uri + '/api', cfg)
        if 'login' in auth:
            self.session.auth = auth['login'], auth['password']
        elif 'pkce' in auth:
            self.session.auth = pkce_auth(uri, auth)
        elif 'client_id' in auth:
            self.session.auth = oauth2_auth(auth)

    # things TimeSerie-like
    def insert_from_many(self, insertlist, author):
        bytestr = pack_insert_series(author, insertlist)
        output = self.session.patch(
            '{}/insert_from_many'.format(self._uri),
            data=bytestr,
            headers={'Content-Type': 'application/octet-stream'}
        )
        if output.status_code != 200:
            raise Exception(output.text)

    def get_many(self, inputlist):
        if not inputlist:
            return []
        data = pack_getmany_request(inputlist)
        output = self.session.post(
            '{}/get_many'.format(self._uri),
            data=data,
            headers={'Content-Type': 'application/octet-stream'}
        )
        if output.status_code != 200:
            raise Exception(output.text)
        return unpack_getmany(output.content)
