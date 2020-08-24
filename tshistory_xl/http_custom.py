import requests

from tshistory_xl.codecs import (
    pack_getmany_request,
    pack_insert_series,
    unpack_getmany,
)


class HTTPClient:
    _uri = None

    def __init__(self, uri=None):
        if self._uri is None:
            self._uri = uri
        self.session = requests.Session()
        self.session.trust_env = False

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
