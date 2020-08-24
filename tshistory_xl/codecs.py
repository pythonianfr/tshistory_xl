import zlib
import json
import datetime

import isodate
import pandas as pd
from tshistory import util


# the excel client works under the following assumption
# value_{d}type can change in case of strings
META = {
    'tzaware': False,
    'index_type': 'datetime64[ns]',
    'index_dtype': '<M8[ns]',
    'value_type': 'float64',
    'value_dtype': '<f8'
}


def buildmeta(bvtype):
    meta = META.copy()
    if bvtype == b'object':
        meta['value_type'] = 'object'
        meta['value_dtype'] = '|O'
    elif bvtype == b'bool':
        meta['value_type'] = 'bool'
        meta['value_dtype'] = '?'
    return meta


def buildseries(bname, bindex, bvalues, bvtype):
    meta = buildmeta(bvtype)
    index, values = util.numpy_deserialize(
        bindex, bvalues, meta
    )
    return pd.Series(
        values,
        index=index,
        name=bname.decode('utf-8')
    )


def serialize_series(series):
    if series is None or not len(series):
        return (
            b'', b'', b''
        )
    bvdtype = series.dtype.name.encode('utf-8')
    bindex, bvalues = util.numpy_serialize(
        series,
        bvdtype == b'object'
    )
    return (
        bindex,
        bvalues,
        bvdtype
    )


def pack_insert_series(author, seriesgroup):
    out = [author.encode('utf-8')]
    for name, series in seriesgroup:
        out.append(name.encode('utf-8'))
        bindex, bvalues, bvtype = serialize_series(series)
        out.append(bvtype)
        out.append(bindex)
        out.append(bvalues)
    return zlib.compress(util.nary_pack(*out))


def unpack_insert_series(bytestr):
    byteslist = util.nary_unpack(zlib.decompress(bytestr))
    author = byteslist[0].decode('utf-8')
    iterbseries = zip(*[iter(byteslist[1:])] * 4)
    series = []
    for bname, bvtype, bindex, bvalues in iterbseries:
        series.append(buildseries(bname, bindex, bvalues, bvtype))
    return author, series


class MoreJson(json.JSONEncoder):

    def default(self, o):
        if isinstance(o, datetime.datetime):
            return o.isoformat()
        elif isinstance(o, datetime.timedelta):
            return isodate.duration_isoformat(o)
        return super().default(o)


def pack_getmany_request(querylist):
    """
    [('rw_test', 'bidule',
      {'delta': None,
       'from_value_date': datetime.datetime(2012, 4, 5, 0, 0),
        'name': 'bidule',
        'revision_date': None,
        'to_value_date': datetime.datetime(2012, 4, 14, 0, 0)
       }
      )]
    """
    byteslist = []
    for query in querylist:
        byteslist.append(json.dumps(query, cls=MoreJson).encode('utf-8'))
    return zlib.compress(util.nary_pack(*byteslist))


def unpack_getmany_request(querybytes):
    out = []
    for item in util.nary_unpack(zlib.decompress(querybytes)):
        item = json.loads(item)
        subquery = item[2]
        for attr in ('from_value_date', 'to_value_date', 'revision_date'):
            if subquery[attr]:
                subquery[attr] = pd.Timestamp(subquery[attr])
        if subquery['delta']:
            subquery['delta'] = isodate.parse_duration(subquery['delta'])
        out.append(item)
    return out


def pack_getmany(manyseries):
    out = []
    for zonename, seriesname, revdate, (base, origin, marker) in manyseries:
        out.append(zonename.encode('utf-8'))
        out.append(seriesname.encode('utf-8'))
        out.append(
            revdate.isoformat().encode('utf-8') if revdate else b''
        )
        for ts in (base, origin, marker):
            bindex, bvalues, bvtype = serialize_series(ts)
            out.append(bvtype)
            out.append(bindex)
            out.append(bvalues)
    return zlib.compress(util.nary_pack(*out))


def unpack_getmany(compressedbytes):
    out = []
    byteslist = util.nary_unpack(zlib.decompress(compressedbytes))
    iterbytes = zip(*[iter(byteslist)] * 12)
    for bzone, bname, bdate, btyp1, bi1, bv1, btyp2, bi2, bv2, btyp3, bi3, bv3 in iterbytes:
        base = buildseries(b'base', bi1, bv1, btyp1)
        origin = buildseries(b'origin', bi2, bv2, btyp2)
        markers = buildseries(b'markers', bi3, bv3, btyp3)
        zonename = bzone.decode('utf-8')
        name = bname.decode('utf-8')
        revdate = pd.Timestamp(bdate.decode('utf-8')) if bdate else None
        out.append(
            (zonename, name, revdate, (base, origin, markers))
        )
    return out
