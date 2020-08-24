from datetime import datetime, timedelta
from contextlib import contextmanager
from collections import Counter
import os
import calendar as cal
import itertools
from functools import partial
import time
import platform
import string
import gc

import pandas as pd
import numpy as np
import colorlover as colorlover
from dateutil import parser

from tshistory_xl.http_custom import HTTPClient

PLATFORM = platform.system()
if PLATFORM in ('Darwin', 'Windows'):
    import xlwings as xw
else:
    xw = None

def col2num(col):
    num = 0
    for c in col:
        if c in string.ascii_letters:
            num = num * 26 + (ord(c.upper()) - ord('A')) + 1
    return num


# elements of excel names whose presence
# denote permissions/operations
PUSH = 'w'  # allowed to push
PULL = 'r'  # allowed to read
COLOR = 'c'  # for colored ouput
FILL_TRAIL = 'f'   # for filling the last values with the previous
MONTH = '_month_'  # tagged values must be unfolded daily for a whole month
IGNORE = 'ignore'  # whole name, column will be ignored (equivalent to no name)

CFG_SHEET = '_SATURN_CFG'
MODE = 'mode'
READ_ONLY = 'ro'
NA = 'na'
ZERO = '0'
PREV = 'prev'
BLANK = 'blank'
EMPTY = 'empty'
ASOFDELTA = 'asofdelta='
ASOF = 'asof'
AGG = 'agg'
NA_XL = '#N/A'

BOGUS_CHARACTER = ('(', ')', '-', 'é', 'à', 'è', '&', '@', '+')
# can be present in the sheet name but not in the excel 'names'

AUTHOR = os.environ.get('USERNAME', 'TOX_USER')

COLOR_MARKER = (251, 128, 114)
COLOR_PULL = (219, 229, 241)
COLOR_READ_ONLY = (255, 255, 179)
COLOR_PUSH = (255, 170, 170)

# aggregation option
SUM = 'sum'
MAX = 'max'
MIN = 'min'
MEAN = 'mean'

#for test compatibility
DIE_ON_ERROR = True


_OPS = set([PUSH, PULL, COLOR, FILL_TRAIL])
def operations(name):
    """ extract operations from an excel name """
    if '_' not in name:
        return set()
    ops, _ = name.split('_', 1)
    ops = set(ops)
    if ops - _OPS:
        return set()
    return ops


def hasoperation(name, op):
    return op in operations(name)


def push_ts(wb, pushlist):
    tsh = HTTPClient(get_webapi_uri(wb))
    tsh.insert_from_many(pushlist, AUTHOR)


def parse_revision_date(sheet, top, left):
    range_rev = (top - 1, left - 1)
    range_rev_A1 = sheet.range(range_rev).address
    try:
        sheet.api.Range(range_rev_A1).Comment.Delete()
    except:
        pass
    if PLATFORM != 'Darwin':
        sheet.api.Range(range_rev_A1).AddComment('ASOF')

    mayberevdate = sheet.range(range_rev).value
    if isinstance(mayberevdate, str):
        if ASOFDELTA in mayberevdate:
            delta = float(mayberevdate.strip().split('=')[1])
            return None, delta
    elif isinstance(mayberevdate, datetime):
        if PLATFORM != 'Darwin':
            sheet.range(range_rev).color = COLOR_READ_ONLY
            sheet.range(range_rev).api.Font.Bold = True
            sheet.range(range_rev).api.Borders.LineStyle = 1
            sheet.range(range_rev).api.Borders.Weight = 3
            return mayberevdate, None
    else:
        if PLATFORM != 'Darwin':
            sheet.range(range_rev).color = None
            sheet.range(range_rev).api.Font.Bold = False
            sheet.range(range_rev).api.Borders.LineStyle = 1
            sheet.range(range_rev).api.Borders.Weight = 2
    return None, None


def find_zones_having_keyword(
        wb, parts, sheetname=None,
        # let's be slightly higher order to accomodate for different uses
        predicate=hasoperation):
    """ Scan the workbook (possibly only the given sheet name) for relevant
    zones (i.e groups of series having a common timestamps column).
    The keywords are operation (PUSH, PULL) and other weird things.

    The zone names obey certain rules (see end user documention).
    We map the zone names to a set of properties, amongst which:
    * the series names
    * zone coordinates
    * some fancy metadata
    """
    zones = {}
    for xlname in wb.names:
        present = True
        for word in parts:
            present = present and predicate(xlname.name, word)
        if not present:
            continue
        try:
            sheet_name = xlname.refers_to_range.sheet.name
        except:
            sheet_name_str = xlname.refers_to
            if sheet_name_str.startswith('=#REF'):
                raise Exception(sheet_name_str + ' has no proper cell reference. '
                                'Please check your Name Manager')

        if sheetname and sheet_name != sheetname:
            continue


        range_xls = xlname.refers_to_range.address
        if ':' in range_xls:
            corner_0, corner_1 = range_xls.split(':')
        else:  # the selection contains only one cell
            corner_0 = range_xls
            corner_1 = range_xls

        x_range = col2num(corner_0.split('$')[1]), col2num(corner_1.split('$')[1])
        y_range = int(corner_0.split('$')[2]), int(corner_1.split('$')[2])

        left = min(x_range)
        right = max(x_range)
        top = min(y_range)
        down = max(y_range)

        zones[xlname.name] = zone = {
            'sheetname': sheet_name,
            'range': range_xls,
            'left': left,
            'top': top,
            'right': right,
            'down': down
        }
        sheet = wb.sheets[sheet_name]
        list_head = sheet.range((top - 1, left), (top - 1, right)).options(ndim=1).value
        rev_date, delta = parse_revision_date(sheet, top, left)
        col_names, map_read_only, map_na, map_aggregation = parse_header(
            list_head,
            rev_date
        )
        if rev_date or delta:
            map_read_only = {
                k: True
                for k, _ in map_read_only.items()
            }
        forbid_duplicates(col_names)

        coord = bounding_box(zone)
        timestamps = sheet.range((coord['top'], coord['left'] - 1),
                                 (coord['down'], coord['left'] - 1)).options(ndim=1).value
        zone['from_date'] = min(filter(None, timestamps))
        zone['to_date'] = max(filter(None, timestamps))


        # NOTE: some indications wanted
        zone['col_name'] = col_names                      # serie name
        zone['map_read_only'] = map_read_only             # r/o flag
        zone['map_na'] = map_na                           # N/A flag
        zone['map_aggregation'] = map_aggregation         # type of aggregation to apply
        zone['monthly'] = MONTH in xlname.name            # monthly interpolation flag
        zone['fill_trail'] = hasoperation(xlname.name,
                                          FILL_TRAIL)     # trail interpolation flag
        zone['color'] = hasoperation(xlname.name, COLOR)  # colorization flag (ui control)
        zone['revision_date'] = rev_date
        zone['delta'] = delta
    return zones


find_zone_by_name = partial(find_zones_having_keyword,
                            predicate=lambda name, part: part == name)


def parse_header(list_head, rev_date):
    col_names = []
    map_read_only = {}
    map_na = {}
    map_aggregation = {}

    for head in list_head:
        if head is None:
            col_names.append((None, None))
            map_read_only[None, None] = None
            map_na[None, None] = None
            continue

        if isinstance(head, float):
            head = str(int(head))
        read_only_stuff = False
        na_stuff = None
        asof_date = rev_date
        aggregation = None
        head, options = parse_option(head)
        if options:
            if ASOF in options:
                asof_date = parser.parse(options[ASOF])
                read_only_stuff = True
            if MODE in options:
                read_only_stuff = options[MODE] == READ_ONLY
            if BLANK in options:
                na_stuff = options[BLANK]
                if na_stuff in (ZERO, PREV):
                    read_only_stuff = True
            if AGG in options:
                aggregation = options[AGG]
                read_only_stuff = True
        head = head.replace(' ', '')
        col_names.append((head, asof_date))

        map_na[(head, asof_date)] = na_stuff
        map_read_only[(head, asof_date)] = read_only_stuff
        map_aggregation[(head, asof_date)] = aggregation

    return col_names, map_read_only, map_na, map_aggregation


def parse_option(head_col):
    if '(' not in head_col or ')' not in head_col:
        return head_col, None
    begin, temp = head_col.split('(')
    inside, end = temp.split(')')
    inside = inside.replace(' ', '')
    list_inside = inside.split(',')
    list_inside = [elt.split('=') for elt in list_inside]
    map_inside = {elt[0]: elt[1] for elt in list_inside}
    return begin + end, map_inside


def int_to_str(seq):
    return [str(int(elt)) if isinstance(elt, float) else elt
            for elt in seq]


def bounding_box(zonedata):
    return {
        attr: zonedata[attr]
        for attr in ('top', 'down', 'left', 'right')
    }


def push_parse_ts(wb, zones):
    """
    Use the zone map to parse the full data (values and possibly metadata)
    """
    output = {}
    for zonename, zonedata in zones.items():
        sheet = wb.sheets[zonedata['sheetname']]
        coord = bounding_box(zonedata)
        timestamps = sheet.range(
            (coord['top'], coord['left'] - 1),
            (coord['down'], coord['left'] - 1)
        ).value

        serie_names = zonedata['col_name']
        data = sheet.range(
            (coord['top'], coord['left']),
            (coord['down'], coord['right'])
        ).value

        if isinstance(data, list):
            data_array = np.array(data)
        else:
            data_array = np.array([data])
        # let's put everythin into a big dataframe
        data_extract = pd.DataFrame(
            data_array,
            index=timestamps,
            columns=serie_names
        )
        output[zonename] = {}
        valued_columns = [
            (idx, elt)
            for idx, elt in enumerate(data_extract.columns)
            if elt != (None, None) and IGNORE not in elt
        ]
        for idx_col, colname in valued_columns:
            output[zonename][colname] = {}
            output[zonename][colname]['value'] = data_extract.iloc[:, idx_col]

    return output


def pull_prepare_zones(wb, zones):
    """
    Use the zone map to parse the timestamps and cleanup the metadata
    """
    zones_tstamps = {}
    for zonename, zonedata in zones.items():
        sheet = wb.sheets[zonedata['sheetname']]
        coord = bounding_box(zonedata)
        timestamps = sheet.range(
            (coord['top'], coord['left'] - 1),
            (coord['down'], coord['left'] - 1)
        ).options(ndim=1).value
        # We DONT wipe the holes as they are *significant* for the write back
        zones_tstamps[zonename] = timestamps

    return zones_tstamps


def extrapolate_daily(values):
    mask = values.index.astype('int') != int(pd.NaT)
    values = values[mask]
    first_value = values[values.index == min(values.index)]
    last_value = values[values.index == max(values.index)]
    first_of_month = datetime(
        first_value.index.year[0],
        first_value.index.month[0],
        1
    )
    last_day = cal.monthrange(
        last_value.index[0].year,
        last_value.index[0].month
    )[1]
    last_of_month = datetime(
        last_value.index.year[0],
        last_value.index.month[0],
        last_day
    )
    extrapole = pd.Series(
        index=pd.date_range(
            start=first_of_month,
            end=last_of_month, freq='D'
        )
    )
    for cell in values.iteritems():
        extrapole[str(cell[0].year) + '-' + str(cell[0].month)] = cell[1]
    return extrapole


def push_rolling(wb, zones):
    zones_series = push_parse_ts(wb, zones)
    pushlist = []
    for zonename, series in zones_series.items():
        zonedata = zones[zonename]
        for ts_name, dico_rich in series.items():
            read_only = zonedata['map_read_only'][ts_name]
            if read_only:
                continue

            values = dico_rich['value']

            # removing the NA, parsed as -2146826246
            values[values == -2146826246] = None
            # in case it appears as string (can't reproduce it in the test file,
            # but it has already happened in the wild...)
            if values.dtype == 'O':
                values[values == '-2146826246'] = None
                values[values == NA_XL] = None      # It happens. Sometime....

            # nasty people put strings in the middle of their data
            # turns float64 into object ...
            # hence we filter the bad values and reconstruct the serie
            name = values.name
            goodvalues = values.index.astype('int') != int(pd.NaT)
            values = pd.Series(values[goodvalues].to_dict())
            values.name = name

            # coerce value to numeric if possible
            try:
                values = pd.to_numeric(values, errors='raise')
            except:
                values = values.astype(str)

            if zonedata['monthly']:
                values = extrapolate_daily(values)

            pushlist.append((ts_name[0], values))

        if zonedata['color']:
            wb.sheets[zonedata['sheetname']].range(zonename).color = COLOR_PUSH
            for idx, colname in enumerate(zonedata['col_name']):
                if colname is None:
                    continue
                col_ro = zonedata['map_read_only'][colname]
                if col_ro:
                    wb.sheets[zonedata['sheetname']].range(
                        (zonedata['top'], zonedata['left'] + idx),
                        (zonedata['down'], zonedata['left'] + idx)
                    ).color = COLOR_READ_ONLY
        else:
            paint_header(wb, zonedata, COLOR_PUSH)

    push_ts(wb, pushlist)


def paint_header(wb, coord, color):
    for idx, col_head in enumerate(coord['col_name']):
        if col_head is not None and IGNORE not in col_head:
            cell = coord['top'] - 1, coord['left'] + idx
            wb.sheets[coord['sheetname']].range(cell).color = color


def find_bounds(axis):
    search_first_coord = True
    bounds = []
    for idx in range(len(axis)):
        if search_first_coord:
            if axis[idx] is not None:
                first_coord = idx
                search_first_coord = False
        if not search_first_coord:
            if axis[idx] is None or idx == len(axis) - 1:
                last_coord = idx - 1 if axis[idx] is None else idx
                bounds.append(
                    (first_coord, last_coord)
                )
                search_first_coord = True
    return bounds


def print_data(wb, coord, data_df):
    bounds_col = find_bounds([name for name, _ in coord['col_name']])
    axis_row = data_df.index.astype('str').values
    # We replace the pd.Nat by None
    axis_row[axis_row == 'NaT'] = None
    bounds_row = find_bounds(axis_row)

    for prod in itertools.product(bounds_row, bounds_col):
        df = data_df.iloc[prod[0][0]:prod[0][1] + 1, prod[1][0]:prod[1][1] + 1]
        # NA handeling

        map_na = {
            str(k[0]) + str(k[1]): v
            for k, v in coord['map_na'].items()
        }
        vec_na = np.array(
            [map_na[elt]
             for elt in df.columns]
        )

        if not coord['fill_trail']:
            for i_col in range(len(df.columns)):
                filler = vec_na[i_col]
                if df.iloc[:, i_col].isnull().any():
                    if filler in [None, NA]:
                        df.iloc[:, i_col] = df.iloc[:, i_col].fillna(NA_XL)
                    elif filler in [ZERO]:
                        df.iloc[:, i_col] = df.iloc[:, i_col].fillna(ZERO)
                    elif filler in [PREV]:
                        df.iloc[:, i_col] = df.iloc[:, i_col].fillna(method='ffill')
                    elif filler in [EMPTY]:
                        pass
                    else:
                        raise Exception(f"""{filler} not autorised for (blank= )""")

        elif (coord['fill_trail'] and prod[0][1] == bounds_row[-1][1]
              and not coord['revision_date']):
            # ie. the bottom of the area is the bottom of the zone
            for i_col in range(len(df.columns)):
                ts = df.iloc[:, i_col]
                if ts.isnull().all() or ts.dtype == 'O':
                    continue
                last_valid = ts.index[ts.notnull()][-1]
                df.loc[last_valid:, :].iloc[:, i_col] = ts[last_valid:].fillna(method='ffill')
        wb.sheets[coord['sheetname']].range(
            (coord['top'] + prod[0][0], coord['left'] + prod[1][0])).value = df.values


def forbid_duplicates(somelist, seqname='columns', itemname='name'):
    """ check uniqueness of elements of some sequence
    This will raise an exception """
    reduced_list = [elt for elt in somelist if elt != (None, None)]
    dupes = [
        x
        for x, y in Counter(reduced_list).items()
        if y > 1
    ]
    if dupes:
        strdupes = [name for name, _  in dupes]
        raise Exception(
            'Cannot process {} with identical {}s ({})'.format(
                seqname,
                itemname,
                ', '.join(strdupes))
        )


def pull_series(wb, zones):
    inputlist = []
    for zonename, zonedata in zones.items():
        for seriename, asofdate in zonedata['col_name']:
            if seriename is None:
                continue
            delta_as_td = timedelta(hours=zonedata['delta']) if zonedata['delta'] else None
            inputlist.append(
                (zonename, seriename, {
                    'name': seriename,
                    'revision_date': asofdate,
                    'delta': delta_as_td,
                    'from_value_date': zonedata['from_date'],
                    'to_value_date': zonedata['to_date']
                })
            )
    tsh = HTTPClient(get_webapi_uri(wb))
    start = time.time()
    outputlist = tsh.get_many(inputlist)
    print('get many client done in %s seconds' % (time.time() - start))

    # transform into easier to use mapping
    out = {}
    for zonename, seriename, asofdate, data in outputlist:
        out[(zonename, seriename, asofdate)] = data

    return out


def isnat(dtindex):
    return dtindex.astype('int') == int(pd.NaT)


def build_zones_dataframes(wb, zones, db_series):
    zones_timestamps = pull_prepare_zones(wb, zones)
    zones_dfs = {}

    for zonename in zones:
        timestamps = zones_timestamps[zonename]
        data_df = pd.DataFrame(index=timestamps)
        marker_df = pd.DataFrame(index=timestamps)
        origin_df = pd.DataFrame(index=timestamps)

        # actually the spec/metadata of the zone
        zonedata = zones[zonename]

        idx_ignore = 1
        for colname, asofdate in zonedata['col_name']:
            if colname is None:
                # we still write shit into the big dataframes
                data_df[IGNORE + str(idx_ignore)] = np.nan
                marker_df[IGNORE + str(idx_ignore)] = False
                origin_df[IGNORE + str(idx_ignore)] = None
                idx_ignore += 1
                continue

            #aggregation
            agg_type = zonedata['map_aggregation'][colname, asofdate]

            ts_base, marker, origin = db_series[(zonename, colname, asofdate)]
            new_name = colname + str(asofdate)
            # so that the name of the col data_df match the col_name in zone
            if ts_base is not None:
                if agg_type:
                    ts_base = aggregate_series(ts_base, timestamps, agg_type)
                ts_base.name = new_name
                data_df = data_df.join(ts_base, how='outer')
            else:
                data_df[new_name] = None
            if marker is not None:
                marker.name = new_name
                marker_df = marker_df.join(marker, how='outer')
            else:
                marker_df[new_name] = False
            if origin is not None:
                origin.name = new_name
                origin_df = origin_df.join(origin, how='outer')
            else:
                origin_df[new_name] = None

        try:
            data_df = data_df[~isnat(data_df.index)]
        except TypeError:
            raise Exception('In zone %s, sheet %s the index is bogus %s' % (
                zonename,
                zones[zonename]['sheetname'],
                [str(elt) for elt in data_df.index.values])
            )

        marker_df = marker_df[~isnat(marker_df.index)]
        origin_df = origin_df[~isnat(origin_df.index)]

        data_df.loc[pd.NaT] = np.nan
        data_df = data_df.loc[timestamps]
        marker_df.fillna(False, inplace=True)
        marker_df.loc[pd.NaT] = False
        marker_df = marker_df.reindex(timestamps)
        origin_df = origin_df.reindex(timestamps)
        origin_df = origin_df.where((pd.notnull(origin_df)), None)

        zones_dfs[zonename] = {
            'data': data_df,
            'marker': marker_df,
            'origin': origin_df,
        }
    return zones_dfs


def aggregate_series(ts_base, timstamps, agg_type):
    base_name = ts_base.name
    timstamps = [elt for elt in timstamps if not pd.isnull(elt)]
    pseudo_period = pd.Series(range(len(timstamps)), index=timstamps, name='marker')
    df = ts_base.to_frame().join(pseudo_period, how='outer')
    df['marker'] = df['marker'].fillna(method='ffill')

    if agg_type == SUM:
        # trick so that the NAs are not aggregated into 0
        df_result = df.groupby('marker').apply(
            lambda x: x.sum(skipna=False)
        )
    elif agg_type == MAX:
        df_result = df.groupby('marker').max()
    elif agg_type == MIN:
        df_result = df.groupby('marker').min()
    elif agg_type == MEAN:
        df_result = df.groupby('marker').mean()
    else:
        return ts_base
    invert = pseudo_period.reset_index().set_index('marker')
    ts = df_result.join(invert).set_index('index')[base_name]
    return ts[:-1] # the last value has no proper meaning


def pull_rolling(wb, zones):
    db_series = pull_series(wb, zones)
    update_workbook(wb, zones, db_series)


def compress_series(ts):
    # GIF style
    shifted = ts.shift(1)
    bounds = ts[1:] != shifted[1:]
    pos_index = np.where(bounds)[0]
    bounds = bounds[bounds]  # we only keep the True values
    if not len(bounds):
        return [(ts.iloc[0], (0, len(ts) - 1))]
    result_list = []
    start_pos_index = 0
    for pos in pos_index:
        if ts.iloc[pos] is not None:
            result_list.append((ts.iloc[pos], (start_pos_index, pos)))
        start_pos_index = pos + 1
    if ts.iloc[start_pos_index] is not None:
        result_list.append((ts.iloc[start_pos_index], (start_pos_index, len(ts) - 1)))
    return result_list


def update_workbook(wb, zones, db_series):
    zones_dfs = build_zones_dataframes(wb, zones, db_series)

    for zonename, zone_df in zones_dfs.items():
        data_df = zone_df['data']
        marker_df = zone_df['marker']
        origin_df = zone_df['origin']

        zonedata = zones[zonename]
        if not hasoperation(zonename, PUSH):
            color_pull = COLOR_READ_ONLY
        else:
            color_pull = COLOR_PULL

        bool_color = zonedata['color']
        if bool_color:
            wb.sheets[zonedata['sheetname']].range(zonename).color = color_pull
        else:
            paint_header(wb, zonedata, color_pull)
        print_data(wb, zonedata, data_df)
        # write the header
        list_name = [name for name, _ in zonedata['col_name']]
        list_label = wb.sheets[zonedata['sheetname']].range(
            (int(zonedata['top']) - 1, int(zonedata['left'])),
            (int(zonedata['top']) - 1, int(zonedata['right']))
        ).options(ndim=1).value
        list_label = [str(label) for label in list_label]
        decorate_header(
            wb,
            zonedata['sheetname'],
            (zonedata['top'] - 1, zonedata['left']),
            list_name,
            list_label,
            zonedata['from_date'],
            zonedata['to_date']
        )

        sheet = wb.sheets[zonedata['sheetname']]
        # origin
        for idx_col in range(origin_df.shape[1]):
            origines = origin_df.iloc[:, idx_col]
            col_marker = marker_df.iloc[:, idx_col]
            if len(np.unique(origines[(~origines.isnull())])) < 2:
                if not col_marker.any():
                    continue
                origines[~col_marker] = None
            origines[col_marker] = 'marker'
            cont_coord = compress_series(origines)
            origines = origines[~origines.isnull()]
            nb_color = len(np.unique(origines))
            couleurs = colorlover.scales['7']['qual']['Pastel2'][:nb_color]
            teintes = colorlover.to_numeric(couleurs)
            map_col = dict(zip(np.unique(origines), teintes))
            map_col['marker'] = COLOR_MARKER
            for elt in cont_coord:
                origine = elt[0]
                start = elt[1][0]
                end = elt[1][1]
                if bool_color:
                    abs_col = int(zonedata['left'] + idx_col)
                    sheet.range(
                        (int(zonedata['top'] + start), abs_col),
                        (int(zonedata['top'] + end), abs_col)
                    ).color = map_col[origine]


def d2dt(date, end_of_day=False):
    if not end_of_day:
        return datetime(
            year=date.year,
            month=date.month,
            day=date.day
        )
    return datetime(
        year=date.year,
        month=date.month,
        day=date.day,
        hour=23,
        minute=59
    )


def build_origin(ts):
    # helper to construct an origine series for Mercure data
    if ts is None or len(ts) == 0:
        return None
    return  pd.Series(
        [ts.name] * len(ts),
        index=ts.index,
        name=ts.name
    )


def decorate_header(wb, sheetname, corner_head, list_name, list_label, fromdate, todate):
    uri = get_webapi_uri(wb)
    for idx, (name, label) in enumerate(zip(list_name, list_label)):
        if name is None:
            continue
        position = (corner_head[0], corner_head[1] + idx)
        hypertext = "{}/tseditor/?name={}&startdate={}&enddate={}&author={}".format(
            uri,
            name,
            fromdate.date(),
            todate.date() + timedelta(days=1),
            AUTHOR
        )
        # excel has its own set of rules...
        if len(hypertext) < 196:
            wb.sheets[sheetname].range(position).add_hyperlink(
                hypertext,
                text_to_display=label
            )


@contextmanager
def xlbook(xlpath=None, close=False):
    debug = os.environ.get('XL_DEBUG')
    if xlpath:
        wb = xw.Book(str(xlpath))
    else:
        wb = xw.Book.caller()
    updating = wb.app.screen_updating
    computemode = wb.app.calculation
    if not debug:
        wb.app.screen_updating = False
        wb.app.calculation = 'manual'

    try:
        yield wb
    finally:
        wb.app.screen_updating = updating
        wb.app.calculation = computemode
        if close and not debug:
            wb.close()
        gc.collect(2)


_TESTS = False
def get_webapi_uri(wb):
    if _TESTS:
        return None
    try:
        df_cfg = pd.DataFrame(wb.sheets[CFG_SHEET].range((1, 1), (10, 10)).value)
    except:
        raise Exception('No {} sheet in {}'.format(CFG_SHEET, wb.name))
    if 'webapi' not in df_cfg.iloc[:, 0].tolist():
        raise Exception(f'You must specified in {CFG_SHEET}: webapi | www.saturnurl.net')
    coord_uri_api = (
        np.where(df_cfg == 'webapi')[0][0],
        np.where(df_cfg == 'webapi')[1][0] + 1
    )
    raw_uri = df_cfg.iloc[coord_uri_api[0], coord_uri_api[1]]
    return raw_uri.strip('/')


def macro_pull_all(xlpath=None):
    with xlbook(xlpath) as wb:
        zones = find_zones_having_keyword(wb, [PULL])
        pull_rolling(wb, zones)


def macro_pull_tab(xlpath=None, tab=None):
    with xlbook(xlpath) as wb:
        if not tab:
            tab = wb.sheets.active.name
        zones = find_zones_having_keyword(wb, [PULL], tab)
        pull_rolling(wb, zones)


def macro_push_all(xlpath=None):
    with xlbook(xlpath) as wb:
        zones = find_zones_having_keyword(wb, [PUSH])
        push_rolling(wb, zones)


def macro_push_tab(xlpath=None, tab=None):
    with xlbook(xlpath) as wb:
        if not tab:
            tab = wb.sheets.active.name
        zones = find_zones_having_keyword(wb, [PUSH], tab)
        push_rolling(wb, zones)
