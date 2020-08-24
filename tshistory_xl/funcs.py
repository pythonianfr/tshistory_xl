import pandas as pd

from tshistory.util import patch
from tshistory_formula.registry import func


@func('priority-origin')
def series_priority_origin(*serieslist: pd.Series) -> pd.Series:
    # NO DOC: do not expose it
    # this is a hack to provide the origin of a priority
    # as a second series
    final = serieslist[-1]
    origin = pd.Series(
        [final.name] * len(final.index),
        index=final.index
    )

    for ts in reversed(serieslist[:-1]):
        assert ts.dtype != 'O'
        prune = ts.options.get('prune')
        if prune:
            ts = ts[:-prune]
        final = patch(final, ts)

        # origin
        ids = pd.Series(
            [ts.name] * len(ts.index),
            index=ts.index
        )
        origin = patch(origin, ids)

    return final, origin
