from tshistory.util import ensuretz, extend
from tshistory.api import (
    altsources,
    dbtimeseries
)


@extend(dbtimeseries)
def values_markers_origins(
        self,
        name,
        revision_date=None,
        from_value_date=None,
        to_value_date=None,
        delta=None):

    revision_date = ensuretz(revision_date)

    with self.engine.begin() as cn:
        if self.tsh.exists(cn, name):
            return self.tsh.get_many(
                cn,
                name,
                revision_date,
                from_value_date,
                to_value_date,
                delta
            )

    return self.othersources.values_markers_origins(
        name,
        revision_date,
        from_value_date,
        to_value_date,
        delta
    )


@extend(altsources)
def values_markers_origins(
        self,
        name,
        revision_date=None,
        from_value_date=None,
        to_value_date=None,
        delta=None):

    source = self._findsourcefor(name)
    if source is None:
        return None, None, None

    return source.tsa.values_markers_origins(
        name,
        revision_date,
        from_value_date,
        to_value_date,
        delta
    )
