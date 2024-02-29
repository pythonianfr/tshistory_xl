from tshistory_supervision.tsio import timeseries as supervisionts
from tshistory_formula.tsio import timeseries as formulats
from tshistory_formula import interpreter

# registration
import tshistory_formula.funcs
import tshistory_formula.api  # noqa: F401

import tshistory_supervision.api  # noqa: F401

import tshistory_xl.api
import tshistory_xl.funcs  # noqa: F401


class timeseries(supervisionts, formulats):
    index = 2
    metadata_compat_excluded = ('supervision_status',)

    def get_many(self, cn, name,
                 revision_date=None,
                 from_value_date=None,
                 to_value_date=None,
                 delta=None):

        ts_values = None
        ts_marker = None
        ts_origins = None
        if not self.exists(cn, name):
            return ts_values, ts_marker, ts_origins

        formula = self.formula(cn, name)
        if formula and formula.startswith('(prio') and not delta:
            # now we must take care of the priority formula
            # in this case: we need to compute the origins
            formula = formula.replace('(priority', '(priority-origin', 1)
            kw = {
                'revision_date': revision_date,
                'from_value_date': from_value_date,
                'to_value_date': to_value_date
            }
            i = interpreter.Interpreter(cn, self, kw)
            expanded = self._expanded_formula(
                cn, formula, display=False, remote=False, qargs=kw
            )
            ts_values, ts_origins = i.evaluate(expanded)
            ts_values.name = name
            ts_origins.name = name
            return ts_values, ts_marker, ts_origins

        if delta:
            ts_values = self.staircase(
                cn, name,
                delta=delta,
                from_value_date=from_value_date,
                to_value_date=to_value_date
            )
        elif formula:
            ts_values = self.get(
                cn, name,
                revision_date=revision_date,
                from_value_date=from_value_date,
                to_value_date=to_value_date
            )
        else:
            ts_values, ts_marker = self.get_ts_marker(
                cn, name,
                revision_date=revision_date,
                from_value_date=from_value_date,
                to_value_date=to_value_date
            )
        return ts_values, ts_marker, ts_origins
