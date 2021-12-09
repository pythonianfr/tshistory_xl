from psyl import lisp

from tshistory_supervision.tsio import timeseries as supervisionts
from tshistory_formula.tsio import timeseries as formulats
from tshistory_formula import interpreter

# registration
import tshistory_formula.funcs
import tshistory_formula.api

import tshistory_supervision.api

import tshistory_xl.api
import tshistory_xl.funcs


class timeseries(supervisionts, formulats):
    _forbidden_chars = ' (),;=[]'
    metadata_compat_excluded = ('supervision_status',)

    def update(self, cn, ts, name, author,
               metadata=None,
               insertion_date=None,
               manual=False):
        name = self._sanitize(name)
        return super().update(
            cn, ts, name, author,
            metadata=metadata,
            insertion_date=insertion_date,
            manual=manual
        )

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
            i = interpreter.Interpreter(
                cn, self, {
                    'revision_date': revision_date,
                    'from_value_date': from_value_date,
                    'to_value_date':to_value_date
                }
            )
            ts_values, ts_origins = i.evaluate(lisp.parse(formula))
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

    def _sanitize(self, name):
        for char in self._forbidden_chars:
            name = name.replace(char, '')
        return name
