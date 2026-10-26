"""Learn a rule for every output column from (input, mapping tables, expected output).

Pipeline per output column (first perfect hit wins, otherwise the best partial):
  formula in the output file  ->  constant  ->  copy  ->  text / date part  ->  lookup in a
  mapping table (exact, multi-key, approximate)  ->  arithmetic (incl. looked-up operands,
  linear combos, rounding)  ->  concatenation  ->  decision tree (IF / bands)  ->
  learned value map.

Every candidate is validated by running the *same* ``rules.apply_rule`` that production runs
use, so what is learned is exactly what is replayed.
"""

from __future__ import annotations

import itertools
import math
import re
from collections import Counter
from typing import Any, Dict, List, Optional, Tuple

import numpy as np
import pandas as pd

from .rules import (Ctx, RuleError, apply_rule, date_share, disp, eq_series, is_blank, key_series,
                    nk, numeric_share, rule_columns, to_dates, to_num, _round_half_up, describe)

PERFECT = 0.9995
APPROX = 0.90
SAMPLE = 600


# ================================================================== alignment
def _unique_ratio(s: pd.Series) -> float:
    k = s.map(nk)
    nb = k[k != ""]
    return float(nb.nunique() / len(k)) if len(k) else 0.0


def align(X: pd.DataFrame, Y: pd.DataFrame) -> Tuple[pd.DataFrame, np.ndarray, Dict[str, Any]]:
    """Return (X rows matching Y rows in Y's order, bool mask of X rows present in Y, info)."""
    X = X.reset_index(drop=True)
    Y = Y.reset_index(drop=True)
    common = [(c, d) for c in X.columns for d in Y.columns if nk(c) == nk(d)]
    info: Dict[str, Any] = {"input_rows": len(X), "output_rows": len(Y)}

    pos_score = -1.0
    if len(X) == len(Y):
        if common:
            sc = [eq_series(X[c].iloc[:SAMPLE], Y[d].iloc[:SAMPLE]).mean() for c, d in common[:6]]
            pos_score = float(np.max(sc))
        else:
            pos_score = 0.5
        if pos_score >= 0.97:
            info.update(method="row order", key=None)
            return X.copy(), np.ones(len(X), bool), info

    # key join: a (composite of up to 2) column(s) that identify rows on both sides
    def make_key(df, cols):
        ks = [df[c].map(nk) for c in cols]
        k = ks[0]
        for extra in ks[1:]:
            k = k + "\x1f" + extra
        return k

    options: List[Tuple[float, List[str], List[str]]] = []
    singles = []
    for c, d in common or []:
        singles.append(([c], [d]))
    if not common:  # different names: compare values
        for d in Y.columns:
            if _unique_ratio(Y[d]) < 0.98:
                continue
            ys = set(Y[d].map(nk))
            for c in X.columns:
                xs = set(X[c].map(nk))
                if len(ys & xs) >= 0.98 * len(ys):
                    singles.append(([c], [d]))
    for xc, yc in singles:
        yk = make_key(Y, yc)
        xk = make_key(X, xc)
        if (yk != "").mean() < 0.98:
            continue
        cov = float(yk.isin(set(xk)).mean())
        uq = float(yk.nunique() / len(yk))
        if cov >= 0.98 and uq >= 0.98:
            options.append((cov + uq + 0.01 * (nk(xc[0]) == nk(yc[0])), xc, yc))
    if not options and len(common) >= 2:
        for (c1, d1), (c2, d2) in itertools.combinations(common[:8], 2):
            yk = make_key(Y, [d1, d2])
            xk = make_key(X, [c1, c2])
            cov = float(yk.isin(set(xk)).mean())
            uq = float(yk.nunique() / len(yk))
            if cov >= 0.98 and uq >= 0.98:
                options.append((cov + uq, [c1, c2], [d1, d2]))
    if options:
        options.sort(key=lambda t: -t[0])
        _, xc, yc = options[0]
        xk, yk = make_key(X, xc), make_key(Y, yc)
        first: Dict[str, int] = {}
        for i, k in enumerate(xk):
            first.setdefault(k, i)
        idx = [first[k] for k in yk if k in first]
        if len(idx) == len(Y):
            mask = np.zeros(len(X), bool)
            mask[list(set(idx))] = True
            info.update(method="key", key=xc if xc == yc else [f"{a} = {b}" for a, b in zip(xc, yc)])
            return X.iloc[idx].reset_index(drop=True), mask, info
    if len(X) == len(Y):
        info.update(method="row order", key=None,
                    warning="No identifying key found; assumed the output is in the same row order.")
        return X.copy(), np.ones(len(X), bool), info
    raise ValueError(
        "Could not line up the input rows with the output rows (%d vs %d). Make sure at least one "
        "identifying column (an ID / reference) is carried into the output." % (len(X), len(Y)))


# =============================================================== small helpers
def _pick_sample(n: int) -> np.ndarray:
    if n <= SAMPLE:
        return np.arange(n)
    return np.sort(np.random.RandomState(7).choice(n, SAMPLE, replace=False))


def _decimals(y: np.ndarray) -> int:
    d = 0
    for dd in range(0, 7):
        if np.allclose(np.round(y, dd), y, rtol=0, atol=1e-9):
            return dd
        d = dd
    return 7


def _clean_const(c: float, check) -> float:
    """Shortest decimal representation of a learned constant that still satisfies ``check``."""
    for d in range(0, 10):
        v = round(c, d)
        if check(v):
            return int(v) if float(v).is_integer() else v
    return c


def _num_arrays(work: pd.DataFrame, cols: List[str]) -> Dict[str, np.ndarray]:
    return {c: to_num(work[c]).values.astype(float) for c in cols}


def _fit(pred: np.ndarray, y: np.ndarray, ydec: int) -> Optional[Tuple[Optional[int], Optional[Any]]]:
    """Does ``pred`` (maybe rounded) reproduce ``y``? -> (round_digits, on_error) or None."""
    ok = np.isfinite(pred)
    if ok.sum() < max(1, 0.5 * len(pred)):
        return None
    bad = ~ok
    on_err = None
    if bad.any():
        yb = y[bad]
        yb = yb[~np.isnan(yb)]
        if len(yb) == 0:
            on_err = None
        elif np.allclose(yb, yb[0]):
            on_err = float(yb[0])
        else:
            return None
    for dg in (None, ydec):
        p = pred[ok] if dg is None else _round_half_up(pd.Series(pred[ok]), dg).values
        if np.allclose(p, y[ok], rtol=1e-7, atol=1e-6, equal_nan=False):
            return dg, on_err
    return None


def _col(c: str) -> Dict[str, Any]:
    return {"col": c}


def _op(op: str, l, r) -> Dict[str, Any]:
    return {"op": op, "l": l, "r": r}


def _const(v) -> Dict[str, Any]:
    return {"const": v}


# ============================================================== the learner
class Learner:
    def __init__(self, X: pd.DataFrame, Y: pd.DataFrame, tables: Dict[str, pd.DataFrame],
                 formulas: Optional[Dict[str, Dict[str, Any]]] = None,
                 toggles: Optional[List[Dict[str, Any]]] = None, y_known: bool = True,
                 variants: Optional[Dict[str, Dict[str, Dict[str, Any]]]] = None,
                 sheets: Optional[pd.Series] = None,
                 sheet_columns: Optional[Dict[str, List[str]]] = None):
        self.X = X.reset_index(drop=True)
        self.Y = Y.reset_index(drop=True)
        self.tables = tables
        self.formulas = formulas or {}
        self.sheet_columns = sheet_columns or {}
        self.input_cols = {nk(c): c for c in X.columns}
        self.variants = variants or {}            # column -> {output sheet: formula}
        self.sheets = sheets.reset_index(drop=True) if sheets is not None else None
        self.toggles = toggles or []
        self.y_known = y_known
        self.ctx = Ctx(tables, self.toggles)
        self.work = self.X.copy()
        self.n = len(self.X)
        self.S = _pick_sample(self.n)
        self._tkey_cache: Dict[Tuple[str, str], Tuple[set, float]] = {}

    # ---------------------------------------------------------------- driver
    def learn(self, out_cols: List[str]) -> List[Dict[str, Any]]:
        rules = []
        for c in out_cols:
            y = self.Y[c]
            pm = self._present(c)
            if pm is not None and not pm.all():
                rule, m, status = self._learn_subset(c, pm)
            else:
                rule, m, status = self._learn_one(c, y)
            rule["name"] = c
            rule["match"] = round(float(m), 4) if m is not None else None
            rule["status"] = status
            rules.append(rule)
            self.work[c] = y.values  # downstream columns learn from the true values
        return rules

    def _present(self, col: str) -> Optional[np.ndarray]:
        """Rows whose output sheet actually has this column (None = all rows)."""
        if not self.sheet_columns or self.sheets is None:
            return None
        has = {sh: (col in cs) for sh, cs in self.sheet_columns.items()}
        return self.sheets.map(lambda sh: has.get(sh, True)).values.astype(bool)

    def _learn_subset(self, col: str, pm: np.ndarray):
        """Learn a column that exists on only some output sheets from just those sheets' rows."""
        import copy
        sub = copy.copy(self)
        sub.work = self.work[pm].reset_index(drop=True)
        sub.Y = self.Y[pm].reset_index(drop=True)
        sub.sheets = self.sheets[pm].reset_index(drop=True)
        sub.n = len(sub.work)
        sub.S = _pick_sample(sub.n)
        sub._tkey_cache = {}
        return sub._learn_one(col, sub.Y[col])

    def _score(self, rule: Dict[str, Any], work: pd.DataFrame, y: pd.Series) -> float:
        try:
            p = apply_rule(rule, work, self.ctx)
        except Exception:  # noqa: BLE001
            return -1.0
        return float(eq_series(p, y).mean())

    def _learn_one(self, name: str, y: pd.Series) -> Tuple[Dict[str, Any], Optional[float], str]:
        work = self.work
        best: Tuple[float, Optional[Dict[str, Any]]] = (-1.0, None)

        def consider(rule: Optional[Dict[str, Any]], full: bool = True) -> bool:
            nonlocal best
            if rule is None:
                return False
            m = self._score(rule, work, y)
            if m > best[0]:
                best = (m, rule)
            return m >= PERFECT

        def _unsupported(expr: str) -> bool:     # the translator marks what it cannot calculate
            return "TODO translate" in expr or "MANUAL(" in expr

        # 0a. formula that differs from one output sheet to the next
        if name in self.variants and self.sheets is not None:
            vr = self.variants[name]
            first = next(iter(vr.values()))
            rule = {"kind": "formula", "pandas_expr": first["pandas_expr"],
                    "excel_formula": first.get("excel_formula", ""), "source": "formula",
                    "raw_formula": first.get("sample_excel_formula"), "sample_row": first.get("sample_row"),
                    "by_sheet": {sh: {"pandas_expr": c["pandas_expr"], "excel_formula": c.get("excel_formula", ""),
                                      "raw_formula": c.get("sample_excel_formula"), "sample_row": c.get("sample_row")}
                                 for sh, c in vr.items()}}
            fw = work.copy()
            for c in self.Y.columns:
                if c != name and c not in fw.columns:
                    fw[c] = self.Y[c].values
            from .rules import _formula_one
            unknown = bool(y.map(is_blank).all())
            try:
                out = pd.Series([np.nan] * len(fw), dtype=object)
                for sh, c in vr.items():
                    m = (self.sheets == sh).values
                    if m.any():
                        out[m] = _formula_one(c["pandas_expr"], fw, self.ctx).values[m]
                unknown = bool(y.map(is_blank).all())
                mm = 1.0 if unknown else float(eq_series(out, y).mean())
            except Exception as exc:  # noqa: BLE001
                rule["error"] = str(exc)
                mm = -1.0
            rule["_pending_group"] = True
            if any(_unsupported(c["pandas_expr"]) for c in vr.values()):
                rule["excel_only"] = True
                return rule, None, "excel"
            if mm >= PERFECT:
                return rule, (None if unknown else mm), "unchecked" if unknown else "ok"
            # the sample's own formulas are the truth: keep them even if we cannot reproduce the numbers
            if mm < 0:
                rule["excel_only"] = True
                return rule, None, "excel"
            return rule, mm, "approx"

        # 0. formula found in the output workbook
        f = self.formulas.get(name)
        if f:
            fw = work.copy()
            for c in self.Y.columns:
                if c != name and c not in fw.columns:
                    fw[c] = self.Y[c].values
            rule = {"kind": "formula", "pandas_expr": f["pandas_expr"],
                    "excel_formula": f.get("excel_formula", ""), "source": "formula",
                    "raw_formula": f.get("sample_excel_formula"), "sample_row": f.get("sample_row"),
                    "notes": f.get("notes", [])}
            unknown = (not self.y_known) or bool(y.map(is_blank).all())
            if _unsupported(f["pandas_expr"]):
                rule["excel_only"] = True
                return rule, None, "excel"
            try:
                from .rules import formula_series
                p = formula_series(rule, fw, self.ctx)
                m = 1.0 if unknown else float(eq_series(p, y).mean())
            except Exception as exc:  # noqa: BLE001
                rule["error"] = str(exc)
                m = -1.0
            if m >= PERFECT:
                return rule, (None if unknown else m), "unchecked" if unknown else "ok"
            # the sample's own formula is the truth: keep it (never swap it for a guess learned from the values)
            if m < 0:
                rule["excel_only"] = True          # e.g. SUMIFS / INDEX-MATCH: Excel calculates it when opened
                return rule, None, "excel"
            return rule, m, "approx"
        if not self.y_known:
            return {"kind": "const", "value": None, "source": "unresolved"}, None, "unresolved"

        # pass-through: same name as an input column and same values -> plain copy (even if the sample happens
        # to hold a single value, e.g. every row Region = EMEA - new data must still carry its own Region)
        ic = self.input_cols.get(nk(name))
        if ic is not None and ic in work.columns and eq_series(work[ic], y).mean() >= PERFECT:
            return {"kind": "copy", "col": ic, "source": "input"}, 1.0, "ok"

        yb = y.map(is_blank)
        if yb.all():
            return {"kind": "const", "value": None, "source": "inferred"}, 1.0, "ok"

        steps = (self._const, self._copy, self._datepart, self._text, self._lookup, self._arith,
                 self._concat, self._cond, self._learned_map)
        for step in steps:
            try:
                rules = step(name, y)
            except Exception:  # noqa: BLE001  - a failing heuristic must never stop learning
                continue
            for r in rules if isinstance(rules, list) else [rules]:
                if r is None:
                    continue
                r.setdefault("source", "inferred")
                if consider(r):
                    return r, best[0], "ok"
        m, rule = best
        if rule is not None and m >= APPROX:
            return rule, m, "approx"
        if rule is not None and m >= 0.5:
            return rule, m, "unresolved"
        return {"kind": "const", "value": None, "source": "unresolved"}, max(m, 0.0), "unresolved"

    # ------------------------------------------------------------ constant / copy
    def _const(self, name, y):
        k = y.map(nk)
        if k.nunique() == 1:
            v = next((v for v in y if not is_blank(v)), None)
            return {"kind": "const", "value": _py(v)}
        return None

    def _copy(self, name, y):
        res = []
        for c in self.work.columns:
            if eq_series(self.work[c].iloc[self.S], y.iloc[self.S]).mean() >= 0.97:
                res.append((0 if nk(c) == nk(name) else 1, c))
        res.sort()
        return [{"kind": "copy", "col": c} for _, c in res[:3]]

    # -------------------------------------------------------------------- text
    def _text(self, name, y):
        ys = y.map(disp)
        if (ys == "").all():
            return None
        lens = ys[ys != ""].map(len)
        cands = []
        const_len = int(lens.iloc[0]) if lens.nunique() == 1 else None
        for c in self.work.columns:
            d = self.work[c].map(disp)
            if not (d != "").any():
                continue
            for fn in ("upper", "lower", "strip", "title", "str"):
                cands.append({"kind": "text", "col": c, "fn": fn})
            if const_len:
                cands.append({"kind": "text", "col": c, "fn": "left", "n": const_len})
                cands.append({"kind": "text", "col": c, "fn": "right", "n": const_len})
        # cheap sample screen then full verify by caller
        sw, sy = self.work.iloc[self.S], y.iloc[self.S]
        scored = []
        for r in cands:
            m = self._score(r, sw, sy)
            if m >= 0.9:
                scored.append((-m, r["fn"] == "str", r))
        scored.sort(key=lambda t: (t[0], t[1]))
        return [r for _, _, r in scored[:4]]

    def _datepart(self, name, y):
        out = []
        sw, sy = self.work.iloc[self.S], y.iloc[self.S]
        for c in self.work.columns:
            if numeric_share(self.work[c]) > 0.5 or date_share(sw[c]) < 0.9:
                continue
            for part in ("year", "month", "day", "quarter", "monthname", "weekday", "ym"):
                r = {"kind": "datepart", "col": c, "part": part}
                if self._score(r, sw, sy) >= 0.97:
                    out.append(r)
            for fmt in ("%d/%m/%Y", "%m/%d/%Y", "%Y-%m-%d", "%b-%y", "%b %Y", "%B %Y", "%Y%m", "%d-%b-%Y", "%y"):
                r = {"kind": "datepart", "col": c, "part": "fmt", "fmt": fmt}
                if self._score(r, sw, sy) >= 0.97:
                    out.append(r)
        return out

    # ------------------------------------------------------------------ lookups
    def _tkeys(self, tname: str, col: str) -> Tuple[set, float]:
        k = (tname, col)
        if k not in self._tkey_cache:
            ks = self.tables[tname][col].map(nk)
            nb = ks[ks != ""]
            self._tkey_cache[k] = (set(nb), float(nb.nunique() / max(1, len(nb))))
        return self._tkey_cache[k]

    def _key_pairs(self, min_cov: float = 0.5) -> List[Tuple[float, str, str, str]]:
        """(coverage, table, table_key_col, work_col) candidates."""
        sw = self.work.iloc[self.S]
        res = []
        xk = {c: sw[c].map(nk) for c in sw.columns}
        for tname, T in self.tables.items():
            for tc in T.columns:
                tset, uq = self._tkeys(tname, tc)
                if uq < 0.9 or len(tset) == 0:
                    continue
                for xc, ks in xk.items():
                    nb = ks[ks != ""]
                    if len(nb) == 0:
                        continue
                    cov = float(nb.isin(tset).mean())
                    if cov >= min_cov:
                        res.append((cov + (0.001 if nk(xc) == nk(tc) else 0), tname, tc, xc))
        res.sort(key=lambda t: -t[0])
        return res

    def _lookup(self, name, y):
        rules: List[Dict[str, Any]] = []
        pairs = self._key_pairs(0.5)
        if not pairs:
            return None
        sw, sy = self.work.iloc[self.S], y.iloc[self.S]
        scored = []
        for cov, tname, tc, xc in pairs[:40]:
            T = self.tables[tname]
            idx = self.ctx.index(tname, [tc])
            keys = sw[xc].map(nk)
            for vc in T.columns:
                if vc == tc:
                    continue
                vals = T[vc].tolist()
                pred = [vals[idx[(k,)]] if (k,) in idx else np.nan for k in keys]
                miss = np.array([(k,) not in idx for k in keys])
                default = None
                if miss.any():
                    ym = sy[miss]
                    kk = ym.map(nk)
                    if kk.nunique() == 1 and (kk.iloc[0] != ""):
                        default = _py(next(v for v in ym if not is_blank(v)))
                p = pd.Series([default if m_ and default is not None else v for v, m_ in zip(pred, miss)],
                              dtype=object)
                m = float(eq_series(p, sy).mean())
                if m >= 0.9:
                    scored.append((-m, -cov, tname, tc, xc, vc, default))
        scored.sort(key=lambda t: t[:2])
        for _, _, tname, tc, xc, vc, default in scored[:3]:
            rules.append({"kind": "lookup", "table": tname, "keys": [xc], "table_keys": [tc],
                          "value": vc, "default": default})
        if rules:
            return rules
        # two-column key
        goods = [(cov, tname, tc, xc) for cov, tname, tc, xc in self._key_pairs(0.9)]
        by_t: Dict[str, List] = {}
        for g in goods:
            by_t.setdefault(g[1], []).append(g)
        for tname, gs in by_t.items():
            T = self.tables[tname]
            tcs = list(dict.fromkeys(g[2] for g in gs))
            xcs = list(dict.fromkeys(g[3] for g in gs))
            for t1, t2 in itertools.permutations(tcs, 2):
                for x1, x2 in itertools.permutations(xcs, 2):
                    idx = self.ctx.index(tname, [t1, t2])
                    if len(idx) < 0.9 * len(T):
                        continue
                    k1, k2 = sw[x1].map(nk), sw[x2].map(nk)
                    for vc in T.columns:
                        if vc in (t1, t2):
                            continue
                        vals = T[vc].tolist()
                        p = pd.Series([vals[idx[(a, b)]] if (a, b) in idx else np.nan
                                       for a, b in zip(k1, k2)], dtype=object)
                        if eq_series(p, sy).mean() >= PERFECT:
                            rules.append({"kind": "lookup", "table": tname, "keys": [x1, x2],
                                          "table_keys": [t1, t2], "value": vc, "default": None})
                            return rules
        # approximate (band) lookups on a numeric key
        for tname, T in self.tables.items():
            for tc in T.columns:
                if numeric_share(T[tc]) < 0.98:
                    continue
                for xc in sw.columns:
                    if numeric_share(sw[xc]) < 0.98:
                        continue
                    for vc in T.columns:
                        if vc == tc:
                            continue
                        r = {"kind": "lookup", "table": tname, "keys": [xc], "table_keys": [tc],
                             "value": vc, "default": None, "approx": True}
                        if self._score(r, sw, sy) >= PERFECT:
                            return [r]
        return None

    # --------------------------------------------------------------- arithmetic
    def _lookup_operands(self, y) -> List[Tuple[str, Dict[str, Any]]]:
        """Numeric columns of mapping tables reachable from a well-covered key."""
        ops = []
        for cov, tname, tc, xc in self._key_pairs(0.9)[:6]:
            T = self.tables[tname]
            for vc in T.columns:
                if vc == tc or numeric_share(T[vc]) < 0.98:
                    continue
                spec = {"table": tname, "keys": [xc], "table_keys": [tc], "value": vc, "default": None}
                ops.append((f"{tname}.{vc}", spec))
        return ops[:6]

    def _arith(self, name, y):
        if numeric_share(y) < 0.95:
            return None
        sw = self.work.iloc[self.S].reset_index(drop=True)
        ys = pd.to_numeric(y.iloc[self.S].where(~y.iloc[self.S].map(is_blank)), errors="coerce").fillna(0).values
        if np.allclose(ys, 0):
            return None
        ydec = _decimals(ys)
        cols = [c for c in sw.columns if numeric_share(sw[c]) >= 0.95]
        arrs = _num_arrays(sw, cols)
        nodes: Dict[str, Dict[str, Any]] = {c: _col(c) for c in cols}
        from .rules import lookup_series
        for label, spec in self._lookup_operands(y):
            try:
                arrs[label] = to_num(lookup_series(spec, sw, self.ctx)).values.astype(float)
                nodes[label] = {"lookup": spec}
            except Exception:  # noqa: BLE001
                continue
        names = [k for k, v in arrs.items() if np.isfinite(v).all() or np.isfinite(v).mean() > 0.9]
        if not names:
            return None
        found: List[Dict[str, Any]] = []

        def emit(expr, fit):
            dg, on_err = fit
            found.append({"kind": "arith", "expr": expr, "round": dg, "on_error": on_err})

        def try_pred(expr, pred) -> bool:
            f = _fit(pred, ys, ydec)
            if f is not None:
                emit(expr, f)
                return True
            return False

        def apply_op(op, a, b):
            with np.errstate(all="ignore"):
                if op == "+":
                    return a + b
                if op == "-":
                    return a - b
                if op == "*":
                    return a * b
                return np.where(b != 0, a / np.where(b == 0, 1, b), np.nan)

        # abs / pairs with max/min
        for a in names:
            if try_pred({"fn": "abs", "args": [nodes[a]]}, np.abs(arrs[a])):
                return found
        for a, b in itertools.combinations(names, 2):
            for fn, f in (("max", np.maximum), ("min", np.minimum)):
                if try_pred({"fn": fn, "args": [nodes[a], nodes[b]]}, f(arrs[a], arrs[b])):
                    return found
        # a op b
        for op in "+-*/":
            for a, b in itertools.permutations(names, 2):
                if op in "+*" and a > b:
                    continue
                if try_pred(_op(op, nodes[a], nodes[b]), apply_op(op, arrs[a], arrs[b])):
                    return found
        # a op const
        for a in names:
            x = arrs[a]
            nz = np.abs(x) > 1e-12
            if nz.sum() >= 2:
                for op, derive in (("*", lambda: np.median(ys[nz] / x[nz])),
                                   ("/", lambda: np.median(x[nz] / np.where(ys[nz] == 0, np.nan, ys[nz]))
                                    if (ys[nz] != 0).any() else math.nan)):
                    c = derive()
                    if not np.isfinite(c) or c == 0:
                        continue
                    chk = lambda v, op=op, x=x: _fit(apply_op(op, x, np.full_like(x, v)), ys, ydec) is not None
                    cv = _clean_const(float(c), chk)
                    if chk(cv):
                        emit(_op(op, nodes[a], _const(cv)), _fit(apply_op(op, x, np.full_like(x, cv)), ys, ydec))
                        return found
            for op, c in (("+", float(np.median(ys - x))), ("-", float(np.median(x - ys)))):
                chk = lambda v, op=op, x=x: _fit(apply_op(op, x, np.full_like(x, v)), ys, ydec) is not None
                cv = _clean_const(c, chk)
                if chk(cv) and cv != 0:
                    emit(_op(op, nodes[a], _const(cv)), _fit(apply_op(op, x, np.full_like(x, cv)), ys, ydec))
                    return found
        # a * (alpha + beta*b)    e.g. exposure * (1 - crm)
        for a, b in itertools.permutations(names, 2):
            x, z = arrs[a], arrs[b]
            M = np.column_stack([x, x * z])
            if np.linalg.matrix_rank(M) < 2:
                continue
            w, *_ = np.linalg.lstsq(M, ys, rcond=None)
            al = _snap(w[0], x, z, ys, ydec, 0)
            be = _snap(w[1], x, z, ys, ydec, 1, al)
            pred = x * (al + be * z)
            if be != 0 and try_pred(_op("*", nodes[a], _lin(al, be, nodes[b])), pred):
                return found
        # (a op b) op2 c   and   c op2 (a op b)
        P = np.column_stack([arrs[n] for n in names])
        for op in "+-*/":
            for ia, ib in itertools.permutations(range(len(names)), 2):
                if op in "+*" and ia > ib:
                    continue
                v = apply_op(op, P[:, ia], P[:, ib])
                for op2 in "+-*/":
                    V = np.column_stack([v] * len(names))
                    R = apply_op(op2, V, P)
                    R2 = apply_op(op2, P, V)
                    for ic in range(len(names)):
                        if ic in (ia, ib) and op2 in "-/":
                            pass
                        for pred, rev in ((R[:, ic], False), (R2[:, ic], True)):
                            f = _fit(pred, ys, ydec)
                            if f is not None:
                                inner = _op(op, nodes[names[ia]], nodes[names[ib]])
                                c = nodes[names[ic]]
                                emit(_op(op2, c, inner) if rev else _op(op2, inner, c), f)
                                return found
                # (a op b) op2 const
                nz = np.abs(v) > 1e-12
                if np.isfinite(v).all() and nz.sum() >= 2:
                    c = float(np.median(ys[nz] / v[nz]))
                    if c != 0:
                        chk = lambda t, v=v: _fit(v * t, ys, ydec) is not None
                        cv = _clean_const(c, chk)
                        if chk(cv):
                            emit(_op("*", _op(op, nodes[names[ia]], nodes[names[ib]]), _const(cv)),
                                 _fit(v * cv, ys, ydec))
                            return found
        # sparse linear combination (a + b - c, 0.3*a + 0.7*b ...)
        if len(names) >= 2 and len(ys) >= 4:
            res = self._omp(names, arrs, nodes, ys, ydec)
            if res:
                return [res]
        return None

    def _omp(self, names, arrs, nodes, ys, ydec):
        chosen: List[str] = []
        resid = ys.copy()
        for _ in range(min(5, len(names))):
            best, bs = None, 0.0
            for nme in names:
                if nme in chosen:
                    continue
                a = arrs[nme]
                d = np.linalg.norm(a)
                if d < 1e-12:
                    continue
                sc = abs(float(a @ resid)) / d
                if sc > bs:
                    best, bs = nme, sc
            if best is None:
                break
            chosen.append(best)
            M = np.column_stack([arrs[c] for c in chosen])
            if len(ys) < len(chosen) + 2:
                break
            w, *_ = np.linalg.lstsq(M, ys, rcond=None)
            resid = ys - M @ w
            wr = np.array([_snap_w(float(x)) for x in w])
            pred = M @ wr
            f = _fit(pred, ys, ydec)
            if f is not None:
                expr = None
                for c, wt in zip(chosen, wr):
                    if wt == 0:
                        continue
                    term = nodes[c] if wt == 1 else _op("*", _const(_numfmt(wt)), nodes[c])
                    if expr is None:
                        expr = term if wt >= 0 or wt == 1 else _op("*", _const(_numfmt(wt)), nodes[c])
                    else:
                        if wt == -1:
                            expr = _op("-", expr, nodes[c])
                        elif wt < 0:
                            expr = _op("-", expr, _op("*", _const(_numfmt(-wt)), nodes[c]))
                        else:
                            expr = _op("+", expr, term)
                if expr is not None:
                    return {"kind": "arith", "expr": expr, "round": f[0], "on_error": f[1]}
        return None

    # -------------------------------------------------------------------- concat
    def _concat(self, name, y):
        ys = y.map(disp)
        probe = [i for i in self.S if ys.iloc[i] != ""][:6]
        if not probe:
            return None
        cols = list(self.work.columns)
        disp_cols = {c: self.work[c].map(disp) for c in cols}
        i0 = probe[0]
        y0 = ys.iloc[i0]
        occ: List[Tuple[int, int, str]] = []
        for c in cols:
            v = disp_cols[c].iloc[i0]
            if not v:
                continue
            if sum(1 for i in probe if disp_cols[c].iloc[i] and disp_cols[c].iloc[i] in ys.iloc[i]) < len(probe):
                continue
            for m in re.finditer(re.escape(v), y0):
                occ.append((m.start(), m.end(), c))
        if not occ:
            return None
        occ.sort()
        patterns = []

        def dfs(pos, seq, budget=[3000]):
            budget[0] -= 1
            if budget[0] < 0:
                return
            if seq:
                patterns.append(list(seq))
            if len(seq) >= 3:
                return
            for s, e, c in occ:
                if s >= pos:
                    seq.append((s, e, c))
                    dfs(e, seq)
                    seq.pop()
        dfs(0, [])
        scored = []
        sw, sy = self.work.iloc[self.S], y.iloc[self.S]
        for seq in patterns:
            parts, pos = [], 0
            for s, e, c in seq:
                if s > pos:
                    parts.append({"text": y0[pos:s]})
                parts.append({"col": c})
                pos = e
            if pos < len(y0):
                parts.append({"text": y0[pos:]})
            r = {"kind": "concat", "parts": parts}
            m = self._score(r, sw, sy)
            if m >= 0.9:
                scored.append((-m, len(parts), r))
        scored.sort(key=lambda t: t[:2])
        return [r for _, _, r in scored[:2]]

    # ------------------------------------------------------------ decision trees
    def _cond(self, name, y):
        tree = build_tree(self.work, y, max_depth=3, max_leaves=10)
        if tree is None:
            return None
        rule = {"kind": "cond", "tree": tree}
        return rule

    def _learned_map(self, name, y):
        sw = self.work
        lowcard = [c for c in sw.columns
                   if 1 < sw[c].map(nk).nunique() <= max(2, self.n // 2) and sw[c].map(nk).nunique() <= 500]
        yk = y.map(nk)
        for size in (1, 2):
            for combo in itertools.combinations(lowcard[:40], size):
                kk = list(zip(*[sw[c].map(nk) for c in combo]))
                grp: Dict[Any, set] = {}
                for k, v in zip(kk, yk):
                    grp.setdefault(k, set()).add(v)
                if all(len(v) == 1 for v in grp.values()) and len(grp) <= self.n / 1.5:
                    first = {}
                    for i, (k, v) in enumerate(zip(kk, y)):
                        first.setdefault(k, (i, _py(v)))
                    pairs = []
                    for k, (i, v) in first.items():
                        pairs.append([_py(sw[c].iloc[i]) for c in combo] + [v])
                    return {"kind": "learned_map", "keys": list(combo), "pairs": pairs, "default": None}
        return None


# ===================================================================== helpers
def _py(v):
    if isinstance(v, (np.integer,)):
        return int(v)
    if isinstance(v, (np.floating,)):
        f = float(v)
        return None if math.isnan(f) else (int(f) if f.is_integer() else f)
    if isinstance(v, float) and math.isnan(v):
        return None
    if v is pd.NaT:
        return None
    if isinstance(v, (pd.Timestamp,)):
        return v.isoformat()
    return v


def _numfmt(x: float):
    return int(x) if float(x).is_integer() else x


def _snap_w(w: float) -> float:
    for d in range(0, 7):
        r = round(w, d)
        if abs(r - w) < 1e-6 * max(1.0, abs(w)):
            return _numfmt(r)
    return w


def _snap(w, x, z, ys, ydec, which, al=None):
    for d in range(0, 7):
        r = round(float(w), d)
        if abs(r - w) < 1e-6 * max(1.0, abs(w)):
            return _numfmt(r)
    return float(w)


def _lin(al, be, b_node):
    if al == 0:
        return b_node if be == 1 else _op("*", _const(be), b_node)
    if be == 1:
        return _op("+", _const(al), b_node)
    if be == -1:
        return _op("-", _const(al), b_node)
    return _op("+", _const(al), _op("*", _const(be), b_node))


# ----------------------------------------------------------------- tree learner
def _nice_threshold(lo: float, hi: float) -> float:
    """Roundest number t with lo < t <= hi (so test is `>= t`)."""
    for k in range(12, -7, -1):
        step = 10.0 ** k
        t = math.floor(lo / step + 1e-12) * step + step
        if t <= hi + 1e-12 and t > lo:
            return float(round(t, 9))
    return hi


def build_tree(df: pd.DataFrame, y: pd.Series, max_depth: int = 3, max_leaves: int = 10,
               feature_cols: Optional[List[str]] = None) -> Optional[Dict[str, Any]]:
    yk = y.map(nk).values
    classes = list(dict.fromkeys(yk))
    if len(classes) < 2 or len(classes) > 12:
        return None
    y_orig = {}
    for k, v in zip(yk, y):
        y_orig.setdefault(k, _py(v))
    cls_idx = {c: i for i, c in enumerate(classes)}
    yi = np.array([cls_idx[k] for k in yk])
    K = len(classes)
    cols = feature_cols or list(df.columns)
    feats = {}
    for c in cols:
        s = df[c]
        if numeric_share(s) >= 0.97:
            feats[c] = ("num", to_num(s).values.astype(float))
        else:
            feats[c] = ("cat", s.map(nk).values)
    num_cols = [c for c, (k, _) in feats.items() if k == "num"][:12]
    pairs = [(a, b, op) for a in num_cols for b in num_cols if a != b for op in (">", ">=")]
    leaves = [0]

    def leaf(idx):
        cnt = np.bincount(yi[idx], minlength=K)
        return {"leaf": {"const": y_orig[classes[int(cnt.argmax())]]}}

    def best_split(idx):
        base = np.bincount(yi[idx], minlength=K).max()
        best = (base, None)
        for c, (kind, arr) in feats.items():
            v = arr[idx]
            yy = yi[idx]
            if kind == "num":
                order = np.argsort(v, kind="stable")
                vs, ys_ = v[order], yy[order]
                oh = np.zeros((len(vs), K))
                oh[np.arange(len(vs)), ys_] = 1
                cum = np.cumsum(oh, axis=0)
                tot = cum[-1]
                for i in range(len(vs) - 1):
                    if vs[i] == vs[i + 1]:
                        continue
                    score = cum[i].max() + (tot - cum[i]).max()
                    if score > best[0] + 1e-9:
                        t = _nice_threshold(float(vs[i]), float(vs[i + 1]))
                        best = (score, {"col": c, "op": ">=", "value": _numfmt(t)})
            else:
                for val in np.unique(v):
                    if val == "":
                        continue
                    m = v == val
                    if m.all():
                        continue
                    score = np.bincount(yy[m], minlength=K).max() + np.bincount(yy[~m], minlength=K).max()
                    if score > best[0] + 1e-9:
                        orig = next(df[c].iloc[j] for j in idx if arr[j] == val)
                        best = (score, {"col": c, "op": "==", "value": _py(orig)})
        for a, b, op in pairs:
            va, vb = feats[a][1][idx], feats[b][1][idx]
            m = (va > vb) if op == ">" else (va >= vb)
            if m.all() or not m.any():
                continue
            score = np.bincount(yy_all(idx)[m], minlength=K).max() + np.bincount(yy_all(idx)[~m], minlength=K).max()
            if score > best[0] + 1e-9:
                best = (score, {"col": a, "op": op, "value_col": b})
        return best[1]

    def yy_all(idx):
        return yi[idx]

    def split_mask(test):
        if test.get("value_col"):
            xa, xb = feats[test["col"]][1], feats[test["value_col"]][1]
            return (xa > xb) if test["op"] == ">" else (xa >= xb)
        arr_kind, arr = feats[test["col"]]
        if test["op"] == ">=":
            return arr >= float(test["value"])
        return arr == nk(test["value"])

    def rec(idx, depth):
        if len(set(yi[idx])) == 1 or depth >= max_depth:
            leaves[0] += 1
            return leaf(idx)
        test = best_split(idx)
        if test is None:
            leaves[0] += 1
            return leaf(idx)
        m = split_mask(test)
        a, b = idx[m[idx]], idx[~m[idx]]
        if len(a) == 0 or len(b) == 0:
            leaves[0] += 1
            return leaf(idx)
        return {"test": test, "then": rec(a, depth + 1), "else": rec(b, depth + 1)}

    tree = rec(np.arange(len(df)), 0)
    if leaves[0] > max_leaves or "leaf" in tree:
        return None
    return tree
