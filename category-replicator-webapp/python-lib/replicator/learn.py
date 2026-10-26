"""Helpers that work out a sample's *row routing*: which input rows each output sheet holds.

``align``      lines the input rows up with a sample output sheet (by row order or by an identifying column)
``build_tree`` finds a small IF-chain explaining which rows were kept (used for the input filter)
"""

from __future__ import annotations

import itertools
import math
from typing import Any, Dict, List, Optional, Tuple

import numpy as np
import pandas as pd

from .rules import eq_series, nk, numeric_share, to_num

SAMPLE = 600


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
