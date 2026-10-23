"""Rule engine: one *rule* describes how to build one output column.

A rule is a plain JSON dict (so it can be stored, shown and edited):

    {"name": "Risk Weight", "kind": "lookup", ...kind specific fields..., "source": "inferred"}

Kinds
    copy         {"col"}
    const        {"value"}
    text         {"col", "fn": upper|lower|strip|title|left|right|mid|str, "n", "start"}
    datepart     {"col", "part": year|month|day|quarter|monthname|weekday|ym|fmt, "fmt"}
    arith        {"expr": <node>, "round": int|None, "on_error": value|None}
                 node = {"col"} | {"const"} | {"lookup": <lookup spec>} |
                        {"op": + - * /, "l": node, "r": node} | {"fn": abs|max|min, "args": [node..]}
    concat       {"parts": [{"col"} | {"text"}]}
    lookup       {"table", "keys": [cols], "table_keys": [cols], "value", "default", "approx"}
    cond         {"tree"}   tree = {"test": {"col","op","value"}, "then": tree, "else": tree}
                 or a leaf {"leaf": {"const"} | {"col"}}
    learned_map  {"keys": [cols], "pairs": [[k.., v]], "default"}
    formula      {"pandas_expr", "excel_formula"}   (Excel formula translated by the Studio parser,
                 or a manual pandas expression)

Everything here is dataiku-free and vectorised over a pandas DataFrame (``work``).
"""

from __future__ import annotations

import datetime as dt
import math
import re
from typing import Any, Callable, Dict, List, Optional, Tuple

import numpy as np
import pandas as pd

_WS = re.compile(r"\s+")


class RuleError(Exception):
    pass


# ------------------------------------------------------------------ scalar helpers
def is_blank(v: Any) -> bool:
    if v is None:
        return True
    if isinstance(v, float) and math.isnan(v):
        return True
    if v is pd.NaT:
        return True
    if isinstance(v, str) and v.strip() == "":
        return True
    return False


def disp(v: Any) -> str:
    """Excel-ish text rendering of a value (integers lose the '.0')."""
    if is_blank(v):
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, (int, np.integer)):
        return str(int(v))
    if isinstance(v, (float, np.floating)):
        f = float(v)
        return str(int(f)) if f.is_integer() else repr(f)
    if isinstance(v, (dt.datetime, pd.Timestamp)):
        return v.strftime("%Y-%m-%d") if (v.hour, v.minute, v.second) == (0, 0, 0) else v.isoformat(" ")
    if isinstance(v, dt.date):
        return v.isoformat()
    return str(v)


def nk(v: Any) -> str:
    """Normalised comparison/lookup key: case-, space- and 1 vs 1.0-insensitive."""
    return _WS.sub(" ", disp(v)).strip().lower()


def key_series(s: pd.Series) -> pd.Series:
    return s.map(nk)


def to_num(s: pd.Series) -> pd.Series:
    """Excel arithmetic coercion: numbers stay, blanks -> 0, text that is not a number -> NaN."""
    blank = s.map(is_blank)
    out = pd.to_numeric(s.where(~blank), errors="coerce")
    return out.where(~blank, 0.0).astype(float)


def numeric_share(s: pd.Series) -> float:
    nb = ~s.map(is_blank)
    if not nb.any():
        return 0.0
    return float(pd.to_numeric(s[nb], errors="coerce").notna().mean())


def to_dates(s: pd.Series) -> pd.Series:
    def one(v):
        if is_blank(v):
            return pd.NaT
        if isinstance(v, (dt.datetime, dt.date, pd.Timestamp)):
            return pd.Timestamp(v)
        if isinstance(v, (int, float, np.integer, np.floating)):
            return pd.NaT
        try:
            return pd.Timestamp(pd.to_datetime(str(v), errors="raise"))
        except Exception:  # noqa: BLE001
            return pd.NaT
    return pd.to_datetime(s.map(one), errors="coerce")


def date_share(s: pd.Series) -> float:
    nb = ~s.map(is_blank)
    if not nb.any():
        return 0.0
    return float(to_dates(s[nb]).notna().mean())


def eq_series(pred: pd.Series, expect: pd.Series) -> pd.Series:
    """Cell-by-cell 'same value?' with Excel-friendly tolerance."""
    p = pred.reset_index(drop=True)
    y = expect.reset_index(drop=True)
    pb, yb = p.map(is_blank), y.map(is_blank)
    pn = pd.to_numeric(p.where(~pb), errors="coerce")
    yn = pd.to_numeric(y.where(~yb), errors="coerce")
    both = pn.notna() & yn.notna()
    numeq = pd.Series(np.isclose(pn.fillna(0).values, yn.fillna(0).values, rtol=1e-7, atol=1e-6),
                      index=p.index)
    streq = p.map(lambda v: disp(v).strip()) == y.map(lambda v: disp(v).strip())
    return (both & numeq) | (~both & streq) | (pb & yb)


# ------------------------------------------------------------------------ context
class Ctx:
    """Tables (name -> DataFrame) + toggles + small caches, shared by every rule."""

    def __init__(self, tables: Optional[Dict[str, pd.DataFrame]] = None,
                 toggles: Optional[List[Dict[str, Any]]] = None):
        self.tables: Dict[str, pd.DataFrame] = tables or {}
        self.toggles = toggles or []
        self._idx: Dict[Tuple, Dict] = {}
        self.notes: List[str] = []            # run-time warnings (unmapped keys ...)
        self.stats: Dict[str, Dict[str, Any]] = {}

    def table(self, name: str) -> pd.DataFrame:
        if name in self.tables:
            return self.tables[name]
        low = {k.lower(): k for k in self.tables}
        if name.lower() in low:
            return self.tables[low[name.lower()]]
        raise RuleError(f"mapping table '{name}' is not available")

    def index(self, table: str, keys: List[str]) -> Dict:
        ck = (table, tuple(keys))
        if ck not in self._idx:
            df = self.table(table)
            cols = [resolve_col(df, k) for k in keys]
            kk = [key_series(df[c]) for c in cols]
            idx: Dict = {}
            for i, tup in enumerate(zip(*kk)):
                idx.setdefault(tup, i)
            self._idx[ck] = idx
        return self._idx[ck]


def resolve_col(df: pd.DataFrame, name: str) -> str:
    if name in df.columns:
        return name
    low = {nk(c): c for c in df.columns}
    if nk(name) in low:
        return low[nk(name)]
    raise RuleError(f"column '{name}' not found")


def get_col(work: pd.DataFrame, name: str) -> pd.Series:
    return work[resolve_col(work, name)]


# -------------------------------------------------------------- lookup application
def lookup_series(spec: Dict[str, Any], work: pd.DataFrame, ctx: Ctx,
                  rule_name: str = "") -> pd.Series:
    df = ctx.table(spec["table"])
    vcol = resolve_col(df, spec["value"])
    vals = df[vcol].tolist()
    keys = spec["keys"]
    if spec.get("approx"):
        kcol = resolve_col(df, spec["table_keys"][0])
        tk = pd.to_numeric(df[kcol], errors="coerce")
        order = np.argsort(tk.fillna(np.inf).values, kind="stable")
        sk = tk.values[order]
        x = to_num(get_col(work, keys[0])).values
        pos = np.searchsorted(sk, x, side="right") - 1
        out = [vals[order[p]] if p >= 0 and not math.isnan(xx) else np.nan for p, xx in zip(pos, x)]
        return pd.Series(out, index=work.index, dtype=object)
    idx = ctx.index(spec["table"], spec["table_keys"])
    kcols = [key_series(get_col(work, k)) for k in keys]
    default = spec.get("default")
    out, miss = [], 0
    missing_keys: Dict[str, int] = {}
    for tup in zip(*kcols):
        i = idx.get(tup)
        if i is None:
            miss += 1
            out.append(default if default is not None else np.nan)
            if len(missing_keys) < 8:
                lab = " | ".join(tup)
                missing_keys[lab] = missing_keys.get(lab, 0) + 1
        else:
            out.append(vals[i])
    if miss and rule_name:
        ctx.stats[rule_name] = {"unmatched": miss, "examples": list(missing_keys)[:8]}
    return pd.Series(out, index=work.index, dtype=object)


# ------------------------------------------------------------------ arithmetic tree
def _node(node: Dict[str, Any], work: pd.DataFrame, ctx: Ctx) -> pd.Series:
    if "col" in node:
        return to_num(get_col(work, node["col"]))
    if "const" in node:
        return pd.Series(float(node["const"]), index=work.index)
    if "lookup" in node:
        return to_num(lookup_series(node["lookup"], work, ctx))
    if "fn" in node:
        args = [_node(a, work, ctx) for a in node["args"]]
        fn = node["fn"]
        if fn == "abs":
            return args[0].abs()
        if fn == "max":
            return pd.concat(args, axis=1).max(axis=1)
        if fn == "min":
            return pd.concat(args, axis=1).min(axis=1)
        raise RuleError(f"unknown function {fn}")
    op = node["op"]
    l, r = _node(node["l"], work, ctx), _node(node["r"], work, ctx)
    with np.errstate(all="ignore"):
        if op == "+":
            return l + r
        if op == "-":
            return l - r
        if op == "*":
            return l * r
        if op == "/":
            return (l / r).replace([np.inf, -np.inf], np.nan)
    raise RuleError(f"unknown operator {op}")


def _round_half_up(s: pd.Series, n: int) -> pd.Series:
    f = 10.0 ** n
    v = s * f
    return np.sign(v) * np.floor(np.abs(v) + 0.5 + 1e-9) / f  # Excel ROUND (half away from zero)


# ----------------------------------------------------------------------- conditions
def cmp_mask(s: pd.Series, op: str, value: Any) -> pd.Series:
    op = {"=": "==", "<>": "!="}.get(op, op)
    blank = s.map(is_blank)
    if op in ("blank", "is empty"):
        return blank
    if op in ("not blank", "not empty"):
        return ~blank
    if op in ("in", "not in"):
        vals = value if isinstance(value, list) else [x.strip() for x in str(value).split(",")]
        m = key_series(s).isin({nk(v) for v in vals})
        return m if op == "in" else ~m
    if op in ("contains", "not contains", "startswith", "endswith"):
        t = key_series(s)
        v = nk(value)
        m = {"contains": t.str.contains(re.escape(v), regex=True),
             "not contains": ~t.str.contains(re.escape(v), regex=True),
             "startswith": t.str.startswith(v), "endswith": t.str.endswith(v)}[op]
        return m.fillna(False)
    vnum = pd.to_numeric(pd.Series([value]), errors="coerce").iloc[0]
    sn = pd.to_numeric(s.where(~blank), errors="coerce")
    if op in ("==", "!="):
        if not pd.isna(vnum) and sn.notna().mean() > 0.5:
            eq = pd.Series(np.isclose(sn.fillna(np.nan).values, float(vnum), rtol=1e-9, atol=1e-9), index=s.index)
        else:
            eq = key_series(s) == nk(value)
        return eq if op == "==" else ~eq
    if pd.isna(vnum):
        a, b = key_series(s), nk(value)
        return {">": a > b, ">=": a >= b, "<": a < b, "<=": a <= b}[op]
    return {">": sn > vnum, ">=": sn >= vnum, "<": sn < vnum, "<=": sn <= vnum}[op].fillna(False)


def _leaf(node: Dict[str, Any], work: pd.DataFrame) -> pd.Series:
    if "const" in node:
        return pd.Series([node["const"]] * len(work), index=work.index, dtype=object)
    return get_col(work, node["col"]).astype(object)


def _tree(tree: Dict[str, Any], work: pd.DataFrame) -> pd.Series:
    if "leaf" in tree:
        return _leaf(tree["leaf"], work)
    t = tree["test"]
    m = cmp_mask(get_col(work, t["col"]), t["op"], t.get("value")).values
    a, b = _tree(tree["then"], work), _tree(tree["else"], work)
    return pd.Series(np.where(m, a.values, b.values), index=work.index, dtype=object)


# -------------------------------------------------------------------------- formula
def formula_series(rule: Dict[str, Any], work: pd.DataFrame, ctx: Ctx) -> pd.Series:
    from replicator_formulas import compute  # lazy: keeps rules importable on its own
    lk = compute.Lookups({
        n: {"headers": list(df.columns), "records": df.to_dict("records"), "header_row": 0}
        for n, df in ctx.tables.items()
    })
    recipe = {"computed_columns": [{"name": "__out__", "pandas_expr": rule["pandas_expr"]}],
              "toggles": ctx.toggles}
    res = compute.evaluate(work, recipe, lk)
    errs = res.attrs.get("compute_errors")
    if errs:
        raise RuleError("; ".join(errs))
    return res["__out__"]


# ----------------------------------------------------------------------- dispatcher
def apply_rule(rule: Dict[str, Any], work: pd.DataFrame, ctx: Ctx) -> pd.Series:
    kind = rule["kind"]
    n = len(work)
    if kind == "copy":
        return get_col(work, rule["col"]).copy()
    if kind == "const":
        return pd.Series([rule.get("value")] * n, index=work.index, dtype=object)
    if kind == "text":
        s = get_col(work, rule["col"])
        d = s.map(disp)
        fn = rule["fn"]
        k = int(rule.get("n") or 1)
        out = {"upper": d.str.upper(), "lower": d.str.lower(), "strip": d.str.strip(),
               "title": d.str.title(), "left": d.str[:k], "right": d.str[-k:] if k else d.str[:0],
               "mid": d.str[int(rule.get("start", 1)) - 1:int(rule.get("start", 1)) - 1 + k],
               "str": d}[fn]
        return out
    if kind == "datepart":
        d = to_dates(get_col(work, rule["col"]))
        part = rule["part"]
        if part == "year":
            return d.dt.year.astype(float)
        if part == "month":
            return d.dt.month.astype(float)
        if part == "day":
            return d.dt.day.astype(float)
        if part == "quarter":
            return d.dt.quarter.astype(float)
        if part == "weekday":
            return d.dt.day_name()
        if part == "monthname":
            return d.dt.month_name()
        if part == "ym":
            return d.dt.strftime("%Y-%m")
        if part == "fmt":
            return d.dt.strftime(rule["fmt"])
        raise RuleError(f"unknown date part {part}")
    if kind == "arith":
        v = _node(rule["expr"], work, ctx)
        if rule.get("round") is not None:
            v = _round_half_up(v, int(rule["round"]))
        if rule.get("on_error") is not None:
            v = v.where(v.notna(), rule["on_error"])
        return v.astype(object).where(v.notna(), np.nan)
    if kind == "concat":
        parts = []
        for p in rule["parts"]:
            if "col" in p:
                parts.append(get_col(work, p["col"]).map(disp))
            else:
                parts.append(pd.Series(str(p["text"]), index=work.index))
        out = parts[0]
        for p in parts[1:]:
            out = out + p
        return out
    if kind == "lookup":
        return lookup_series(rule, work, ctx, rule.get("name", ""))
    if kind == "cond":
        return _tree(rule["tree"], work)
    if kind == "learned_map":
        keys = rule["keys"]
        mp = {tuple(nk(x) for x in p[:-1]): p[-1] for p in rule["pairs"]}
        cols = [key_series(get_col(work, k)) for k in keys]
        default = rule.get("default")
        out, miss, ex = [], 0, []
        for tup in zip(*cols):
            if tup in mp:
                out.append(mp[tup])
            else:
                miss += 1
                out.append(default if default is not None else np.nan)
                if len(ex) < 8 and " | ".join(tup) not in ex:
                    ex.append(" | ".join(tup))
        if miss and rule.get("name"):
            ctx.stats[rule["name"]] = {"unmatched": miss, "examples": ex}
        return pd.Series(out, index=work.index, dtype=object)
    if kind == "formula":
        return formula_series(rule, work, ctx)
    raise RuleError(f"unknown rule kind '{kind}'")


# ------------------------------------------------------------------- introspection
def _node_cols(node: Dict[str, Any], out: List[str]) -> None:
    if "col" in node:
        out.append(node["col"])
    elif "lookup" in node:
        out.extend(node["lookup"]["keys"])
    elif "fn" in node:
        for a in node["args"]:
            _node_cols(a, out)
    elif "op" in node:
        _node_cols(node["l"], out)
        _node_cols(node["r"], out)


def _tree_cols(t: Dict[str, Any], out: List[str]) -> None:
    if "leaf" in t:
        if "col" in t["leaf"]:
            out.append(t["leaf"]["col"])
        return
    out.append(t["test"]["col"])
    _tree_cols(t["then"], out)
    _tree_cols(t["else"], out)


def rule_columns(rule: Dict[str, Any]) -> List[str]:
    """Columns (input or earlier output) this rule reads."""
    k, out = rule["kind"], []
    if k in ("copy", "text", "datepart"):
        out.append(rule["col"])
    elif k == "arith":
        _node_cols(rule["expr"], out)
    elif k == "concat":
        out.extend(p["col"] for p in rule["parts"] if "col" in p)
    elif k in ("lookup", "learned_map"):
        out.extend(rule["keys"])
    elif k == "cond":
        _tree_cols(rule["tree"], out)
    elif k == "formula":
        out.extend(re.findall(r"df\[\s*['\"]([^'\"]+)['\"]\s*\]", rule.get("pandas_expr", "")))
    return list(dict.fromkeys(out))


def rule_tables(rule: Dict[str, Any]) -> List[str]:
    out: List[str] = []

    def walk(n):
        if isinstance(n, dict):
            if "lookup" in n:
                out.append(n["lookup"]["table"])
            for v in n.values():
                walk(v)
    if rule["kind"] == "lookup":
        out.append(rule["table"])
    walk(rule.get("expr", {}))
    if rule["kind"] == "formula":
        out.extend(re.findall(r"VLOOKUP\([^,]+,\s*['\"]([^'\"]+)['\"]", rule.get("pandas_expr", "")))
    return list(dict.fromkeys(out))


# --------------------------------------------------------------------- description
def _fmt_const(c: Any) -> str:
    return disp(c) if not isinstance(c, str) else f'"{c}"'


def _desc_node(n: Dict[str, Any]) -> str:
    if "col" in n:
        return f"[{n['col']}]"
    if "const" in n:
        return _fmt_const(n["const"])
    if "lookup" in n:
        lk = n["lookup"]
        return f"{lk['table']}.{lk['value']}[by {', '.join(lk['keys'])}]"
    if "fn" in n:
        return f"{n['fn'].upper()}({', '.join(_desc_node(a) for a in n['args'])})"
    return f"({_desc_node(n['l'])} {n['op']} {_desc_node(n['r'])})"


def _desc_tree(t: Dict[str, Any]) -> str:
    if "leaf" in t:
        return _desc_node(t["leaf"])
    x = t["test"]
    return f"IF([{x['col']}] {x['op']} {_fmt_const(x.get('value'))}, {_desc_tree(t['then'])}, {_desc_tree(t['else'])})"


def describe(rule: Dict[str, Any]) -> str:
    k = rule["kind"]
    if k == "copy":
        return f"Copy of [{rule['col']}]"
    if k == "const":
        return f"Constant {_fmt_const(rule.get('value'))}"
    if k == "text":
        extra = f", {rule.get('n')}" if rule["fn"] in ("left", "right", "mid") else ""
        return f"{rule['fn'].upper()}([{rule['col']}]{extra})"
    if k == "datepart":
        return f"{rule['part']} of [{rule['col']}]" + (f" ({rule['fmt']})" if rule["part"] == "fmt" else "")
    if k == "arith":
        s = _desc_node(rule["expr"])
        if s.startswith("(") and s.endswith(")"):
            s = s[1:-1]
        if rule.get("round") is not None:
            s = f"ROUND({s}, {rule['round']})"
        if rule.get("on_error") is not None:
            s += f"  (error -> {rule['on_error']})"
        return s
    if k == "concat":
        return " & ".join(f"[{p['col']}]" if "col" in p else f'"{p["text"]}"' for p in rule["parts"])
    if k == "lookup":
        how = "approximate" if rule.get("approx") else "exact"
        d = f", default {_fmt_const(rule['default'])}" if rule.get("default") is not None else ""
        return (f"Lookup {rule['table']}.{rule['value']} where {', '.join(rule['table_keys'])} "
                f"= {', '.join('[' + c + ']' for c in rule['keys'])} ({how}{d})")
    if k == "cond":
        return _desc_tree(rule["tree"])
    if k == "learned_map":
        return f"Learned mapping from {', '.join('[' + c + ']' for c in rule['keys'])} ({len(rule['pairs'])} entries)"
    if k == "formula":
        return rule.get("excel_formula") or rule.get("pandas_expr", "")
    return k


# --------------------------------------------------------------- live Excel formulas
def _xl_const(c: Any) -> str:
    if isinstance(c, str):
        return '"' + c.replace('"', '""') + '"'
    if isinstance(c, bool):
        return "TRUE" if c else "FALSE"
    return disp(c) if c is not None else '""'


def sheet_token(name: str) -> str:
    return "'" + name.replace("'", "''") + "'"


def excel_formula(rule: Dict[str, Any], ref: Callable[[str], str],
                  table_ref: Callable[[str, str], str]) -> Optional[str]:
    """Return an Excel formula for row ``{r}`` or None if this rule can't be expressed.

    ``ref(col)`` -> cell reference like ``C{r}``; ``table_ref(table, col)`` -> an absolute
    column range like ``'Rates'!$B:$B``.
    """
    k = rule["kind"]

    def node(n):
        if "col" in n:
            return ref(n["col"])
        if "const" in n:
            return _xl_const(n["const"])
        if "lookup" in n:
            return lk(n["lookup"])
        if "fn" in n:
            return f"{n['fn'].upper()}({','.join(node(a) for a in n['args'])})"
        return f"({node(n['l'])}{n['op']}{node(n['r'])})"

    def lk(spec):
        if len(spec["keys"]) != 1:
            raise ValueError("multi-key")
        m = 1 if spec.get("approx") else 0
        base = (f"INDEX({table_ref(spec['table'], spec['value'])},"
                f"MATCH({ref(spec['keys'][0])},{table_ref(spec['table'], spec['table_keys'][0])},{m}))")
        d = spec.get("default")
        return f"IFERROR({base},{_xl_const(d)})" if d is not None and not spec.get("approx") else base

    def tree(t):
        if "leaf" in t:
            return node(t["leaf"])
        x = t["test"]
        c = ref(x["col"])
        v = x.get("value")
        op = {"==": "=", "!=": "<>"}.get(x["op"], x["op"])
        if op in ("blank", "is empty"):
            cond = f'{c}=""'
        elif op in ("not blank", "not empty"):
            cond = f'{c}<>""'
        elif op in ("in", "not in", "contains", "not contains", "startswith", "endswith"):
            raise ValueError("op")
        else:
            cond = f"{c}{op}{_xl_const(v)}"
        return f"IF({cond},{tree(t['then'])},{tree(t['else'])})"

    try:
        if k == "copy":
            return "=" + ref(rule["col"])
        if k == "const":
            return "=" + _xl_const(rule.get("value"))
        if k == "text":
            c = ref(rule["col"])
            f = {"upper": f"UPPER({c})", "lower": f"LOWER({c})", "strip": f"TRIM({c})",
                 "title": f"PROPER({c})", "left": f"LEFT({c},{rule.get('n')})",
                 "right": f"RIGHT({c},{rule.get('n')})",
                 "mid": f"MID({c},{rule.get('start', 1)},{rule.get('n')})", "str": f'{c}&""'}[rule["fn"]]
            return "=" + f
        if k == "datepart":
            c = ref(rule["col"])
            p = rule["part"]
            f = {"year": f"YEAR({c})", "month": f"MONTH({c})", "day": f"DAY({c})",
                 "quarter": f"ROUNDUP(MONTH({c})/3,0)", "monthname": f'TEXT({c},"mmmm")',
                 "weekday": f'TEXT({c},"dddd")', "ym": f'TEXT({c},"yyyy-mm")'}.get(p)
            return "=" + f if f else None
        if k == "arith":
            f = node(rule["expr"])
            if rule.get("round") is not None:
                f = f"ROUND({f},{int(rule['round'])})"
            if rule.get("on_error") is not None:
                f = f"IFERROR({f},{_xl_const(rule['on_error'])})"
            return "=" + f
        if k == "concat":
            return "=" + "&".join(ref(p["col"]) if "col" in p else _xl_const(p["text"])
                                  for p in rule["parts"])
        if k == "lookup":
            return "=" + lk(rule)
        if k == "cond":
            return "=" + tree(rule["tree"])
        if k == "formula":
            return rule.get("excel_formula") or None
    except (ValueError, KeyError):
        return None
    return None
