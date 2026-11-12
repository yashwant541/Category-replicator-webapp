"""Knobs and helpers for running big data: how many worker processes, when to split a sheet, a stage timer.

Everything can be set from the environment (so a Dataiku admin needs no code change):
    REPLICATOR_WORKERS            worker processes for reading / writing (default: CPU count, at most 8; 1 = no parallelism)
    REPLICATOR_ROWS_PER_SHEET     data rows after which an output sheet continues on a new tab (default 1,000,000)
    REPLICATOR_FAST_WRITE_ROWS    output rows from which the streaming (multi-process) writer is used (default 100,000)
    REPLICATOR_PARALLEL_MIN_ROWS  below this many rows no worker processes are started (default 100,000)
"""

from __future__ import annotations

import multiprocessing as mp
import os
import time
from collections import OrderedDict
from concurrent.futures import ProcessPoolExecutor
from contextlib import contextmanager
from typing import Any, Dict, Iterator, Optional

EXCEL_MAX_ROWS = 1048576                 # rows on one Excel sheet, header rows included
CHUNK_ROWS = 25000

IN_WORKER = False                        # set inside worker processes: never start a pool from a pool


def _env_int(name: str, default: int) -> int:
    try:
        v = int(str(os.environ.get(name, "")).replace(",", "").replace("_", "").strip())
        return v if v > 0 else default
    except ValueError:
        return default


def workers() -> int:
    return max(1, _env_int("REPLICATOR_WORKERS", min(8, os.cpu_count() or 1)))


def rows_per_sheet() -> int:
    return min(_env_int("REPLICATOR_ROWS_PER_SHEET", 1000000), EXCEL_MAX_ROWS - 1)


def fast_write_rows() -> int:
    return _env_int("REPLICATOR_FAST_WRITE_ROWS", 100000)


def parallel_min_rows() -> int:
    return _env_int("REPLICATOR_PARALLEL_MIN_ROWS", 100000)


def parallel_min_bytes() -> int:
    return _env_int("REPLICATOR_PARALLEL_MIN_BYTES", 5000000)


def start_method() -> str:
    return "fork" if "fork" in mp.get_all_start_methods() else "spawn"


class _NoStage:
    def __enter__(self) -> None:
        return None

    def __exit__(self, *a: Any) -> bool:
        return False


def stage(timer: Optional["Timer"], name: str):
    """``with perf.stage(timer, 'read'):`` - times the block when there is a timer, does nothing otherwise."""
    return timer.stage(name) if timer is not None else _NoStage()


def _init_worker() -> None:
    global IN_WORKER
    IN_WORKER = True


def make_pool(n: int) -> Optional[ProcessPoolExecutor]:
    """A pool of ``n`` worker processes, or None when that is pointless or impossible (n <= 1, already in a worker)."""
    if IN_WORKER or n <= 1:
        return None
    try:
        return ProcessPoolExecutor(max_workers=n, mp_context=mp.get_context(start_method()), initializer=_init_worker)
    except Exception:  # noqa: BLE001 - no processes available here: run serially
        return None


def map_ordered(pool: Optional[ProcessPoolExecutor], fn, tasks, window: int = 0) -> Iterator[Any]:
    """fn(task) for every task, results in task order; at most ``window`` in flight so memory stays bounded."""
    if pool is None:
        for t in tasks:
            yield fn(t)
        return
    window = window or 2 * getattr(pool, "_max_workers", 2)
    pending = []
    for t in tasks:
        pending.append(pool.submit(fn, t))
        if len(pending) >= window:
            yield pending.pop(0).result()
    while pending:
        yield pending.pop(0).result()


class Timer:
    """Wall-clock time of the stages of one run, plus the figures that make it readable (rows, rows per second)."""

    def __init__(self) -> None:
        self.t0 = time.perf_counter()
        self.stages: "OrderedDict[str, float]" = OrderedDict()
        self.info: Dict[str, Any] = {}

    @contextmanager
    def stage(self, name: str) -> Iterator[None]:
        t = time.perf_counter()
        try:
            yield
        finally:
            self.stages[name] = round(self.stages.get(name, 0.0) + time.perf_counter() - t, 3)

    def report(self, rows_in: int, rows_out: int, **extra: Any) -> Dict[str, Any]:
        total = max(time.perf_counter() - self.t0, 1e-6)
        return {"seconds": round(total, 3), "stages": dict(self.stages), "rows_in": int(rows_in), "rows_out": int(rows_out),
                "rows_per_sec": int(round(max(rows_in, rows_out) / total)), **self.info, **extra}


def human_seconds(s: float) -> str:
    if s < 60:
        return "%.1f s" % s
    m, sec = divmod(int(round(s)), 60)
    return "%d min %02d s" % (m, sec) if m < 60 else "%d h %02d min" % divmod(m, 60)
