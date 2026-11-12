"""Time a heavy run: python dev/benchmark.py [rows] [workers]

Builds the system from dev/sample_data (dummy data), makes a big multi-sheet input workbook with ``rows`` rows (written with the
fast writer), runs it through the engine and prints how long each stage took.  Nothing is stored outside a temp folder.
"""
import io
import pathlib
import shutil
import sys
import tempfile
import time
import zipfile

ROOT = pathlib.Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "python-lib"))

from replicator import fastxlsx, perf, service              # noqa: E402
from replicator.store import LocalStore                      # noqa: E402
from replicator.tables import read_workbook                   # noqa: E402

D = ROOT / "dev" / "sample_data"


def big_input(n_rows: int) -> bytes:
    sheets = read_workbook((D / "4_new_input_exposures.xlsx").read_bytes(), "new.xlsx")
    per = max(1, n_rows // len(sheets))
    specs = []
    for sh in sheets:
        df = sh["df"]
        base = [list(r) for r in df.itertuples(index=False, name=None)]
        idc = 0                                               # first column is the loan id: keep every id unique
        names = list(df.columns)

        def rows(base=base, names=names, per=per):
            yield 1, [(j, h, 0) for j, h in enumerate(names)]
            for i in range(per):
                r = base[i % len(base)]
                yield i + 2, [(0, "L%08d" % i, 0)] + [(j, v, 0) for j, v in enumerate(r) if j and v is not None]
        specs.append({"name": sh["sheet"], "static": rows()})
    return fastxlsx.write_package(specs, fastxlsx.Styles())


def main() -> None:
    n = int(sys.argv[1]) if len(sys.argv) > 1 else 200000
    workers = int(sys.argv[2]) if len(sys.argv) > 2 else None
    tmp = tempfile.mkdtemp()
    try:
        st = LocalStore(tmp)
        did = service.new_draft(st)
        for name, role in (("1_input_exposures.xlsx", "input"), ("2_mapping_tables.xlsx", "mapping"), ("3_expected_output.xlsx", "output")):
            service.add_draft_files(st, did, [(name, (D / name).read_bytes())], role)
        s = service.create_system(st, did, {"name": "bench"})
        t = time.perf_counter()
        data = big_input(n)
        print("input workbook: %s rows, %.1f MB, built in %.1f s" % (format(n, ","), len(data) / 1048576, time.perf_counter() - t))
        opts = {"workers": workers} if workers else {}
        r = service.run_system(st, s["id"], [("big.xlsx", data)], options=opts)
        t = r["timing"]
        print("rows in %s -> out %s | total %s | %s rows/s | writer %s | workers %s | tabs %s | file %.1f MB" % (
            format(t["rows_in"], ","), format(t["rows_out"], ","), perf.human_seconds(t["seconds"]), format(t["rows_per_sec"], ","),
            t["writer"], t["workers"], t["tabs"], t["output_mb"]))
        print("stages:", {k: round(v, 1) for k, v in t["stages"].items()})
        print("sheets:", r["sheets"], "| split:", r.get("split"))
        out, _ = service.get_run_file(st, s["id"], r["run_id"])
        with zipfile.ZipFile(io.BytesIO(out)) as z:
            print("output package ok:", z.testzip() is None, "| parts:", len(z.namelist()))
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


if __name__ == "__main__":
    main()
