import pathlib, sys, tempfile
import conftest
from replicator import service
from replicator.store import LocalStore

D = pathlib.Path(__file__).resolve().parents[1] / "dev" / "sample_data"


def build():
    st = LocalStore(tempfile.mkdtemp())
    did = service.new_draft(st)
    for f, hint in (("1_input_exposures.xlsx", "input"), ("2_mapping_tables.xlsx", "mapping"), ("3_expected_output.xlsx", "output")):
        d = service.add_draft_files(st, did, [(f, (D / f).read_bytes())], hint)
    for f in d["files"]:
        print(f["name"], {s["sheet"]: s["role"] for s in f["sheets"]})
    return st, service.create_system(st, did, {"name": "dummy"})


if __name__ == "__main__":
    st, s = build()
    print("warnings:", s["warnings"])
    print("columns:", s["output"]["columns"])
    for r in s["rules"]:
        print("%-18s %-8s %-10s %s | %s" % (r["name"], r["kind"], r["status"], r["match"], (r["description"] or "")[:90]))
    print("split:", s["filters"]["split"]["enabled"], s["filters"]["split"].get("column"), "replay", s["training"].get("replay_match"))
