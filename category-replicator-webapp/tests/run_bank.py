import pathlib, tempfile, sys
import conftest
from replicator import service
from replicator.store import LocalStore

DL = pathlib.Path("C:/Users/yavnv/Downloads")


def build(names=None):
    st = LocalStore(tempfile.mkdtemp())
    did = service.new_draft(st)
    for f, hint in (("bank_financials_source_data.xlsx", "input"), ("bank_financials_mapping.xlsx", "mapping"),
                    ("bank_financials_outputs.xlsx", "output")):
        d = service.add_draft_files(st, did, [(f, (DL / f).read_bytes())], hint)
    for f in d["files"]:
        print(f["name"], {s["sheet"]: s["role"] for s in f["sheets"]})
    return st, did


if __name__ == "__main__":
    st, did = build()
    try:
        s = service.create_system(st, did, {"name": "bank"})
    except Exception as e:
        print("FAILED:", type(e).__name__, e)
        sys.exit()
    print("warnings:", s["warnings"])
    for r in s["rules"]:
        print("%-22s %-8s %-10s %s | %s" % (r["name"], r["kind"], r["status"], r["match"], (r["description"] or "")[:80]))
    print("split:", s["filters"]["split"], "replay", s["training"].get("replay_match"))
