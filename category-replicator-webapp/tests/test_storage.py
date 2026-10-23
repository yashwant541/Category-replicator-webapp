import io
import os
import tempfile

from helpers import make_case
from replicator import service
from replicator.store import LocalStore, get_json, put_json
from test_api import upload, client


def test_layout_and_reload():
    c = client()
    from replicator.store import make_store
    root = os.environ["REPLICATOR_LOCAL_ROOT"]
    inp, mp, out, *_ = make_case(n=40)
    did = c.post("/api/drafts").get_json()["draft_id"]
    for n, d, h in (("in.xlsx", inp, "input"), ("map.xlsx", mp, "mapping"), ("out.xlsx", out, "output")):
        upload(c, f"/api/drafts/{did}/files", n, d, hint=h)
    sid = c.post(f"/api/drafts/{did}/learn", json={"name": "S"}).get_json()["system"]["id"]
    files = {os.path.relpath(os.path.join(b, f), root).replace("\\", "/") for b, _, fs in os.walk(root) for f in fs}
    assert {"index.json", f"systems/{sid}/system.json", f"systems/{sid}/rules.json",
            f"systems/{sid}/tables/Rates.json", f"systems/{sid}/tables/Groups.json"} <= files
    st = LocalStore(root)
    assert "tables" not in get_json(st, f"systems/{sid}/system.json")
    s = service.get_system(st, sid)
    assert len(s["tables"]) == 2 and s["rules"]
    # legacy single-file system still loads and migrates on save
    root2 = tempfile.mkdtemp()
    st2 = LocalStore(root2)
    put_json(st2, f"systems/{sid}/system.json", s)
    assert service.get_system(st2, sid)["tables"].keys() == s["tables"].keys()
    service.save_system(st2, service.get_system(st2, sid))
    assert st2.get(f"systems/{sid}/rules.json") and service.list_systems(st2)[0]["name"] == "S"
