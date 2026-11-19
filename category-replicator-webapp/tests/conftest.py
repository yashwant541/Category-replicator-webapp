import pathlib
import sys

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parents[1] / "python-lib"))


# The run / inspect-input endpoints start a background job and return its id. The tests want the finished answer, so the
# Flask test client waits for the job and presents it the way the old synchronous endpoints did.
import time as _time

import pytest as _pytest


class _Done:
    def __init__(self, status, payload):
        self.status_code, self._p = status, payload

    def get_json(self, *a, **k):
        return self._p


@_pytest.fixture(autouse=True)
def _wait_for_jobs(monkeypatch):
    from flask.testing import FlaskClient
    orig = FlaskClient.post

    def post(self, path, *a, **k):
        r = orig(self, path, *a, **k)
        if not (path.endswith("/run") or path.endswith("/inspect-input")) or r.status_code != 200:
            return r
        j = r.get_json() or {}
        if not j.get("job_id"):
            return r
        for _ in range(6000):
            job = self.get("/api/jobs/" + j["job_id"]).get_json()["job"]
            if job["state"] == "done":
                res = job["result"]
                return _Done(200, {"success": True, **(res if path.endswith("inspect-input") else {"run": res})})
            if job["state"] == "error":
                return _Done(400, {"success": False, "error": job["error"], "details": job.get("details")})
            _time.sleep(0.05)
        raise AssertionError("job did not finish")
    monkeypatch.setattr(FlaskClient, "post", post)
