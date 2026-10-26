"""Local preview: runs the real backend against a local folder store.

    python dev/serve.py          -> http://localhost:5181

Also writes sample files into dev/samples/ (input / mapping / expected output) to try the wizard.
"""
import os
import pathlib
import sys

from flask import Flask, Response

HERE = pathlib.Path(__file__).resolve().parent
ROOT = HERE.parent
sys.path.insert(0, str(ROOT / "python-lib"))
sys.path.insert(0, str(ROOT / "tests"))
os.environ.setdefault("REPLICATOR_LOCAL_ROOT", str(HERE / "_data"))

app = Flask(__name__)
_g = {"app": app, "__name__": "backend_dev"}
exec(compile((ROOT / "backend" / "backend.py").read_text(encoding="utf-8"), "backend.py", "exec"), _g)  # noqa: S102


def _samples():
    from helpers import make_case
    d = HERE / "samples"
    d.mkdir(exist_ok=True)
    for tag, n, seed in (("train", 80, 1), ("new", 50, 7)):
        inp, mp, out, *_ = make_case(n=n, seed=seed, split=True)
        (d / ("%s_input.xlsx" % tag)).write_bytes(inp)
        if tag == "train":
            (d / "mapping.xlsx").write_bytes(mp)
        (d / ("%s_expected_output.xlsx" % tag)).write_bytes(out)


@app.route("/dev/sample/<name>")
def _sample(name):
    for folder in ("sample_data", "samples", "sample_data/bank"):
        f = HERE / folder / pathlib.Path(name).name
        if f.exists():
            return Response(f.read_bytes(), mimetype="application/octet-stream")
    return Response("not found", 404)


@app.route("/")
def _index():
    body = (ROOT / "frontend" / "body.html").read_text(encoding="utf-8")
    css = (ROOT / "frontend" / "style.css").read_text(encoding="utf-8")
    js = (ROOT / "frontend" / "app.js").read_text(encoding="utf-8")
    return Response('<!doctype html><html><head><meta charset="utf-8">'
                    '<meta name="viewport" content="width=device-width, initial-scale=1">'
                    '<title>Category Replicator (dev)</title><style>' + css + '</style></head>'
                    '<body>' + body + '<script>' + js + '</script></body></html>', mimetype="text/html")


if __name__ == "__main__":
    _samples()
    app.run(port=5181, debug=False)
