import pathlib, subprocess, sys

ROOT = pathlib.Path(__file__).resolve().parents[1]


def test_single_file_backend_runs_without_python_lib():
    subprocess.check_call([sys.executable, str(ROOT / "build_single.py")])
    code = f"""
import sys, os, tempfile, io, importlib.util
sys.path = [p for p in sys.path if 'python-lib' not in p]
sys.path.insert(0, {str(ROOT / 'tests')!r})
os.environ['REPLICATOR_LOCAL_ROOT'] = tempfile.mkdtemp()
from flask import Flask
from helpers import make_case
app = Flask('x'); g = {{'app': app, '__name__': 'b'}}
exec(compile(open({str(ROOT / 'backend_single.py')!r}, encoding='utf-8').read(), 'backend_single.py', 'exec'), g)
c = app.test_client()
inp, mp, out, *_ = make_case(n=60)
did = c.post('/api/drafts').get_json()['draft_id']
for n, d, h in (('in.xlsx', inp, 'input'), ('map.xlsx', mp, 'mapping'), ('out.xlsx', out, 'output')):
    r = c.post(f'/api/drafts/{{did}}/files', data={{'files': (io.BytesIO(d), n), 'hint': h}}, content_type='multipart/form-data').get_json()
    assert r['success'], r
r = c.post(f'/api/drafts/{{did}}/learn', json={{'name': 's'}}).get_json()
assert r['success'] and r['system']['sheets'], r
print('SINGLE OK')
"""
    res = subprocess.run([sys.executable, "-c", code], capture_output=True, text=True)
    assert "SINGLE OK" in res.stdout, res.stderr[-1500:]
