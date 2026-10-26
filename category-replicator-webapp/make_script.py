"""Learn from example files and write a standalone Dataiku recipe script.

    python make_script.py --input in.xlsx --mapping map.xlsx --output expected.xlsx -o recipe.py \
        [--name Retail] [--data-folder ID] [--mapping-folder ID] [--output-folder ID] [--template-folder ID]

--input/--mapping/--output may be repeated. The output workbook may itself contain formulas and mapping sheets.
"""
import argparse
import pathlib
import sys
import tempfile

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent / "python-lib"))
from replicator import codegen, service  # noqa: E402
from replicator.store import LocalStore  # noqa: E402


def main():
    ap = argparse.ArgumentParser()
    for k in ("input", "mapping", "output"):
        ap.add_argument("--" + k, action="append", default=[], help=k + " example file")
    ap.add_argument("-o", "--out", default="recipe.py")
    ap.add_argument("--name", default="Category")
    ap.add_argument("--data-folder", default="INPUT_FOLDER_ID")
    ap.add_argument("--mapping-folder", default="")
    ap.add_argument("--output-folder", default="OUTPUT_FOLDER_ID")
    ap.add_argument("--template-folder", default="")
    a = ap.parse_args()
    st = LocalStore(tempfile.mkdtemp())
    did = service.new_draft(st)
    for hint in ("input", "mapping", "output"):
        for f in getattr(a, hint):
            p = pathlib.Path(f)
            service.add_draft_files(st, did, [(p.name, p.read_bytes())], hint)
    s = service.create_system(st, did, {"name": a.name})
    bad = [(r["name"], r["status"]) for r in s["rules"] if r["status"] != "ok"]
    print("replay match on the example: %s" % s["training"].get("replay_match"))
    if bad:
        print("columns NOT fully learned (review):", bad)
    pathlib.Path(a.out).write_text(codegen.generate_script(
        s, data_folder=a.data_folder, mapping_folder=a.mapping_folder, output_folder=a.output_folder,
        template_folder=a.template_folder, category=a.name), encoding="utf-8")
    print("wrote", a.out)


main()
