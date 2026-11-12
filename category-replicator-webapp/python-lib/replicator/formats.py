"""The run's output in the format the input came in, next to the .xlsx (which is always produced).

  .xlsm .xltx .xltm       the same workbook, saved as a macro-enabled / template workbook (no macros are involved).
  .xls .xlsb .ods .parquet  cannot be written here (no writer in the Python environment): the xlsx only, and a note says so.
"""

from __future__ import annotations

import io
import zipfile
from typing import Any, Dict, List, Optional, Tuple

from .tables import TEXT_EXT, file_ext

_BOOK_TYPES = {
    ".xlsm": "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
    ".xltx": "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml",
    ".xltm": "application/vnd.ms-excel.template.macroEnabled.main+xml",
}
_PLAIN = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
_MIME = {".xlsm": "application/vnd.ms-excel.sheet.macroEnabled.12", ".xltx": "application/vnd.openxmlformats-officedocument.spreadsheetml.template",
         ".xltm": "application/vnd.ms-excel.template.macroEnabled.12"}


def retype_workbook(xlsx: bytes, ext: str) -> bytes:
    """The same workbook with the content type of .xlsm / .xltx / .xltm (only the package manifest changes)."""
    out = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(xlsx)) as zin, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED, allowZip64=True) as zout:
        for info in zin.infolist():
            if info.filename == "[Content_Types].xml":
                text = zin.read(info).decode("utf-8")
                if _PLAIN not in text:
                    raise ValueError("unexpected workbook manifest")
                zout.writestr(info.filename, text.replace(_PLAIN, _BOOK_TYPES[ext], 1))
            else:
                with zin.open(info) as src, zout.open(info.filename, "w", force_zip64=True) as dst:
                    while True:
                        blk = src.read(1 << 20)
                        if not blk:
                            break
                        dst.write(blk)
    return out.getvalue()


def original_outputs(sheets: Dict[str, Dict[str, Any]], parts: Dict[str, Tuple[str, int, int]], xlsx: bytes,
                     used: List[Dict[str, Any]], base: str) -> Tuple[List[Dict[str, Any]], List[str]]:
    """The output in each distinct source format other than .xlsx -> ([{format, filename, data, mime, note}], notes)."""
    out: List[Dict[str, Any]] = []
    notes: List[str] = []
    seen: List[str] = []
    for u in used:
        ext = file_ext(u["file"])
        gz = u["file"].lower().endswith(".gz")
        key = ext + (".gz" if gz else "")
        if ext == ".xlsx" or key in seen:
            continue
        seen.append(key)
        if ext in TEXT_EXT:
            continue                                     # text files are input only: the .xlsx is the output
        if ext in _BOOK_TYPES:
            out.append({"format": ext, "filename": base + ext, "mime": _MIME[ext], "data": retype_workbook(xlsx, ext),
                        "note": "The same workbook saved as %s (no macros)." % ext})
        else:
            notes.append("The input was a %s file; this server cannot write that format, so only the .xlsx was produced "
                         "(open it in Excel and use Save As %s)." % (ext, ext))
    for o in out:
        notes.append("Also produced in the input's format (%s): %s" % (o["format"], o["note"]))
    return out, notes
