"""Persistence: a Dataiku managed folder (production) or a local directory (dev / tests).

Folder id comes from, in order: the ``REPLICATOR_FOLDER`` project variable, the env var of the
same name, or the ``FOLDER_FALLBACK`` constant below.  Without Dataiku, everything lives under
``$REPLICATOR_LOCAL_ROOT`` (default ``./_replicator_data``).
"""

from __future__ import annotations

import io
import json
import os
from typing import Any, List, Optional

FOLDER_FALLBACK = ""  # hard-wire a managed folder id here if you prefer

try:  # pragma: no cover - only inside Dataiku
    import dataiku
    _HAS_DATAIKU = True
except Exception:  # noqa: BLE001
    dataiku = None
    _HAS_DATAIKU = False


class LocalStore:
    def __init__(self, root: str):
        self.root = os.path.abspath(root)
        os.makedirs(self.root, exist_ok=True)

    def _p(self, path: str) -> str:
        full = os.path.abspath(os.path.join(self.root, path.lstrip("/")))
        if not full.startswith(self.root):
            raise ValueError("bad path")
        return full

    def get(self, path: str) -> Optional[bytes]:
        try:
            with open(self._p(path), "rb") as fh:
                return fh.read()
        except FileNotFoundError:
            return None

    def put(self, path: str, data: bytes) -> None:
        full = self._p(path)
        os.makedirs(os.path.dirname(full), exist_ok=True)
        with open(full, "wb") as fh:
            fh.write(data)

    def list(self, prefix: str = "") -> List[str]:
        out = []
        for base, _, files in os.walk(self.root):
            for f in files:
                rel = "/" + os.path.relpath(os.path.join(base, f), self.root).replace("\\", "/")
                if rel.startswith("/" + prefix.lstrip("/")):
                    out.append(rel)
        return sorted(out)

    def delete(self, path: str) -> None:
        try:
            os.remove(self._p(path))
        except FileNotFoundError:
            pass


class DataikuStore:  # pragma: no cover - only inside Dataiku
    def __init__(self, folder_id: str):
        self.f = dataiku.Folder(folder_id)

    def get(self, path: str) -> Optional[bytes]:
        try:
            with self.f.get_download_stream("/" + path.lstrip("/")) as s:
                return s.read()
        except Exception:  # noqa: BLE001
            return None

    def put(self, path: str, data: bytes) -> None:
        self.f.upload_stream("/" + path.lstrip("/"), io.BytesIO(data))

    def list(self, prefix: str = "") -> List[str]:
        return sorted(p for p in self.f.list_paths_in_partition() if p.startswith("/" + prefix.lstrip("/")))

    def delete(self, path: str) -> None:
        try:
            self.f.delete_path("/" + path.lstrip("/"))
        except Exception:  # noqa: BLE001
            pass


def _folder_id() -> str:
    fid = os.environ.get("REPLICATOR_FOLDER", "")
    if _HAS_DATAIKU:
        try:
            fid = dataiku.get_custom_variables().get("REPLICATOR_FOLDER") or fid
        except Exception:  # noqa: BLE001
            pass
    return fid or FOLDER_FALLBACK


def make_store() -> Any:
    fid = _folder_id()
    if _HAS_DATAIKU and fid:
        return DataikuStore(fid)
    return LocalStore(os.environ.get("REPLICATOR_LOCAL_ROOT", "_replicator_data"))


def put_json(store: Any, path: str, obj: Any) -> None:
    store.put(path, json.dumps(obj, ensure_ascii=False, default=str).encode("utf-8"))


def get_json(store: Any, path: str) -> Optional[Any]:
    raw = store.get(path)
    return json.loads(raw.decode("utf-8")) if raw else None


def delete_prefix(store: Any, prefix: str) -> None:
    for p in store.list(prefix):
        store.delete(p)
