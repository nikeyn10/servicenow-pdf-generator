import os
import httpx
import shutil
from src.log import log_event
from pathlib import Path

def sanitize_filename(name):
    return "".join(c if c.isalnum() or c in "-_" else "_" for c in name)

def download_asset(asset, out_dir):
    filename = f"{asset.id}-{sanitize_filename(asset.name)}.{asset.file_extension}"
    path = Path(out_dir) / filename
    if path.exists():
        return str(path)
    try:
        resp = httpx.get(asset.public_url, timeout=60)
        resp.raise_for_status()
        with open(path, "wb") as f:
            f.write(resp.content)
        log_event(asset_id=asset.id, action="download", status="success")
        return str(path)
    except Exception as e:
        log_event(asset_id=asset.id, action="download", status="fail", warning=str(e))
        return None

def dedupe_assets(assets):
    seen = set()
    unique = []
    for asset in assets:
        key = (asset.id, asset.name)
        if key not in seen:
            seen.add(key)
            unique.append(asset)
    return unique
