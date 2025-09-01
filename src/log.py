import json
import time

def log_event(item_id=None, asset_id=None, action=None, status=None, duration_ms=None, warning=None):
    event = {
        "timestamp": time.time(),
        "item_id": item_id,
        "asset_id": asset_id,
        "action": action,
        "status": status,
        "duration_ms": duration_ms,
        "warning": warning,
    }
    print(json.dumps({k: v for k, v in event.items() if v is not None}))
