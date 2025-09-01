import calendar
from datetime import datetime
import yaml

# Month range calculation
def get_month_range(month: str):
    year, mon = map(int, month.split('-'))
    first_day = f"{year:04d}-{mon:02d}-01"
    last_day = f"{year:04d}-{mon:02d}-{calendar.monthrange(year, mon)[1]:02d}"
    return first_day, last_day

# Status label → index
def get_status_index(settings_str: str, required_label: str) -> int:
    import json
    settings = json.loads(settings_str)
    labels = settings.get('labels', {})
    for idx, label in labels.items():
        if label.lower() == required_label.lower():
            return int(idx)
    raise ValueError(f"Status label '{required_label}' not found.")
