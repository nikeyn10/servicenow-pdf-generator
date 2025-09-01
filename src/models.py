from pydantic import BaseModel
from typing import List, Optional

class Asset(BaseModel):
    id: str
    name: str
    file_extension: str
    public_url: Optional[str]
    url: Optional[str]
    size: Optional[int] = None

class TicketRow(BaseModel):
    item_id: str
    item_name: str
    open_date: str
    close_date: Optional[str]
    attachments: List[Asset]

class Item(BaseModel):
    id: str
    name: str
    assets: List[Asset]
    open_date: str
    close_date: Optional[str]
    status_label: str
