import httpx
import os
import tenacity
import yaml
from src.queries import GET_STATUS_COLUMN, GET_ITEMS_PAGE, NEXT_ITEMS_PAGE
from src.log import log_event

class MondayClient:
    def __init__(self, config):
        self.token = os.getenv(config['api']['token_env'])
        self.api_url = config['api']['api_url']
        self.headers = {
            "Authorization": self.token,
            "Content-Type": "application/json",
        }
        if 'api_version' in config['api']:
            self.headers["API-Version"] = config['api']['api_version']

    @tenacity.retry(stop=tenacity.stop_after_attempt(5), wait=tenacity.wait_exponential(), retry=tenacity.retry_if_exception_type(httpx.HTTPStatusError))
    def graphql(self, query, variables):
        resp = httpx.post(self.api_url, headers=self.headers, json={"query": query, "variables": variables}, timeout=30.0)
        resp.raise_for_status()
        data = resp.json()
        if 'errors' in data:
            raise Exception(data['errors'])
        return data['data']

    def get_status_column(self, board_id):
        return self.graphql(GET_STATUS_COLUMN, {"boardId": [board_id]})

    def get_items_page(self, board_id, from_date, to_date, status_idx, limit=500):
        return self.graphql(GET_ITEMS_PAGE, {
            "boardId": [board_id],
            "limit": limit
        })

    def next_items_page(self, cursor, limit=500):
        return self.graphql(NEXT_ITEMS_PAGE, {"cursor": cursor, "limit": limit})
