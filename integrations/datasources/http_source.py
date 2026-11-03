import os

import requests

from integrations.datasources.base import check_params, resolve_location


class HttpSource:
    def __init__(self, definition: dict):
        self.definition = definition

    def fetch(self, params=None) -> dict:
        params = check_params(self.definition, params)
        headers = {}
        if self.definition.get("auth") == "bearer":
            headers["Authorization"] = f"Bearer {os.getenv(self.definition.get('token_env') or 'DATA_API_TOKEN', '')}"
        r = requests.get(resolve_location(self.definition["location"], params), headers=headers,
                         timeout=self.definition.get("timeout", 60))
        r.raise_for_status()
        return r.json()
