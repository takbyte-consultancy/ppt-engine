import json

from integrations.datasources.base import check_params, local_path


class JsonFileSource:
    def __init__(self, definition: dict):
        self.definition = definition

    def fetch(self, params=None) -> dict:
        params = check_params(self.definition, params)
        return json.loads(local_path(self.definition["location"], params).read_text(encoding="utf-8"))
