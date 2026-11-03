"""Thin JIRA REST client (Cloud: email + API token, Server/DC: bearer PAT)."""
from __future__ import annotations

import requests


class JiraClient:
    def __init__(self, base_url: str, auth_type: str = "basic", email: str = "", token: str = "",
                 timeout: int = 60):
        if not base_url:
            raise ValueError("JIRA_BASE_URL is not configured")
        self.base = base_url.rstrip("/")
        self.timeout = timeout
        self.s = requests.Session()
        self.s.headers["Accept"] = "application/json"
        if auth_type == "bearer":
            self.s.headers["Authorization"] = f"Bearer {token}"
        else:
            self.s.auth = (email, token)

    def get_issue(self, key: str, fields: str = "*all") -> dict:
        r = self.s.get(f"{self.base}/rest/api/2/issue/{key}", params={"fields": fields},
                       timeout=self.timeout)
        r.raise_for_status()
        return r.json()

    def download(self, url: str) -> bytes:
        r = self.s.get(url, timeout=self.timeout * 2)
        r.raise_for_status()
        return r.content
