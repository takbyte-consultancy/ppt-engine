"""Central configuration, read from environment / .env file."""
import os
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()
ROOT = Path(__file__).parent

# --- LLM ---
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY", "")
OPENAI_MODEL = os.getenv("OPENAI_MODEL", "gpt-4.1")
# For reasoning models: low / medium / high. Leave empty for non-reasoning models (e.g. gpt-4.1).
OPENAI_REASONING_EFFORT = os.getenv("OPENAI_REASONING_EFFORT", "").strip()

# --- JIRA ---
JIRA_BASE_URL = os.getenv("JIRA_BASE_URL", "").rstrip("/")
JIRA_AUTH_TYPE = os.getenv("JIRA_AUTH_TYPE", "basic")
JIRA_EMAIL = os.getenv("JIRA_EMAIL", "")
JIRA_API_TOKEN = os.getenv("JIRA_API_TOKEN", "")
# Custom field ids holding these values (optional; description "Label: value" lines are the fallback)
JIRA_FIELD_IDS = {
    "asset_type": os.getenv("JIRA_FIELD_ASSET_TYPE", ""),
    "acronym": os.getenv("JIRA_FIELD_ACRONYM", ""),
    "datasource_id": os.getenv("JIRA_FIELD_DATASOURCE", ""),
}

# --- Template key: which fields make a template unique (order matters) ---
TEMPLATE_KEY_FIELDS = [f.strip() for f in os.getenv("TEMPLATE_KEY_FIELDS", "asset_type,acronym").split(",")]

# --- Workflow ---
MAX_MAPPING_ATTEMPTS = int(os.getenv("MAX_MAPPING_ATTEMPTS", "3"))

# --- Storage (JSON files today; swapped for DB repositories later) ---
DATA_DIR = ROOT / os.getenv("DATA_DIR", "data")
TEMPLATES_DIR = ROOT / os.getenv("TEMPLATES_DIR", "templates")
EXHIBITS_DIR = ROOT / os.getenv("EXHIBITS_DIR", "exhibits")
PITCHBOOKS_DIR = ROOT / os.getenv("PITCHBOOKS_DIR", "pitchbooks")
WORK_DIR = ROOT / "work"
MOCK_DIR = ROOT / "mock"
for _d in (DATA_DIR, TEMPLATES_DIR, EXHIBITS_DIR, PITCHBOOKS_DIR, WORK_DIR):
    _d.mkdir(exist_ok=True)
