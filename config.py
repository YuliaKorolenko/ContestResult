"""
Configuration module that loads settings from local_config.py (gitignored).
Create local_config.py from local_config.example.py template.
"""
import importlib.util
import sys
from pathlib import Path
from typing import List, Optional

local_config_path = Path(__file__).parent / "local_config.py"

if not local_config_path.exists():
    raise ImportError(
        "local_config.py not found! "
        "Please copy local_config.example.py to local_config.py and fill in your secrets."
    )


spec = importlib.util.spec_from_file_location("local_config", local_config_path)
if spec is None or spec.loader is None:
    raise ImportError("Failed to load local_config.py")

local_config = importlib.util.module_from_spec(spec)
sys.modules["local_config"] = local_config
spec.loader.exec_module(local_config)


CF_API_KEY: str = local_config.CF_API_KEY
CF_API_SECRET: str = local_config.CF_API_SECRET
GOOGLE_SA_PATH: str = local_config.GOOGLE_SA_PATH
TABLE_ID: Optional[str] = local_config.TABLE_ID
TABLE_LINK: Optional[str] = local_config.TABLE_LINK
WORKSHEET_NAME: str = local_config.WORKSHEET_NAME
CONTEST_IDS: List[int] = local_config.CONTEST_IDS

# Re-export for convenience
__all__ = [
    "CF_API_KEY",
    "CF_API_SECRET",
    "GOOGLE_SA_PATH",
    "TABLE_ID",
    "TABLE_LINK",
    "WORKSHEET_NAME",
    "CONTEST_IDS",
]
