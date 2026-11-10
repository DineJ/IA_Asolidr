"""Shared pytest fixtures for the IA_Asolidr extraction tests."""
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))


@pytest.fixture(scope="session")
def project_root():
    return ROOT


@pytest.fixture(scope="session")
def raw_dir(project_root):
    """Real Excel inputs, organized as data/raw/YYYY/."""
    return project_root / "data" / "raw"
