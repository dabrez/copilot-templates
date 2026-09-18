import pathlib

import pytest


@pytest.fixture
def repo_root():
    return pathlib.Path(__file__).resolve().parent.parent
