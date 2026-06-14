import pytest


@pytest.fixture
def labels():
    return {
        "set_prefix": "Set",
        "vocabulary_suffix": "Vocabulary",
    }
