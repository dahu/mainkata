import pytest

from mainkata.domain.selection import random_sets
from mainkata.domain.types import VocabPair


def make_vocab(n: int) -> list[VocabPair]:
    return [(f"term-{i}", f"definition-{i}") for i in range(n)]


def test_random_sets_raises_if_vocab_too_small() -> None:
    vocab = make_vocab(5)

    with pytest.raises(
        ValueError,
        match="Need at least 10 unique items; got 5.",
    ):
        random_sets(vocab, set_count=1, set_size=10)


def test_random_sets_respects_set_count_and_size() -> None:
    vocab = make_vocab(20)

    sets = random_sets(vocab, set_count=3, set_size=7, seed=42)

    assert len(sets) == 3
    assert all(len(s) == 7 for s in sets)


def test_random_sets_is_deterministic_with_same_seed() -> None:
    vocab = make_vocab(30)

    sets_a = random_sets(vocab, set_count=4, set_size=6, seed=123)
    sets_b = random_sets(vocab, set_count=4, set_size=6, seed=123)

    assert sets_a == sets_b


def test_random_sets_differs_with_different_seeds() -> None:
    vocab = make_vocab(30)

    sets_a = random_sets(vocab, set_count=2, set_size=6, seed=1)
    sets_b = random_sets(vocab, set_count=2, set_size=6, seed=2)

    # They might coincidentally match, but with random.sample over this size,
    # it's overwhelmingly likely they differ; we assert inequality as a sanity check.
    assert sets_a != sets_b
