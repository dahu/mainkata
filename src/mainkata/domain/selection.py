import random

from mainkata.domain.types import VocabPair


def random_sets(
    vocab: list[VocabPair],
    set_count: int = 6,
    set_size: int = 10,
    seed: int = 42,
):
    if len(vocab) < set_size:
        raise ValueError(f"Need at least {set_size} unique items; got {len(vocab)}.")
    rng = random.Random(seed)
    return [rng.sample(vocab, set_size) for _ in range(set_count)]
