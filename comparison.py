"""Deterministic paragraph alignment with whitespace-tolerant word differences."""

from __future__ import annotations

import re
import unicodedata
from collections import Counter, defaultdict
from dataclasses import dataclass
from difflib import SequenceMatcher
from itertools import groupby

_LIGATURES = str.maketrans({
    "ﬀ": "ff", "ﬁ": "fi", "ﬂ": "fl", "ﬃ": "ffi",
    "ﬄ": "ffl", "ﬅ": "st", "ﬆ": "st",
})
_WORD_PATTERN = re.compile(r"\w+|[^\w\s]")


@dataclass(frozen=True)
class Block:
    text: str
    kind: str = "paragraph"
    page: int | None = None


@dataclass(frozen=True)
class Span:
    text: str
    changed: bool = False


@dataclass(frozen=True)
class Change:
    kind: str
    original: tuple[Block, ...]
    modified: tuple[Block, ...]
    before: tuple[Span, ...]
    after: tuple[Span, ...]
    move: int | None = None


@dataclass(frozen=True)
class Comparison:
    original_name: str
    modified_name: str
    rows: tuple[Change, ...]

    @property
    def summary(self) -> dict[str, int]:
        counts = Counter(row.kind for row in self.rows)
        return {
            "inserted": counts["inserted"],
            "deleted": counts["deleted"],
            "modified": counts["modified"],
            "moved": counts["moved"] // 2,
        }


def normalize(text: str) -> str:
    text = re.sub(r"\u00ad[ \t]*\n[ \t]*", "", text)
    text = unicodedata.normalize("NFC", text).replace("\u00ad", "").translate(_LIGATURES)
    return " ".join(text.split())


def block_key(block: Block) -> tuple[str, str]:
    return block.kind, normalize(block.text)


def _word_diff(left: str, right: str) -> tuple[tuple[Span, ...], tuple[Span, ...]]:
    old, new = list(_WORD_PATTERN.finditer(left)), list(_WORD_PATTERN.finditer(right))
    matcher = SequenceMatcher(None, [m[0] for m in old], [m[0] for m in new], autojunk=False)

    def spans(text, tokens, changed_indices):
        result = []
        cursor = 0
        for start, end in changed_indices:
            if start == end:
                continue
            left_edge, right_edge = tokens[start].start(), tokens[end - 1].end()
            if cursor < left_edge:
                result.append(Span(text[cursor:left_edge]))
            result.append(Span(text[left_edge:right_edge], True))
            cursor = right_edge
        if cursor < len(text):
            result.append(Span(text[cursor:]))
        return tuple(result)

    changes = [opcode for opcode in matcher.get_opcodes() if opcode[0] != "equal"]
    return (
        spans(left, old, [(i, j) for _, i, j, _, _ in changes]),
        spans(right, new, [(i, j) for _, _, _, i, j in changes]),
    )


def _row(kind, original, modified, move=None) -> Change:
    left = "\n\n".join(normalize(block.text) for block in original)
    right = "\n\n".join(normalize(block.text) for block in modified)
    if kind == "modified":
        before, after = _word_diff(left, right)
    else:
        before = (Span(left, kind == "deleted"),) if left else ()
        after = (Span(right, kind == "inserted"),) if right else ()
    return Change(kind, tuple(original), tuple(modified), before, after, move)


def _structure(blocks):
    # PDF line wrapping and page breaks can split or merge body paragraphs.
    groups = []
    for kind, items in groupby(blocks, key=lambda block: block.kind):
        keys = [normalize(block.text) for block in items]
        groups.append((kind, (" ".join(keys),) if kind == "paragraph" else tuple(keys)))
    return groups


def _replacement(left, right):
    if _structure(left) == _structure(right):
        yield _row("equal", left, right)
    elif len(left) == len(right):
        yield from (_row("modified", [a], [b]) for a, b in zip(left, right))
    else:
        yield _row("modified", left, right)


def _similar_keys(old_keys, new_keys):
    """Find unambiguous edited paragraphs that can anchor the alignment."""
    old_set, new_set = set(old_keys), set(new_keys)
    old_candidates = [key for key, count in Counter(old_keys).items() if count == 1 and key not in new_set]
    new_candidates = [key for key, count in Counter(new_keys).items() if count == 1 and key not in old_set]
    index = defaultdict(list)
    for key in new_candidates:
        kind, text = key
        for word in set(_WORD_PATTERN.findall(text.casefold())):
            if word.isalnum():
                index[kind, word].append(key)

    matches, reverse = defaultdict(list), defaultdict(list)
    for key in old_candidates:
        kind, text = key
        candidates = set()
        for word in set(_WORD_PATTERN.findall(text.casefold())):
            entries = index[kind, word]
            if len(entries) <= 24:
                candidates.update(entries)
        for other in candidates:
            matcher = SequenceMatcher(None, text, other[1])
            if matcher.quick_ratio() >= 0.8 and matcher.ratio() >= 0.8:
                matches[key].append(other)
                reverse[other].append(key)
    return {
        candidates[0]: key for key, candidates in matches.items()
        if len(candidates) == 1 and len(reverse[candidates[0]]) == 1
    }


def compare(
    original: list[Block], modified: list[Block], original_name: str, modified_name: str
) -> Comparison:
    old_keys, new_keys = list(map(block_key, original)), list(map(block_key, modified))
    aliases = _similar_keys(old_keys, new_keys)
    aligned_new_keys = [aliases.get(key, key) for key in new_keys]
    opcodes = SequenceMatcher(None, old_keys, aligned_new_keys, autojunk=False).get_opcodes()
    removed, added = Counter(), Counter()
    for tag, i, j, k, l in opcodes:
        if tag == "delete":
            removed.update(old_keys[i:j])
        elif tag == "insert":
            added.update(new_keys[k:l])
    old_counts, new_counts = Counter(old_keys), Counter(new_keys)
    moves = {
        key: idx
        for idx, key in enumerate(
            (key for key in removed if removed[key] == added[key] == old_counts[key] == new_counts[key] == 1),
            start=1,
        )
    }

    rows = []
    for tag, i, j, k, l in opcodes:
        left, right = original[i:j], modified[k:l]
        if tag == "equal":
            rows.extend(
                _row("equal" if block_key(a) == block_key(b) else "modified", [a], [b])
                for a, b in zip(left, right)
            )
        elif tag == "replace":
            rows.extend(_replacement(left, right))
        elif tag == "delete":
            for block in left:
                move = moves.get(block_key(block))
                rows.append(_row("moved" if move else "deleted", [block], [], move))
        else:
            for block in right:
                move = moves.get(block_key(block))
                rows.append(_row("moved" if move else "inserted", [], [block], move))
    return Comparison(original_name, modified_name, tuple(rows))
