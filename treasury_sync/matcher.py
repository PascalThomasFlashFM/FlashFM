from __future__ import annotations

from .excel_treasury import TargetRow, normalize

MIN_SUBSTRING_LENGTH = 5


def _substring_candidates(tiers_name: str, targets: list[TargetRow]) -> list[TargetRow]:
    normalized_tiers = normalize(tiers_name)
    found = []
    for t in targets:
        normalized_label = normalize(t.label)
        if len(normalized_label) < MIN_SUBSTRING_LENGTH:
            continue
        if normalized_label in normalized_tiers or normalized_tiers in normalized_label:
            found.append(t)
    return found


def resolve(
    tiers_name: str | None,
    label: str,
    amount: float,
    category_labels: list[str],
    label_index: dict[str, list[TargetRow]],
    mapping: dict,
) -> tuple[TargetRow | None, str, list[TargetRow]]:
    """Returns (target_row_or_None, reason, candidates).
    reason is one of: "ok", "ok_disambiguated", "ok_override",
    "ambigu", "non_trouve"."""

    candidates: list[TargetRow] = []

    if tiers_name:
        candidates = label_index.get(normalize(tiers_name), [])

    if not candidates and tiers_name:
        alias = mapping.get("tiers_to_label", {}).get(tiers_name)
        if alias:
            candidates = label_index.get(normalize(alias), [])

    if not candidates and tiers_name:
        all_targets = [t for targets in label_index.values() for t in targets]
        candidates = _substring_candidates(tiers_name, all_targets)

    if not candidates:
        for keyword, alias in mapping.get("keyword_to_label", {}).items():
            if keyword.upper() in label.upper():
                candidates = label_index.get(normalize(alias), [])
                if candidates:
                    break

    if len(candidates) == 1:
        return candidates[0], "ok", candidates

    if len(candidates) > 1:
        # Try to disambiguate using the Pennylane category vs. the bloc name.
        normalized_categories = [normalize(c) for c in category_labels]
        bloc_matches = [
            c for c in candidates if any(
                normalize(c.bloc) in cat or cat in normalize(c.bloc)
                for cat in normalized_categories
            )
        ]
        if len(bloc_matches) == 1:
            return bloc_matches[0], "ok_disambiguated", candidates

        if tiers_name:
            for override in mapping.get("ambiguous_overrides", []):
                if override.get("tiers", "").upper() != tiers_name.upper():
                    continue
                if override["min"] <= abs(amount) <= override["max"]:
                    match = next(
                        (c for c in candidates if normalize(c.label) == normalize(override["target"])),
                        None,
                    )
                    if match:
                        return match, "ok_override", candidates

        return None, "ambigu", candidates

    return None, "non_trouve", []
