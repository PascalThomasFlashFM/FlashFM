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


def _filter_by_section(candidates: list[TargetRow], amount: float) -> list[TargetRow]:
    """A transaction can only land on a "recette" row if its amount is
    positive, and on a "dépense" row if negative — this rules out
    accidental cross-section matches (e.g. a revenue transaction whose label
    happens to contain the name of an unrelated expense row)."""
    is_recette = amount > 0
    filtered = [c for c in candidates if c.is_recette == is_recette]
    return filtered if filtered else candidates


def should_ignore(label: str, category_labels: list[str], mapping: dict) -> bool:
    """Some Pennylane transactions represent the same money moving twice
    (e.g. an individual Stripe charge notification, its batched payout, and
    the internal transfer between the Stripe sub-account and the real bank
    account) and must be excluded entirely rather than matched to a row,
    otherwise the same revenue/expense gets counted more than once."""
    normalized_categories = [normalize(c) for c in category_labels]
    for category in mapping.get("ignore_categories", []):
        if normalize(category) in normalized_categories:
            return True
    for keyword in mapping.get("ignore_keywords", []):
        if keyword.upper() in label.upper():
            return True
    return False


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

    all_targets = [t for targets in label_index.values() for t in targets]

    # High-confidence, explicit row assignments (e.g. distinguishing several
    # rows that share the exact same label, like three "URSSAF" lines told
    # apart by the payment reference in the bank label) take priority over
    # the generic cascade below.
    for rule in mapping.get("row_overrides", []):
        if rule["keyword"].upper() not in label.upper():
            continue
        if "min" in rule and abs(amount) < rule["min"]:
            continue
        if "max" in rule and abs(amount) > rule["max"]:
            continue
        match = next((t for t in all_targets if t.row == rule["target_row"]), None)
        if match:
            return match, "ok_override", [match]

    candidates: list[TargetRow] = []

    if tiers_name:
        candidates = label_index.get(normalize(tiers_name), [])

    if not candidates and tiers_name:
        alias = mapping.get("tiers_to_label", {}).get(tiers_name)
        if alias:
            candidates = label_index.get(normalize(alias), [])

    if not candidates and tiers_name:
        candidates = _substring_candidates(tiers_name, all_targets)

    if not candidates:
        for keyword, alias in mapping.get("keyword_to_label", {}).items():
            if keyword.upper() in label.upper():
                candidates = label_index.get(normalize(alias), [])
                if candidates:
                    break

    if not candidates:
        normalized_categories = [normalize(c) for c in category_labels]
        for category, alias in mapping.get("category_to_label", {}).items():
            if normalize(category) in normalized_categories:
                candidates = label_index.get(normalize(alias), [])
                if candidates:
                    break

    candidates = _filter_by_section(candidates, amount)

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

        return None, "ambigu", candidates

    return None, "non_trouve", []
