"""Utilities for merging translations into multiple target workbooks.

The existing single-target merge flow expects the caller to provide a mapping
for each source Excel file and merges them into one common workbook. This
module extends that workflow to allow several target workbooks while keeping
compatibility with :func:`core.merge_columns.merge_excel_columns`.

Usage outline::

    plans = [
        TargetMergePlan(
            main_file="projectA.xlsx",
            mappings=[{"source_id": "projA", "target_sheet": "Sheet1",
                       "source_columns": ["B"], "target_columns": ["B"]}]
        ),
        TargetMergePlan(
            main_file="projectB.xlsx",
            mappings=[{"source_id": "projB", "target_sheet": "Sheet1",
                       "source_columns": ["B"], "target_columns": ["B"]}]
        ),
    ]

    outputs = merge_translations_to_many_targets(plans, source_files)

The automatic matcher pairs sources with targets using filename similarity.
Callers can override the assignment with ``manual_mapping``.
"""

from __future__ import annotations

import os
import re
from dataclasses import dataclass
from difflib import SequenceMatcher
from typing import Callable, Dict, Iterable, List, Mapping

from core.merge_columns import merge_excel_columns

ProgressFn = Callable[[int, str], None]


@dataclass
class TargetMergePlan:
    """Configuration for merging translations into a single workbook.

    Attributes:
        main_file: Path to the target workbook that will receive translations.
        mappings: A list of dictionaries compatible with
            :func:`core.merge_columns.merge_excel_columns`. The ``source`` key
            may be omitted and replaced with ``source_id`` for automatic
            resolution based on filename similarity.
        output_file: Optional explicit output path. If omitted the helper
            mirrors ``merge_excel_columns`` behaviour and suffixes the target
            file name with ``_merged``.
    """

    main_file: str
    mappings: List[Dict[str, object]]
    output_file: str | None = None


def normalize_name(path: str) -> str:
    """Return a simplified, comparable name derived from the file path."""

    base = os.path.splitext(os.path.basename(path))[0]
    return re.sub(r"[^a-z0-9]+", "", base.lower())


def suggest_target_mapping(
    source_files: Iterable[str],
    target_files: Iterable[str],
    manual_mapping: Mapping[str, str] | None = None,
    target_hints: Mapping[str, Iterable[str]] | None = None,
) -> Dict[str, List[str]]:
    """Assign translation files to target workbooks.

    The matcher uses filename similarity to guess the correct target when
    several options are available. Callers can provide ``manual_mapping`` to
    force particular assignments and ``target_hints`` (expected source IDs for
    each target) to improve automatic matching. Any source not present in
    ``manual_mapping`` will be matched automatically. If there is only a
    single target file, all remaining sources are routed there.
    """

    manual_mapping = manual_mapping or {}
    target_hints = target_hints or {}
    target_files = list(target_files)
    assignments: Dict[str, List[str]] = {t: [] for t in target_files}
    remaining: List[str] = []

    for src in source_files:
        target = manual_mapping.get(src)
        if target:
            assignments.setdefault(target, []).append(src)
        else:
            remaining.append(src)

    if len(target_files) == 1:
        assignments[target_files[0]].extend(remaining)
        return assignments

    normalized_targets = {t: normalize_name(t) for t in target_files}

    for src in remaining:
        norm_src = normalize_name(src)
        best_target = None
        best_score = -1.0
        for target, norm_target in normalized_targets.items():
            if not norm_target:
                continue
            score = SequenceMatcher(None, norm_src, norm_target).ratio()
            if norm_target and norm_target in norm_src:
                score += 0.5  # prefer clear substring matches
            for hint in target_hints.get(target, []):
                hint_score = SequenceMatcher(None, norm_src, normalize_name(str(hint))).ratio() + 0.2
                score = max(score, hint_score)
            if score > best_score:
                best_score = score
                best_target = target

        if best_target is None or best_score < 0.35:
            counts = {t: len(assignments.get(t, [])) for t in target_files}
            best_target = sorted(counts.items(), key=lambda item: (item[1], item[0]))[0][0]
        assignments[best_target].append(src)

    return assignments


def _resolve_source_hint(source_hint: str | None, candidates: List[str]) -> str | None:
    if not source_hint:
        return candidates[0] if candidates else None
    norm_hint = normalize_name(source_hint)
    if not candidates:
        return None
    best = None
    best_score = -1.0
    for candidate in candidates:
        score = SequenceMatcher(None, norm_hint, normalize_name(candidate)).ratio()
        if score > best_score:
            best = candidate
            best_score = score
    return best


def _build_plan_mappings(plan: TargetMergePlan, sources: List[str]) -> List[Dict[str, object]]:
    prepared: List[Dict[str, object]] = []
    for mapping in plan.mappings:
        if mapping.get("source"):
            prepared.append(dict(mapping))
            continue
        source_hint = mapping.get("source_id") or mapping.get("source_name")
        matched_source = _resolve_source_hint(str(source_hint) if source_hint is not None else None, sources)
        if not matched_source:
            raise ValueError(
                f"Не удалось сопоставить источник для '{source_hint}' при обработке {plan.main_file}"
            )
        new_mapping = dict(mapping)
        new_mapping["source"] = matched_source
        prepared.append(new_mapping)
    return prepared


def merge_translations_to_many_targets(
    plans: Iterable[TargetMergePlan],
    source_files: Iterable[str],
    manual_mapping: Mapping[str, str] | None = None,
    progress_callback: ProgressFn | None = None,
) -> List[str]:
    """Merge translations from ``source_files`` into several workbooks.

    Args:
        plans: Iterable of :class:`TargetMergePlan` describing each target.
        source_files: Paths to translation Excel files.
        manual_mapping: Optional explicit mapping ``{source: target}`` that
            overrides the automatic matcher.
        progress_callback: Optional callback ``(percent, message)`` invoked with
            aggregated progress across all targets.

    Returns:
        List of output file paths produced by the merges.
    """

    plan_list = list(plans)
    target_files = [plan.main_file for plan in plan_list]
    target_hints = {
        plan.main_file: [
            m.get("source_id") or m.get("source_name")
            for m in plan.mappings
            if m.get("source_id") or m.get("source_name")
        ]
        for plan in plan_list
    }
    assignments = suggest_target_mapping(
        source_files, target_files, manual_mapping, target_hints=target_hints
    )

    outputs: List[str] = []
    total_targets = len(plan_list)

    for index, plan in enumerate(plan_list, start=1):
        assigned_sources = assignments.get(plan.main_file, [])
        prepared_mappings = _build_plan_mappings(plan, assigned_sources)

        def _callback(idx: int, total: int, mapping: Dict[str, object]):
            if not progress_callback:
                return
            base_progress = int(((index - 1) / total_targets) * 100)
            merge_progress = int((idx / max(total, 1)) * (100 / total_targets))
            progress_callback(base_progress + merge_progress, str(mapping))

        output = merge_excel_columns(
            plan.main_file,
            prepared_mappings,
            output_file=plan.output_file,
            progress_callback=_callback if progress_callback else None,
        )
        outputs.append(output)

    if progress_callback:
        progress_callback(100, "Все объединения завершены")

    return outputs
