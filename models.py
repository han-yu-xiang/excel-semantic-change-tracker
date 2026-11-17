"""Typed records shared by the Excel reader, matcher, detector, and writer."""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any


@dataclass(frozen=True)
class ColumnInfo:
    index: int
    letter: str
    header: str

    @property
    def number(self) -> int:
        return self.index + 1


@dataclass(frozen=True)
class ExcelRecord:
    row_number: int
    values: tuple[Any, ...]


@dataclass(frozen=True)
class SheetPreview:
    path: str | Path
    sheet_name: str
    columns: tuple[ColumnInfo, ...]
    data_row_count: int


@dataclass(frozen=True)
class WorkbookSnapshot:
    path: str | Path
    sheet_name: str
    columns: tuple[ColumnInfo, ...]
    records: tuple[ExcelRecord, ...]
    source_data_row_count: int

    @property
    def file_name(self) -> str:
        return Path(self.path).name

    @property
    def filtered_record_count(self) -> int:
        return len(self.records)


@dataclass(frozen=True)
class Match:
    before: ExcelRecord
    after: ExcelRecord
    score: float
    method: str
    manual: bool = False


@dataclass(frozen=True)
class MatchCandidate:
    after: ExcelRecord
    score: float
    method: str


@dataclass(frozen=True)
class ReviewCase:
    before: ExcelRecord
    candidates: tuple[MatchCandidate, ...]


@dataclass(frozen=True)
class ScoredMatchEdge:
    before: ExcelRecord
    after: ExcelRecord
    score: float
    method: str


@dataclass
class MatchPlan:
    matches: list[Match] = field(default_factory=list)
    review_cases: list[ReviewCase] = field(default_factory=list)
    unmatched_before: list[ExcelRecord] = field(default_factory=list)
    unmatched_after: list[ExcelRecord] = field(default_factory=list)
    candidate_edges: tuple[ScoredMatchEdge, ...] = ()
    threshold: float = 0.77
    ambiguity_margin: float = 0.04
    rejected_before_rows: tuple[int, ...] = ()


@dataclass(frozen=True)
class ChangeRow:
    change_type: str
    before_location: str | int
    after_location: str | int
    before_values: tuple[Any, ...]
    after_values: tuple[Any, ...]
    changed_column_indices: tuple[int, ...]


@dataclass(frozen=True)
class ReportStats:
    changed_count: int
    new_count: int
    out_count: int
    still_count: int
    before_count: int
    after_count: int
    matched_count: int
    manual_confirmed_count: int
    pending_review_count: int = 0


@dataclass(frozen=True)
class ChangeReport:
    rows: list[ChangeRow]
    stats: ReportStats
    pending_candidates: tuple[ReviewCase, ...] = ()
