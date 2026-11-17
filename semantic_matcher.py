"""Conservative multilingual matching with exact matching, embeddings, and review."""

from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass
import os
import re
import unicodedata
from typing import Callable, Protocol, Sequence

import numpy as np
from scipy.sparse import csr_matrix
from scipy.sparse.csgraph import min_weight_full_bipartite_matching

from models import ExcelRecord, Match, MatchCandidate, MatchPlan, ReviewCase, ScoredMatchEdge


MODEL_NAME = "sentence-transformers/paraphrase-multilingual-MiniLM-L12-v2"
DEFAULT_THRESHOLD = 0.68
REVIEW_SCORE_MARGIN = 0.15
MAX_REVIEW_CANDIDATES = 5
RETRIEVAL_PER_SIDE = 10

_model_instance = None


class Embedder(Protocol):
    def encode(self, texts: Sequence[str], **kwargs): ...


class LocalSentenceEmbedder:
    """Lazy CPU model wrapper; first tries the Hugging Face cache before downloading."""

    def __init__(self, progress_callback: Callable[[str], None] | None = None):
        global _model_instance
        if _model_instance is None:
            if progress_callback:
                progress_callback("正在加载本地多语言语义模型；首次运行可能需要下载模型文件…")
            os.environ.setdefault("HF_HUB_DISABLE_PROGRESS_BARS", "1")
            os.environ.setdefault("TRANSFORMERS_VERBOSITY", "error")
            try:
                from sentence_transformers import SentenceTransformer
            except ImportError as exc:
                raise RuntimeError("缺少 sentence-transformers 依赖，请重新运行 run.bat 安装依赖。") from exc
            try:
                _model_instance = SentenceTransformer(MODEL_NAME, device="cpu", local_files_only=True)
            except Exception:
                try:
                    if progress_callback:
                        progress_callback("本地未找到模型，正在联网下载（之后可离线使用）…")
                    _model_instance = SentenceTransformer(MODEL_NAME, device="cpu", local_files_only=False)
                except Exception as exc:
                    raise RuntimeError(
                        "无法加载语义模型。首次使用请联网后重试；也可确认模型已缓存且未损坏。"
                    ) from exc
        self.model = _model_instance

    def encode(self, texts: Sequence[str], **kwargs):
        return self.model.encode(
            list(texts),
            batch_size=64,
            show_progress_bar=False,
            normalize_embeddings=True,
            convert_to_numpy=True,
        )


def normalize_match_text(value: object) -> str:
    if value is None:
        return ""
    text = unicodedata.normalize("NFKC", str(value)).replace("\r\n", "\n").replace("\r", "\n")
    return " ".join(text.split()).casefold()


_WORD_RE = re.compile(r"[a-z][a-z0-9_/-]*|[\u4e00-\u9fff]+", re.IGNORECASE)
_NUMBER_RE = re.compile(r"(?<![a-z])\d+(?:\.\d+)?(?:\s*(?:ms|msec|s|sec|seconds?|hz|khz|mhz|v|mv|a|ma|%|℃|°c))?", re.I)
_STOP_WORDS = {
    "a", "an", "the", "and", "or", "to", "of", "for", "in", "on", "at", "by", "as", "with",
    "from", "shall", "must", "should", "will", "can", "may", "be", "is", "are", "was", "were",
    "system", "ecu", "vehicle", "function", "requirement", "signal", "value", "data", "it", "this",
    "that", "when", "if", "then", "each", "every", "all", "any", "within", "during", "after", "before",
}
_GENERIC_CHINESE = {"系统", "车辆", "应当", "应该", "必须", "可以", "每个", "每次", "信号", "数据", "当", "如果", "则"}
_NEGATION = re.compile(r"\b(?:not|no|never|without|unable|cannot|can't|neither|nor)\b|不|未|无|不能|不得|禁止|避免", re.I)

_ACTION_GROUPS = (
    {"send", "sends", "sent", "transmit", "transmits", "transmitted", "report", "reports", "provide", "provides", "broadcast", "output", "发送", "传输", "上报", "输出", "提供"},
    {"detect", "detects", "identify", "identifies", "monitor", "monitors", "measure", "measures", "diagnose", "diagnoses", "检查", "检测", "监测", "识别", "测量"},
    {"control", "controls", "adjust", "adjusts", "regulate", "regulates", "limit", "limits", "控制", "调节", "调整", "限制"},
    {"receive", "receives", "read", "reads", "acquire", "acquires", "collect", "collects", "接收", "读取", "采集", "获取"},
    {"store", "stores", "save", "saves", "log", "logs", "record", "records", "存储", "保存", "记录"},
    {"activate", "activates", "enable", "enables", "open", "opens", "start", "starts", "激活", "启用", "开启", "启动"},
    {"deactivate", "deactivates", "disable", "disables", "stop", "stops", "close", "closes", "关闭", "停用", "停止"},
    {"warn", "warns", "alert", "alerts", "notify", "notifies", "报警", "警告", "通知"},
    {"calculate", "calculates", "compute", "computes", "estimate", "estimates", "计算", "估算"},
)
_ACTION_LOOKUP = {word: group_index for group_index, group in enumerate(_ACTION_GROUPS) for word in group}
_CONTRAST_PAIRS = (
    ("overvoltage", "undervoltage"),
    ("overcurrent", "undercurrent"),
    ("overspeed", "underspeed"),
    ("overtemperature", "undertemperature"),
    ("high voltage", "low voltage"),
    ("high pressure", "low pressure"),
    ("high temperature", "low temperature"),
    ("above limit", "below limit"),
    ("positive", "negative"),
    ("clockwise", "counterclockwise"),
    ("过压", "欠压"),
    ("过流", "欠流"),
    ("高电压", "低电压"),
    ("高温", "低温"),
    ("正向", "反向"),
    ("顺时针", "逆时针"),
)


def _opposite_object_concepts(left_text: str, right_text: str) -> bool:
    left, right = normalize_match_text(left_text), normalize_match_text(right_text)
    return any(
        (first in left and second in right) or (second in left and first in right)
        for first, second in _CONTRAST_PAIRS
    )


def _core_tokens(text: str) -> set[str]:
    tokens: set[str] = set()
    for token in _WORD_RE.findall(normalize_match_text(text)):
        if token.isascii():
            if token not in _STOP_WORDS and not token.isdigit():
                tokens.add(token)
        else:
            if token not in _GENERIC_CHINESE:
                if len(token) <= 2:
                    tokens.add(token)
                else:
                    tokens.update(token[i : i + 2] for i in range(len(token) - 1))
    return tokens


def _lexical_similarity(left: str, right: str) -> float:
    left_tokens, right_tokens = _core_tokens(left), _core_tokens(right)
    if not left_tokens or not right_tokens:
        return 0.0
    return 2.0 * len(left_tokens & right_tokens) / (len(left_tokens) + len(right_tokens))


def _action_groups(text: str) -> set[int]:
    result = set()
    normalized = normalize_match_text(text)
    for word, group_index in _ACTION_LOOKUP.items():
        if word.isascii():
            if re.search(rf"\b{re.escape(word)}\b", normalized):
                result.add(group_index)
        elif word in normalized:
            result.add(group_index)
    return result


def _auxiliary_similarity(
    before: ExcelRecord,
    after: ExcelRecord,
    before_content_column: int,
    after_content_column: int,
    before_headers: Sequence[str] | None = None,
    after_headers: Sequence[str] | None = None,
) -> float:
    left_headers = list(before_headers or ())
    right_headers = list(after_headers or ())
    after_by_header: dict[str, list[int]] = defaultdict(list)
    for index, header in enumerate(right_headers):
        normalized = normalize_match_text(header)
        if normalized and index != after_content_column:
            after_by_header[normalized].append(index)
    used_right: set[int] = set()
    pairs: list[tuple[int, int, str]] = []
    for left_index in range(len(before.values)):
        if left_index == before_content_column:
            continue
        header = left_headers[left_index] if left_index < len(left_headers) else ""
        normalized = normalize_match_text(header)
        named = next((index for index in after_by_header.get(normalized, ()) if index not in used_right), None) if normalized else None
        if named is not None:
            right_index = named
        elif left_index < len(after.values) and left_index != after_content_column and left_index not in used_right:
            right_index = left_index
        else:
            continue
        used_right.add(right_index)
        pairs.append((left_index, right_index, normalized))

    weighted_matches = 0.0
    weights = 0.0
    for left_index, right_index, header in pairs:
        left = before.values[left_index]
        right = after.values[right_index]
        if left in (None, "") and right in (None, ""):
            continue
        is_identifier = header in {"id", "identifier", "object id", "requirement id", "req id", "item id"}
        weight = 2.0 if is_identifier else 1.0
        weights += weight
        if normalize_match_text(left) == normalize_match_text(right):
            weighted_matches += weight
    return weighted_matches / weights if weights else 0.0


def _semantic_score(
    before: ExcelRecord,
    after: ExcelRecord,
    left_text: str,
    right_text: str,
    cosine: float,
    before_content_column: int,
    after_content_column: int,
    before_headers: Sequence[str] | None = None,
    after_headers: Sequence[str] | None = None,
) -> float:
    lexical = _lexical_similarity(left_text, right_text)
    auxiliary = _auxiliary_similarity(
        before, after, before_content_column, after_content_column, before_headers, after_headers
    )
    score = 0.92 * float(np.clip(cosine, -1.0, 1.0)) + 0.06 * lexical + 0.02 * auxiliary
    if bool(_NEGATION.search(left_text)) != bool(_NEGATION.search(right_text)):
        score -= 0.22
    if _opposite_object_concepts(left_text, right_text):
        score -= 0.25
    left_actions, right_actions = _action_groups(left_text), _action_groups(right_text)
    if left_actions and right_actions and not (left_actions & right_actions):
        score -= 0.24
    return float(np.clip(score, 0.0, 1.0))


@dataclass(frozen=True)
class _Edge:
    left_index: int
    right_index: int
    score: float
    method: str


def _stable_row_signature(record: ExcelRecord, content_column: int) -> tuple[str, ...]:
    return tuple(normalize_match_text(value) for index, value in enumerate(record.values) if index != content_column)


def _candidate_edges_from_similarity(
    left_records: Sequence[ExcelRecord],
    right_records: Sequence[ExcelRecord],
    before_content_column: int,
    after_content_column: int,
    threshold: float,
    embedder: Embedder,
    before_headers: Sequence[str] | None = None,
    after_headers: Sequence[str] | None = None,
) -> list[_Edge]:
    left_texts = [str(record.values[before_content_column] or "") for record in left_records]
    right_texts = [str(record.values[after_content_column] or "") for record in right_records]
    left_vectors = np.asarray(embedder.encode(left_texts, batch_size=64, normalize_embeddings=True), dtype=np.float32)
    right_vectors = np.asarray(embedder.encode(right_texts, batch_size=64, normalize_embeddings=True), dtype=np.float32)
    if left_vectors.ndim != 2 or right_vectors.ndim != 2:
        raise ValueError("语义模型返回了无效的向量维度。")
    # Normalize even test/custom encoders. Similarity is computed in bounded batches.
    left_norms = np.maximum(np.linalg.norm(left_vectors, axis=1, keepdims=True), 1e-12)
    right_norms = np.maximum(np.linalg.norm(right_vectors, axis=1, keepdims=True), 1e-12)
    left_vectors /= left_norms
    right_vectors /= right_norms
    per_left: list[set[int]] = [set() for _ in left_records]
    per_right: list[list[tuple[float, int]]] = [[] for _ in right_records]
    batch_size = 128
    top_k = min(RETRIEVAL_PER_SIDE, len(left_records), len(right_records))
    for start in range(0, len(left_records), batch_size):
        end = min(len(left_records), start + batch_size)
        similarities = left_vectors[start:end] @ right_vectors.T
        for local_index, row in enumerate(similarities):
            left_index = start + local_index
            best = np.argpartition(row, -top_k)[-top_k:]
            per_left[left_index].update(int(index) for index in best)
        if top_k:
            batch_top_k = min(top_k, similarities.shape[0])
            best_left = np.argpartition(similarities, -batch_top_k, axis=0)[-batch_top_k:, :]
            for right_index in range(len(right_records)):
                candidates = per_right[right_index]
                candidates.extend(
                    (float(similarities[local_index, right_index]), start + int(local_index))
                    for local_index in best_left[:, right_index]
                )
                per_right[right_index] = sorted(candidates, key=lambda item: (-item[0], item[1]))[:top_k]

    pairs = {(left_index, right_index) for left_index, choices in enumerate(per_left) for right_index in choices}
    pairs.update(
        (left_index, right_index)
        for right_index, choices in enumerate(per_right)
        for _, left_index in choices
    )
    floor = max(0.0, threshold - REVIEW_SCORE_MARGIN)
    edges: list[_Edge] = []
    for left_index, right_index in sorted(pairs):
        cosine = float(np.dot(left_vectors[left_index], right_vectors[right_index]))
        score = _semantic_score(
            left_records[left_index], right_records[right_index],
            left_texts[left_index], right_texts[right_index], cosine,
            before_content_column, after_content_column, before_headers, after_headers,
        )
        if score >= floor:
            edges.append(_Edge(left_index, right_index, score, "semantic"))
    return edges


def _connected_components(edges: Sequence[_Edge]) -> list[tuple[list[int], list[int], list[_Edge]]]:
    left_to_edges: dict[int, list[_Edge]] = defaultdict(list)
    right_to_edges: dict[int, list[_Edge]] = defaultdict(list)
    for edge in edges:
        left_to_edges[edge.left_index].append(edge)
        right_to_edges[edge.right_index].append(edge)
    seen_left: set[int] = set()
    components = []
    for seed in sorted(left_to_edges):
        if seed in seen_left:
            continue
        stack = [(0, seed)]
        component_left: set[int] = set()
        component_right: set[int] = set()
        component_edges: dict[tuple[int, int], _Edge] = {}
        while stack:
            side, index = stack.pop()
            if side == 0:
                if index in component_left:
                    continue
                component_left.add(index)
                seen_left.add(index)
                for edge in left_to_edges[index]:
                    component_edges[(edge.left_index, edge.right_index)] = edge
                    stack.append((1, edge.right_index))
            else:
                if index in component_right:
                    continue
                component_right.add(index)
                for edge in right_to_edges[index]:
                    component_edges[(edge.left_index, edge.right_index)] = edge
                    stack.append((0, edge.left_index))
        components.append((sorted(component_left), sorted(component_right), list(component_edges.values())))
    return components


def _rank_component(
    left_records: Sequence[ExcelRecord],
    right_records: Sequence[ExcelRecord],
    edges: Sequence[_Edge],
    threshold: float,
    ambiguity_margin: float,
) -> tuple[list[Match], list[ReviewCase]]:
    auto_matches: list[Match] = []
    review_cases: list[ReviewCase] = []
    floor = max(0.0, threshold - REVIEW_SCORE_MARGIN)
    for left_indices, right_indices, component_edges in _connected_components(edges):
        left_pos = {value: index for index, value in enumerate(left_indices)}
        right_pos = {value: index for index, value in enumerate(right_indices)}
        edge_lookup: dict[tuple[int, int], _Edge] = {}
        by_left: dict[int, list[_Edge]] = defaultdict(list)
        by_right: dict[int, list[_Edge]] = defaultdict(list)
        for edge in component_edges:
            edge_lookup[(edge.left_index, edge.right_index)] = edge
            by_left[edge.left_index].append(edge)
            by_right[edge.right_index].append(edge)
        sparse_rows, sparse_columns, sparse_costs = [], [], []
        for edge in component_edges:
            sparse_rows.append(left_pos[edge.left_index])
            sparse_columns.append(right_pos[edge.right_index])
            sparse_costs.append(-max(0.0, edge.score - floor) - 1e-12)
        for row_index in range(len(left_indices)):
            # One private dummy column per Before record represents a real unmatched option.
            sparse_rows.append(row_index)
            sparse_columns.append(len(right_indices) + row_index)
            sparse_costs.append(1e-9)
        costs = csr_matrix(
            (np.asarray(sparse_costs), (np.asarray(sparse_rows), np.asarray(sparse_columns))),
            shape=(len(left_indices), len(right_indices) + len(left_indices)),
        )
        selected_rows, selected_cols = min_weight_full_bipartite_matching(costs)
        assignment = {
            left_indices[row_index]: right_indices[column_index]
            for row_index, column_index in zip(selected_rows, selected_cols)
            if column_index < len(right_indices)
            and (left_indices[row_index], right_indices[column_index]) in edge_lookup
            and edge_lookup[(left_indices[row_index], right_indices[column_index])].score >= threshold
        }
        auto_left: set[int] = set()
        auto_right: set[int] = set()
        for left_index, right_index in assignment.items():
            edge = edge_lookup[(left_index, right_index)]
            left_options = sorted((candidate.score for candidate in by_left[left_index]), reverse=True)
            right_options = sorted((candidate.score for candidate in by_right[right_index]), reverse=True)
            left_gap = left_options[0] - left_options[1] if len(left_options) > 1 else 1.0
            right_gap = right_options[0] - right_options[1] if len(right_options) > 1 else 1.0
            if edge.score >= threshold and left_gap >= ambiguity_margin and right_gap >= ambiguity_margin:
                auto_matches.append(Match(left_records[left_index], right_records[right_index], edge.score, edge.method))
                auto_left.add(left_index)
                auto_right.add(right_index)

        for left_index in left_indices:
            if left_index in auto_left:
                continue
            options = sorted(by_left[left_index], key=lambda item: (-item.score, right_records[item.right_index].row_number))
            options = [item for item in options if item.score >= floor and item.right_index not in auto_right]
            if options:
                review_cases.append(
                    ReviewCase(
                        left_records[left_index],
                        tuple(
                            MatchCandidate(right_records[edge.right_index], edge.score, edge.method)
                            for edge in options[:MAX_REVIEW_CANDIDATES]
                        ),
                    )
                )
    return auto_matches, review_cases


def match_records(
    before_records: Sequence[ExcelRecord],
    after_records: Sequence[ExcelRecord],
    content_column: int,
    threshold: float = DEFAULT_THRESHOLD,
    embedder: Embedder | None = None,
    ambiguity_margin: float = 0.04,
    progress_callback: Callable[[str], None] | None = None,
    before_content_column: int | None = None,
    after_content_column: int | None = None,
    semantic_enabled: bool = True,
    before_headers: Sequence[str] | None = None,
    after_headers: Sequence[str] | None = None,
) -> MatchPlan:
    """Match records one-to-one; uncertain pairs are returned for explicit user review."""
    if not 0.0 <= threshold <= 1.0:
        raise ValueError("匹配阈值必须在 0 到 1 之间。")
    before_column = content_column if before_content_column is None else before_content_column
    after_column = content_column if after_content_column is None else after_content_column
    if any(before_column < 0 or before_column >= len(record.values) for record in before_records):
        raise ValueError("Before 内容列选择超出记录字段范围。")
    if any(after_column < 0 or after_column >= len(record.values) for record in after_records):
        raise ValueError("After 内容列选择超出记录字段范围。")

    left_groups: dict[str, list[int]] = defaultdict(list)
    right_groups: dict[str, list[int]] = defaultdict(list)
    for index, record in enumerate(before_records):
        content = normalize_match_text(record.values[before_column])
        if content:
            left_groups[content].append(index)
    for index, record in enumerate(after_records):
        content = normalize_match_text(record.values[after_column])
        if content:
            right_groups[content].append(index)

    matched: list[Match] = []
    review_cases: list[ReviewCase] = []
    used_left: set[int] = set()
    used_right: set[int] = set()
    duplicate_edges: list[_Edge] = []
    for content in sorted(set(left_groups) & set(right_groups)):
        left_indices, right_indices = left_groups[content], right_groups[content]
        if len(left_indices) == len(right_indices) == 1:
            left_index, right_index = left_indices[0], right_indices[0]
            matched.append(Match(before_records[left_index], after_records[right_index], 1.0, "exact"))
            used_left.add(left_index)
            used_right.add(right_index)
            continue
        left_signatures = {_stable_row_signature(before_records[i], before_column) for i in left_indices}
        right_signatures = {_stable_row_signature(after_records[i], after_column) for i in right_indices}
        if len(left_signatures) == 1 and len(right_signatures) == 1 and left_signatures == right_signatures:
            for left_index, right_index in zip(
                sorted(left_indices, key=lambda i: before_records[i].row_number),
                sorted(right_indices, key=lambda i: after_records[i].row_number),
            ):
                matched.append(Match(before_records[left_index], after_records[right_index], 1.0, "exact"))
                used_left.add(left_index)
                used_right.add(right_index)
            continue
        # Bounded candidate graph; IDs and other fields resolve duplicate text without row-order pairing.
        pair_scores = []
        for left_index in left_indices:
            options = sorted(
                ((_auxiliary_similarity(before_records[left_index], after_records[right_index], before_column, after_column, before_headers, after_headers), right_index)
                 for right_index in right_indices),
                key=lambda item: (-item[0], after_records[item[1]].row_number),
            )[:RETRIEVAL_PER_SIDE]
            pair_scores.extend((left_index, right_index, score) for score, right_index in options)
        for right_index in right_indices:
            options = sorted(
                ((_auxiliary_similarity(before_records[left_index], after_records[right_index], before_column, after_column, before_headers, after_headers), left_index)
                 for left_index in left_indices),
                key=lambda item: (-item[0], before_records[item[1]].row_number),
            )[:RETRIEVAL_PER_SIDE]
            pair_scores.extend((left_index, right_index, score) for score, left_index in options)
        for left_index, right_index, aux_score in sorted(set(pair_scores)):
            duplicate_edges.append(_Edge(left_index, right_index, 0.80 + 0.20 * aux_score, "exact-duplicate"))

    if duplicate_edges:
        duplicate_matches, duplicate_reviews = _rank_component(
            before_records, after_records, duplicate_edges, threshold=0.0, ambiguity_margin=ambiguity_margin
        )
        matched.extend(duplicate_matches)
        review_cases.extend(duplicate_reviews)
        matched_rows_before = {match.before.row_number for match in duplicate_matches}
        matched_rows_after = {match.after.row_number for match in duplicate_matches}
        used_left.update(i for i, record in enumerate(before_records) if record.row_number in matched_rows_before)
        used_right.update(i for i, record in enumerate(after_records) if record.row_number in matched_rows_after)

    reserved_review_before = {case.before.row_number for case in review_cases}
    reserved_review_after = {
        candidate.after.row_number for case in review_cases for candidate in case.candidates
    }

    remaining_left_indices = [
        i for i, record in enumerate(before_records)
        if i not in used_left and record.row_number not in reserved_review_before
    ]
    remaining_right_indices = [
        i for i, record in enumerate(after_records)
        if i not in used_right and record.row_number not in reserved_review_after
    ]
    remaining_left = [before_records[i] for i in remaining_left_indices]
    remaining_right = [after_records[i] for i in remaining_right_indices]
    fuzzy_left = [record for record in remaining_left if str(record.values[before_column] or "").strip()]
    fuzzy_right = [record for record in remaining_right if str(record.values[after_column] or "").strip()]
    semantic_edges: list[_Edge] = []
    all_scored_edges: list[ScoredMatchEdge] = [
        ScoredMatchEdge(before_records[edge.left_index], after_records[edge.right_index], edge.score, edge.method)
        for edge in duplicate_edges
    ]
    if semantic_enabled and fuzzy_left and fuzzy_right:
        if progress_callback:
            progress_callback("正在计算多语言语义向量并查找候选对应项…")
        active_embedder = embedder or LocalSentenceEmbedder(progress_callback)
        local_edges = _candidate_edges_from_similarity(
            fuzzy_left, fuzzy_right, before_column, after_column, threshold, active_embedder,
            before_headers, after_headers,
        )
        fuzzy_left_positions = {id(record): i for i, record in enumerate(remaining_left)}
        fuzzy_right_positions = {id(record): i for i, record in enumerate(remaining_right)}
        semantic_edges = [
            _Edge(
                fuzzy_left_positions[id(fuzzy_left[edge.left_index])],
                fuzzy_right_positions[id(fuzzy_right[edge.right_index])],
                edge.score,
                edge.method,
            )
            for edge in local_edges
        ]
        all_scored_edges.extend(
            ScoredMatchEdge(
                remaining_left[edge.left_index], remaining_right[edge.right_index], edge.score, edge.method
            )
            for edge in semantic_edges
        )
        auto_matches, semantic_reviews = _rank_component(
            remaining_left, remaining_right, semantic_edges, threshold, ambiguity_margin
        )
        matched.extend(auto_matches)
        review_cases.extend(semantic_reviews)

    matched_before_rows = {match.before.row_number for match in matched}
    matched_after_rows = {match.after.row_number for match in matched}
    # A review candidate is deliberately still unpaired; later user decisions can consume either side.
    unmatched_before = [record for record in before_records if record.row_number not in matched_before_rows]
    unmatched_after = [record for record in after_records if record.row_number not in matched_after_rows]
    return MatchPlan(
        matches=sorted(matched, key=lambda match: (match.before.row_number, match.after.row_number)),
        review_cases=sorted(review_cases, key=lambda case: case.before.row_number),
        unmatched_before=unmatched_before,
        unmatched_after=unmatched_after,
        candidate_edges=tuple(all_scored_edges),
        threshold=threshold,
        ambiguity_margin=ambiguity_margin,
    )


def apply_review_decisions(plan: MatchPlan, decisions: dict[int, int | None]) -> MatchPlan:
    """Apply row-number keyed confirmations/rejections while preserving one-to-one matches."""
    review_by_before = {case.before.row_number: case for case in plan.review_cases}
    known_after = {record.row_number for record in plan.unmatched_after}
    used_before = {match.before.row_number for match in plan.matches}
    used_after = {match.after.row_number for match in plan.matches}
    resolved: list[Match] = list(plan.matches)
    handled: set[int] = set()
    rejected = set(plan.rejected_before_rows)
    for before_row, after_row in decisions.items():
        case = review_by_before.get(before_row)
        if case is None:
            raise ValueError(f"第 {before_row} 行没有待复核的候选匹配。")
        if before_row in used_before:
            raise ValueError(f"Before 第 {before_row} 行已经匹配。")
        handled.add(before_row)
        if after_row is None:
            rejected.add(before_row)
            continue
        candidate = next((item for item in case.candidates if item.after.row_number == after_row), None)
        if after_row not in known_after:
            raise ValueError(f"After 第 {after_row} 行已被占用或不在筛选结果中。")
        if after_row in used_after:
            raise ValueError(f"After 第 {after_row} 行不能重复匹配。")
        after_record = candidate.after if candidate is not None else next(
            record for record in plan.unmatched_after if record.row_number == after_row
        )
        resolved.append(
            Match(case.before, after_record, candidate.score if candidate is not None else 0.0, "manual", manual=True)
        )
        used_before.add(before_row)
        used_after.add(after_row)

    fixed_before = {match.before.row_number for match in resolved}
    fixed_after = {match.after.row_number for match in resolved}
    available_before = [record for record in plan.unmatched_before if record.row_number not in fixed_before and record.row_number not in rejected]
    available_after = [record for record in plan.unmatched_after if record.row_number not in fixed_after]
    before_position = {record.row_number: index for index, record in enumerate(available_before)}
    after_position = {record.row_number: index for index, record in enumerate(available_after)}
    available_edges = [
        edge for edge in plan.candidate_edges
        if edge.before.row_number in before_position and edge.after.row_number in after_position
    ]
    recomputed_matches: list[Match] = []
    remaining_reviews: list[ReviewCase] = []
    for method in ("exact-duplicate", "semantic"):
        method_edges = [edge for edge in available_edges if edge.method == method]
        if not method_edges:
            continue
        indexed_edges = [
            _Edge(before_position[edge.before.row_number], after_position[edge.after.row_number], edge.score, edge.method)
            for edge in method_edges
        ]
        matches, reviews = _rank_component(
            available_before,
            available_after,
            indexed_edges,
            threshold=0.0 if method == "exact-duplicate" else plan.threshold,
            ambiguity_margin=plan.ambiguity_margin,
        )
        recomputed_matches.extend(matches)
        remaining_reviews.extend(reviews)
    resolved.extend(recomputed_matches)
    resolved_before = {match.before.row_number for match in resolved}
    resolved_after = {match.after.row_number for match in resolved}
    return MatchPlan(
        matches=sorted(resolved, key=lambda match: (match.before.row_number, match.after.row_number)),
        review_cases=sorted(remaining_reviews, key=lambda case: case.before.row_number),
        unmatched_before=[record for record in plan.unmatched_before if record.row_number not in resolved_before],
        unmatched_after=[record for record in plan.unmatched_after if record.row_number not in resolved_after],
        candidate_edges=tuple(available_edges),
        threshold=plan.threshold,
        ambiguity_margin=plan.ambiguity_margin,
        rejected_before_rows=tuple(sorted(rejected)),
    )
