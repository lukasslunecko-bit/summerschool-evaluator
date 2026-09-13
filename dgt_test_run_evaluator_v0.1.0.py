#!/usr/bin/env python3
"""DGT Test Run Evaluator — append tool findings and evaluation metrics to a register.

The evaluator reads one golden-standard CSV, one modified CSV, a dataset-register
XLSX workbook, and a findings file produced by a checking tool.  It appends
detailed rows to ``findings`` and a run summary to ``evaluation``.  Missing
worksheets are created automatically.

Like the DGT Data Converter, this is a standalone, standard-library-only tool.
Run it directly for the desktop UI or use ``--headless`` for automation.  The
CSV/XML/JSON inputs are always read-only.  The selected register is replaced
atomically and, by default, backed up first.
"""

from __future__ import annotations

import argparse
import csv
import getpass
import hashlib
import io
import json
import os
import posixpath
import queue
import re
import shutil
import sys
import threading
import traceback
import unicodedata
import zipfile
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Iterable, Sequence
from xml.etree import ElementTree as ET

try:
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk
    from tkinter.scrolledtext import ScrolledText
except ImportError:  # pragma: no cover - only relevant to minimal Python builds
    tk = None


APP_NAME = "DGT Test Run Evaluator"
APP_VERSION = "0.1.0"
SETTINGS_SCHEMA_VERSION = 1
SETTINGS_PATH = Path(__file__).resolve().with_name("dgt_test_run_evaluator_settings.json")
SUPPORTED_FINDINGS_EXTENSIONS = {".csv", ".tsv", ".json", ".xml"}

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
DOC_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
NS = {"m": MAIN_NS, "r": DOC_REL_NS, "p": PKG_REL_NS, "ct": CONTENT_TYPES_NS}
ET.register_namespace("", MAIN_NS)
ET.register_namespace("r", DOC_REL_NS)

FINDINGS_SHEET = "findings"
EVALUATION_SHEET = "evaluation"

FINDINGS_COLUMNS = (
    "Run_ID",
    "Finding_ID",
    "Tester",
    "Tool",
    "Filename",
    "Segment_number",
    "ORI",
    "Correct_TRA",
    "Modified_TRA",
    "Finding",
    "Severity",
    "Source_details",
    "Matches_seeded_error",
    "Matched_Error_IDs",
    "Match_method",
    "Seeded_comment",
    "Seeded_changed_words",
    "Seeded_similar_segments",
    "Findings_file",
    "Processed_at",
)

EVALUATION_COLUMNS = (
    "Run_ID",
    "Processed_at",
    "Findings_file",
    "Findings_SHA256",
    "Tool",
    "Golden_file",
    "Modified_file",
    "Total_findings",
    "True_positives",
    "False_positives",
    "Unique_seeded_errors_matched",
    "Seeded_errors_in_scope",
    "False_negatives",
    "Precision",
    "Recall",
    "Changed_segments_in_CSVs",
    "Notes",
)

HEADER_ALIASES: dict[str, set[str]] = {
    "error_id": {"errorid", "seedederrorid", "id"},
    "tester": {"username", "tester", "user", "author"},
    "filename": {"filename", "file", "sourcefile", "document", "documentname"},
    "segment": {
        "segment", "segmentnumber", "segmentno", "segmentid", "segmentkey",
        "row", "rownumber", "line", "linenumber", "unit", "unitid",
    },
    "ori": {"ori", "source", "sourcetext", "original", "originaltext"},
    "correct_tra": {
        "correcttra", "goldentra", "goldtranslation", "expectedtra",
        "expectedtranslation", "correcttranslation", "reference",
    },
    "modified_tra": {
        "modifiedtra", "tra", "target", "targettext", "translation",
        "actualtra", "actualtranslation", "modifiedtranslation",
    },
    "comment": {
        "commentoptional", "comment", "comments", "description", "details",
        "message", "finding", "issue", "errormessage", "reason",
    },
    "changed_words": {"changedwords", "changedword", "differences", "difference"},
    "similar_segments": {"similarsegments", "similarsegment", "relatedsegments"},
    "finding_id": {"findingid", "issueid", "resultid", "checkid", "id"},
    "severity": {"severity", "level", "priority", "risk", "category"},
}

ROLE_WORDS = {
    "adjusted", "golden", "standard", "gold", "modified", "modify", "errors",
    "error", "with", "inconsistent", "inconsistencified", "comparison",
    "merged", "notepad", "plain", "converted", "output", "result", "results",
    "report", "reports", "reference", "file", "files", "raw", "copy",
}
FILE_EXT_WORDS = {
    "csv", "tsv", "xlsx", "xls", "doc", "docx", "sdlxliff", "xliff",
    "xlf", "tmx", "xml", "json",
}


class EvaluationError(ValueError):
    """A clear, expected problem with selected files or mappings."""


class DuplicateRunError(EvaluationError):
    """Raised when the same findings content is already registered."""


@dataclass
class DatasetRow:
    segment: str
    ori: str
    tra: str
    row_number: int


@dataclass
class DatasetFile:
    path: Path
    rows: list[DatasetRow]
    delimiter: str
    encoding: str
    segment_column: str | None
    ori_column: str | None
    tra_column: str | None
    implicit_segments: bool

    def by_segment(self) -> dict[str, list[DatasetRow]]:
        result: dict[str, list[DatasetRow]] = {}
        for row in self.rows:
            tokens = parse_segment_tokens(row.segment) or {normalise_segment(row.segment)}
            for token in tokens:
                if token:
                    result.setdefault(token, []).append(row)
        return result


@dataclass
class FindingsData:
    path: Path
    records: list[dict[str, str]]
    headers: list[str]
    format_name: str
    delimiter: str = ""
    encoding: str = ""


@dataclass
class SeededError:
    error_id: str
    tester: str
    filename: str
    segment: str
    ori: str
    correct_tra: str
    modified_tra: str
    comment: str
    changed_words: str
    similar_segments: str

    @property
    def segments(self) -> set[str]:
        return parse_segment_tokens(self.segment)

    @property
    def substantive(self) -> bool:
        return any((self.filename, self.segment, self.ori, self.correct_tra, self.modified_tra, self.comment))


@dataclass
class ColumnMapping:
    segment: str | None = None
    filename: str | None = None
    message: str | None = None
    severity: str | None = None
    finding_id: str | None = None
    ori: str | None = None
    correct_tra: str | None = None
    modified_tra: str | None = None


@dataclass
class EvaluationOptions:
    golden: Path
    modified: Path
    register: Path
    findings: Path
    output_register: Path | None = None
    update_in_place: bool = True
    create_backup: bool = True
    allow_duplicate: bool = False
    tester: str = ""
    tool_name: str = ""
    segment_offset: int = 0
    mappings: ColumnMapping = field(default_factory=ColumnMapping)
    dry_run: bool = False


@dataclass
class EvaluationResult:
    output_register: Path
    backup: Path | None
    run_id: str
    total_findings: int
    true_positives: int
    false_positives: int
    unique_seeded_errors_matched: int
    seeded_errors_in_scope: int
    false_negatives: int
    precision: float | None
    recall: float | None
    changed_segments: int
    findings_rows: list[dict[str, Any]]
    evaluation_row: dict[str, Any]
    warnings: list[str]
    written: bool = True

    def summary(self) -> dict[str, Any]:
        return {
            "result": "completed" if self.written else "dry-run",
            "output_register": str(self.output_register),
            "backup": str(self.backup) if self.backup else None,
            "run_id": self.run_id,
            "total_findings": self.total_findings,
            "true_positives": self.true_positives,
            "false_positives": self.false_positives,
            "unique_seeded_errors_matched": self.unique_seeded_errors_matched,
            "seeded_errors_in_scope": self.seeded_errors_in_scope,
            "false_negatives": self.false_negatives,
            "precision": self.precision,
            "recall": self.recall,
            "changed_segments_in_csvs": self.changed_segments,
            "warnings": self.warnings,
        }


def default_username() -> str:
    try:
        return getpass.getuser()
    except (OSError, KeyError, ImportError):
        return ""


def safe_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, bool):
        return "Yes" if value else "No"
    if isinstance(value, float):
        if value != value or value in {float("inf"), float("-inf")}:
            return ""
    text = str(value)
    text = "".join(character for character in text if character in "\t\n\r" or ord(character) >= 0x20)
    return text[:32_767]


def normalise_header(value: Any) -> str:
    text = unicodedata.normalize("NFKD", safe_text(value).lstrip("\ufeff"))
    return "".join(character for character in text.casefold() if character.isalnum())


def clean_header(value: Any, index: int) -> str:
    text = safe_text(value).lstrip("\ufeff").strip()
    text = re.sub(r"[,;|\t]+$", "", text).strip()
    return text or f"Column_{index + 1}"


def normalise_text(value: Any) -> str:
    text = unicodedata.normalize("NFKC", safe_text(value)).replace("\u00a0", " ")
    text = text.replace("\u2018", "'").replace("\u2019", "'")
    text = text.replace("\u201c", '"').replace("\u201d", '"')
    return re.sub(r"\s+", " ", text).strip().casefold()


def normalise_segment(value: Any) -> str:
    text = normalise_text(value)
    if re.fullmatch(r"0*\d+", text):
        return str(int(text))
    return text


def parse_segment_tokens(value: Any, offset: int = 0) -> set[str]:
    text = safe_text(value).strip()
    if not text:
        return set()
    result: set[str] = set()
    for start, end in re.findall(r"(?<![\w.])(\d+)\s*[-–]\s*(\d+)(?![\w.])", text):
        lower, upper = int(start), int(end)
        if lower <= upper and upper - lower <= 1_000:
            result.update(str(number + offset) for number in range(lower, upper + 1))
    without_ranges = re.sub(r"(?<![\w.])\d+\s*[-–]\s*\d+(?![\w.])", " ", text)
    pieces = re.split(r"[,;|\s]+", without_ranges)
    for piece in pieces:
        piece = piece.strip("()[]{}:,.#")
        if not piece:
            continue
        if re.fullmatch(r"\d+", piece):
            result.add(str(int(piece) + offset))
        elif re.fullmatch(r"[A-Za-z]*\d+(?:[._-][A-Za-z0-9]+)*", piece):
            result.add(normalise_segment(piece))
    return result


def filename_core(value: Any) -> tuple[str, ...]:
    text = unicodedata.normalize("NFKD", Path(safe_text(value)).name)
    tokens = re.findall(r"[A-Za-z0-9]+", text)
    year_index = next((index for index, token in enumerate(tokens) if re.fullmatch(r"20\d{2}|19\d{2}", token)), None)
    start = max(0, year_index - 1) if year_index is not None else 0
    result: list[str] = []
    for token in tokens[start:]:
        folded = token.casefold()
        if folded in FILE_EXT_WORDS or folded in ROLE_WORDS:
            if len(result) >= 3:
                break
            continue
        result.append(folded)
    return tuple(result)


def filename_matches(left: Any, right: Any) -> bool:
    a, b = filename_core(left), filename_core(right)
    if not a or not b:
        return False
    if a == b:
        return True
    shorter, longer = (a, b) if len(a) <= len(b) else (b, a)
    if len(shorter) >= 4 and longer[: len(shorter)] == shorter:
        return True
    common = 0
    for x, y in zip(a, b):
        if x != y:
            break
        common += 1
    return common >= 6 and common >= min(len(a), len(b)) - 1


def file_sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def read_text_file(path: Path) -> tuple[str, str]:
    raw = path.read_bytes()
    for encoding in ("utf-8-sig", "utf-16", "utf-8", "cp1252"):
        try:
            return raw.decode(encoding), encoding
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="replace"), "utf-8 (replacement characters)"


def delimiter_score(header_line: str, delimiter: str) -> tuple[int, int, int]:
    try:
        fields = next(csv.reader([header_line], delimiter=delimiter))
    except (csv.Error, StopIteration):
        return (-1, -1, -1)
    normalised = {normalise_header(clean_header(value, index)) for index, value in enumerate(fields)}
    alias_values = set().union(*HEADER_ALIASES.values())
    known = sum(value in alias_values for value in normalised)
    structural = sum(value in HEADER_ALIASES["segment"] | HEADER_ALIASES["ori"] | HEADER_ALIASES["modified_tra"] for value in normalised)
    return (structural, known, min(len(fields), 30))


def detect_delimiter(text: str, suffix: str = "") -> str:
    header_line = next((line for line in text.splitlines() if line.strip()), "")
    if not header_line:
        return "\t" if suffix.casefold() == ".tsv" else ","
    candidates = ("|", "\t", ",", ";")
    return max(candidates, key=lambda delimiter: delimiter_score(header_line, delimiter))


def make_unique_headers(raw_headers: Sequence[Any]) -> list[str]:
    headers: list[str] = []
    counts: dict[str, int] = {}
    for index, raw in enumerate(raw_headers):
        base = clean_header(raw, index)
        key = base.casefold()
        counts[key] = counts.get(key, 0) + 1
        headers.append(base if counts[key] == 1 else f"{base}_{counts[key]}")
    return headers


def read_delimited(path: Path) -> tuple[list[str], list[dict[str, str]], str, str]:
    text, encoding = read_text_file(path)
    delimiter = detect_delimiter(text, path.suffix)
    reader = csv.reader(io.StringIO(text, newline=""), delimiter=delimiter)
    try:
        raw_headers = next(reader)
    except StopIteration:
        return [], [], delimiter, encoding
    headers = make_unique_headers(raw_headers)
    records: list[dict[str, str]] = []
    for row in reader:
        if not row or not any(safe_text(value).strip() for value in row):
            continue
        if len(row) < len(headers):
            row = list(row) + [""] * (len(headers) - len(row))
        elif len(row) > len(headers):
            extra = row[len(headers) - 1 :]
            row = list(row[: len(headers) - 1]) + [delimiter.join(extra)]
        records.append({header: safe_text(row[index]) for index, header in enumerate(headers)})
    return headers, records, delimiter, encoding


def resolve_column(headers: Sequence[str], role: str, requested: str | None = None) -> str | None:
    if requested and requested.casefold() not in {"auto", "none", "-"}:
        target = normalise_header(requested)
        for header in headers:
            if normalise_header(header) == target:
                return header
        raise EvaluationError(f"Column {requested!r} was not found for {role.replace('_', ' ')}.")
    if requested and requested.casefold() in {"none", "-"}:
        return None
    aliases = HEADER_ALIASES[role]
    for header in headers:
        if normalise_header(header) in aliases:
            return header
    for header in headers:
        normalised = normalise_header(header)
        if any(len(alias) >= 4 and normalised.endswith(alias) for alias in aliases):
            return header
    return None


def read_bilingual_csv(path: Path) -> DatasetFile:
    headers, records, delimiter, encoding = read_delimited(path)
    if not headers:
        raise EvaluationError(f"The CSV is empty: {path}")
    segment_column = resolve_column(headers, "segment")
    ori_column = resolve_column(headers, "ori")
    tra_column = resolve_column(headers, "modified_tra")
    if tra_column is None:
        raise EvaluationError(
            f"Could not find a TRA/target/translation column in {path.name}. "
            f"Detected columns: {', '.join(headers)}"
        )
    rows: list[DatasetRow] = []
    for index, record in enumerate(records, 1):
        segment = record.get(segment_column, "") if segment_column else str(index)
        rows.append(DatasetRow(
            segment=segment or str(index),
            ori=record.get(ori_column, "") if ori_column else "",
            tra=record.get(tra_column, ""),
            row_number=index,
        ))
    return DatasetFile(
        path=path,
        rows=rows,
        delimiter=delimiter,
        encoding=encoding,
        segment_column=segment_column,
        ori_column=ori_column,
        tra_column=tra_column,
        implicit_segments=segment_column is None,
    )


def flatten_json_record(record: dict[str, Any]) -> dict[str, str]:
    result: dict[str, str] = {}

    def visit(prefix: str, value: Any) -> None:
        if isinstance(value, dict):
            for key, child in value.items():
                visit(f"{prefix}.{key}" if prefix else str(key), child)
        elif isinstance(value, list):
            result[prefix or "value"] = json.dumps(value, ensure_ascii=False)
        else:
            result[prefix or "value"] = safe_text(value)

    visit("", record)
    return result


def records_from_json(path: Path) -> tuple[list[str], list[dict[str, str]], str]:
    text, encoding = read_text_file(path)
    try:
        payload = json.loads(text)
    except json.JSONDecodeError as exc:
        raise EvaluationError(f"Invalid JSON in {path.name}: {exc}") from exc
    if isinstance(payload, dict):
        candidates = [payload.get(key) for key in ("findings", "results", "issues", "errors", "items")]
        items = next((value for value in candidates if isinstance(value, list)), None)
        if items is None:
            items = [payload]
    elif isinstance(payload, list):
        items = payload
    else:
        raise EvaluationError("The JSON findings file must contain an object or a list of objects.")
    records = [flatten_json_record(item if isinstance(item, dict) else {"value": item}) for item in items]
    headers = list(dict.fromkeys(key for record in records for key in record))
    return headers, records, encoding


def local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def xml_record(element: ET.Element) -> dict[str, str]:
    result = {f"@{local_name(key)}": safe_text(value) for key, value in element.attrib.items()}
    for child in list(element):
        key = local_name(child.tag)
        value = "".join(child.itertext()).strip()
        if key in result and value:
            result[key] = f"{result[key]} | {value}"
        else:
            result[key] = value
    if not result:
        result[local_name(element.tag)] = "".join(element.itertext()).strip()
    return result


def records_from_xml(path: Path) -> tuple[list[str], list[dict[str, str]], str]:
    text, encoding = read_text_file(path)
    try:
        root = ET.fromstring(text)
    except ET.ParseError as exc:
        raise EvaluationError(f"Invalid XML in {path.name}: {exc}") from exc
    preferred = {"finding", "result", "issue", "error", "item"}
    elements = [element for element in root.iter() if local_name(element.tag).casefold() in preferred]
    if not elements:
        elements = list(root)
    records = [xml_record(element) for element in elements]
    headers = list(dict.fromkeys(key for record in records for key in record))
    return headers, records, encoding


def read_findings(path: Path) -> FindingsData:
    suffix = path.suffix.casefold()
    if suffix in {".csv", ".tsv"}:
        headers, records, delimiter, encoding = read_delimited(path)
        return FindingsData(path, records, headers, "CSV" if suffix == ".csv" else "TSV", delimiter, encoding)
    if suffix == ".json":
        headers, records, encoding = records_from_json(path)
        return FindingsData(path, records, headers, "JSON", encoding=encoding)
    if suffix == ".xml":
        headers, records, encoding = records_from_xml(path)
        return FindingsData(path, records, headers, "XML", encoding=encoding)
    raise EvaluationError(f"Unsupported findings format {suffix!r}. Use CSV, TSV, JSON, or XML.")


def xlsx_column_name(number: int) -> str:
    result = ""
    while number:
        number, remainder = divmod(number - 1, 26)
        result = chr(65 + remainder) + result
    return result


def xlsx_column_index(reference: str) -> int:
    result = 0
    for character in reference:
        if not character.isalpha():
            break
        result = result * 26 + ord(character.upper()) - 64
    return result


def xml_bytes(root: ET.Element) -> bytes:
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def xml_bytes_with_default(root: ET.Element, namespace: str) -> bytes:
    """Serialise package XML with the default namespace strict OpenXML readers expect."""
    ET.register_namespace("", namespace)
    try:
        return ET.tostring(root, encoding="utf-8", xml_declaration=True)
    finally:
        ET.register_namespace("", MAIN_NS)


def relationship_part(part: str) -> str:
    return posixpath.join(posixpath.dirname(part), "_rels", posixpath.basename(part) + ".rels")


def resolve_part(base_part: str, target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(posixpath.dirname(base_part), target))


def next_numeric_id(values: Iterable[str], prefix: str = "rId") -> str:
    numbers = []
    for value in values:
        match = re.fullmatch(re.escape(prefix) + r"(\d+)", value or "")
        if match:
            numbers.append(int(match.group(1)))
    return f"{prefix}{max(numbers, default=0) + 1}"


class XlsxRegister:
    """Minimal OOXML editor that preserves unrelated package parts byte-for-byte."""

    def __init__(self, path: Path):
        self.path = path
        self.entries: dict[str, bytes] = {}
        self.infos: dict[str, zipfile.ZipInfo] = {}
        try:
            with zipfile.ZipFile(path) as archive:
                if archive.testzip() is not None:
                    raise EvaluationError(f"The workbook ZIP package is damaged: {path}")
                for info in archive.infolist():
                    self.entries[info.filename] = archive.read(info.filename)
                    self.infos[info.filename] = info
        except (OSError, zipfile.BadZipFile) as exc:
            raise EvaluationError(f"Could not open register workbook {path}: {exc}") from exc
        for required in ("[Content_Types].xml", "xl/workbook.xml", "xl/_rels/workbook.xml.rels"):
            if required not in self.entries:
                raise EvaluationError(f"The register is missing required XLSX part {required}.")
        self.shared_strings = self._read_shared_strings()
        self.data_style_id = self._detect_data_style_id()

    def _read_shared_strings(self) -> list[str]:
        payload = self.entries.get("xl/sharedStrings.xml")
        if not payload:
            return []
        root = ET.fromstring(payload)
        return ["".join(node.text or "" for node in item.iter() if local_name(node.tag) == "t") for item in root]

    def _detect_data_style_id(self) -> int:
        payload = self.entries.get("xl/styles.xml")
        if not payload:
            return 0
        root = ET.fromstring(payload)
        cell_xfs = root.find(f"{{{MAIN_NS}}}cellXfs")
        count = int(cell_xfs.get("count", "0")) if cell_xfs is not None else 0
        return 1 if count > 1 else 0

    def workbook_roots(self) -> tuple[ET.Element, ET.Element]:
        return ET.fromstring(self.entries["xl/workbook.xml"]), ET.fromstring(self.entries["xl/_rels/workbook.xml.rels"])

    def sheet_map(self) -> dict[str, tuple[ET.Element, str]]:
        workbook, relationships = self.workbook_roots()
        targets = {node.get("Id", ""): node.get("Target", "") for node in relationships}
        result: dict[str, tuple[ET.Element, str]] = {}
        sheets = workbook.find(f"{{{MAIN_NS}}}sheets")
        if sheets is None:
            return result
        for sheet in sheets:
            rid = sheet.get(f"{{{DOC_REL_NS}}}id", "")
            target = targets.get(rid, "")
            part = resolve_part("xl/workbook.xml", target)
            result[sheet.get("name", "").casefold()] = (sheet, part)
        return result

    def cell_value(self, cell: ET.Element) -> str:
        kind = cell.get("t", "")
        if kind == "inlineStr":
            return "".join(node.text or "" for node in cell.iter() if local_name(node.tag) == "t")
        value = cell.find(f"{{{MAIN_NS}}}v")
        raw = value.text if value is not None and value.text is not None else ""
        if kind == "s" and raw:
            try:
                return self.shared_strings[int(raw)]
            except (IndexError, ValueError):
                return raw
        if kind == "b":
            return "Yes" if raw == "1" else "No"
        return raw

    def sheet_rows(self, sheet_name: str) -> list[list[str]]:
        item = self.sheet_map().get(sheet_name.casefold())
        if item is None:
            return []
        _sheet, part = item
        root = ET.fromstring(self.entries[part])
        rows: list[list[str]] = []
        for row in root.findall(f".//{{{MAIN_NS}}}sheetData/{{{MAIN_NS}}}row"):
            values: list[str] = []
            for cell in row.findall(f"{{{MAIN_NS}}}c"):
                index = xlsx_column_index(cell.get("r", "A1")) - 1
                while len(values) <= index:
                    values.append("")
                values[index] = self.cell_value(cell)
            while values and values[-1] == "":
                values.pop()
            rows.append(values)
        return rows

    def sheet_records(self, sheet_name: str) -> list[dict[str, str]]:
        rows = self.sheet_rows(sheet_name)
        if not rows:
            return []
        headers = make_unique_headers(rows[0])
        records: list[dict[str, str]] = []
        for row in rows[1:]:
            values = list(row) + [""] * max(0, len(headers) - len(row))
            records.append({header: values[index] for index, header in enumerate(headers)})
        return records

    def _new_cell(self, reference: str, value: Any) -> ET.Element:
        cell = ET.Element(f"{{{MAIN_NS}}}c", {"r": reference})
        if self.data_style_id:
            cell.set("s", str(self.data_style_id))
        if value is None or value == "":
            return cell
        if isinstance(value, bool):
            value = "Yes" if value else "No"
        if isinstance(value, int) and not isinstance(value, bool):
            child = ET.SubElement(cell, f"{{{MAIN_NS}}}v")
            child.text = str(value)
            return cell
        if isinstance(value, float) and value == value and value not in {float("inf"), float("-inf")}:
            child = ET.SubElement(cell, f"{{{MAIN_NS}}}v")
            child.text = format(value, ".15g")
            return cell
        cell.set("t", "inlineStr")
        inline = ET.SubElement(cell, f"{{{MAIN_NS}}}is")
        text = ET.SubElement(inline, f"{{{MAIN_NS}}}t", {"{http://www.w3.org/XML/1998/namespace}space": "preserve"})
        text.text = safe_text(value)
        return cell

    def _new_row(self, row_number: int, values: Sequence[Any]) -> ET.Element:
        row = ET.Element(f"{{{MAIN_NS}}}row", {"r": str(row_number), "spans": f"1:{len(values)}"})
        for column_number, value in enumerate(values, 1):
            row.append(self._new_cell(f"{xlsx_column_name(column_number)}{row_number}", value))
        return row

    def _unique_table_name(self, preferred: str) -> str:
        existing: set[str] = set()
        for name, payload in self.entries.items():
            if name.startswith("xl/tables/") and name.endswith(".xml"):
                try:
                    root = ET.fromstring(payload)
                    existing.add((root.get("displayName") or root.get("name") or "").casefold())
                except ET.ParseError:
                    continue
        candidate = preferred
        index = 2
        while candidate.casefold() in existing:
            candidate = f"{preferred}{index}"
            index += 1
        return candidate

    def _next_part_number(self, folder: str, stem: str) -> int:
        pattern = re.compile(rf"^{re.escape(folder)}/{re.escape(stem)}(\d+)\.xml$")
        numbers = [int(match.group(1)) for name in self.entries if (match := pattern.match(name))]
        return max(numbers, default=0) + 1

    def _next_table_id(self) -> int:
        ids: list[int] = []
        for name, payload in self.entries.items():
            if name.startswith("xl/tables/") and name.endswith(".xml"):
                try:
                    value = ET.fromstring(payload).get("id")
                    if value:
                        ids.append(int(value))
                except (ET.ParseError, ValueError):
                    pass
        return max(ids, default=0) + 1

    def _ensure_override(self, part_name: str, content_type: str) -> None:
        root = ET.fromstring(self.entries["[Content_Types].xml"])
        for node in root:
            if node.get("PartName", "").casefold() == part_name.casefold():
                return
        ET.SubElement(root, f"{{{CONTENT_TYPES_NS}}}Override", {"PartName": part_name, "ContentType": content_type})
        self.entries["[Content_Types].xml"] = xml_bytes_with_default(root, CONTENT_TYPES_NS)

    def _add_sheet(self, sheet_name: str, columns: Sequence[str], widths: Sequence[int], table_name: str, tab_color: str) -> str:
        workbook, relationships = self.workbook_roots()
        sheet_number = self._next_part_number("xl/worksheets", "sheet")
        sheet_part = f"xl/worksheets/sheet{sheet_number}.xml"
        table_number = self._next_part_number("xl/tables", "table")
        table_part = f"xl/tables/table{table_number}.xml"
        table_id = self._next_table_id()
        relation_ids = [node.get("Id", "") for node in relationships]
        workbook_rid = next_numeric_id(relation_ids)
        ET.SubElement(relationships, f"{{{PKG_REL_NS}}}Relationship", {
            "Id": workbook_rid,
            "Type": f"{DOC_REL_NS}/worksheet",
            "Target": f"worksheets/sheet{sheet_number}.xml",
        })
        sheets = workbook.find(f"{{{MAIN_NS}}}sheets")
        if sheets is None:
            sheets = ET.SubElement(workbook, f"{{{MAIN_NS}}}sheets")
        sheet_ids = [int(node.get("sheetId", "0")) for node in sheets]
        ET.SubElement(sheets, f"{{{MAIN_NS}}}sheet", {
            "name": sheet_name,
            "sheetId": str(max(sheet_ids, default=0) + 1),
            f"{{{DOC_REL_NS}}}id": workbook_rid,
        })
        self.entries["xl/workbook.xml"] = xml_bytes(workbook)
        self.entries["xl/_rels/workbook.xml.rels"] = xml_bytes_with_default(relationships, PKG_REL_NS)

        last_column = xlsx_column_name(len(columns))
        table_ref = f"A1:{last_column}1"
        worksheet = ET.Element(f"{{{MAIN_NS}}}worksheet")
        sheet_pr = ET.SubElement(worksheet, f"{{{MAIN_NS}}}sheetPr")
        ET.SubElement(sheet_pr, f"{{{MAIN_NS}}}tabColor", {"rgb": "FF" + tab_color.lstrip("#").upper()})
        ET.SubElement(worksheet, f"{{{MAIN_NS}}}dimension", {"ref": table_ref})
        views = ET.SubElement(worksheet, f"{{{MAIN_NS}}}sheetViews")
        view = ET.SubElement(views, f"{{{MAIN_NS}}}sheetView", {"workbookViewId": "0", "showGridLines": "0"})
        ET.SubElement(view, f"{{{MAIN_NS}}}pane", {"ySplit": "1", "topLeftCell": "A2", "activePane": "bottomLeft", "state": "frozen"})
        ET.SubElement(worksheet, f"{{{MAIN_NS}}}sheetFormatPr", {"defaultRowHeight": "15"})
        cols = ET.SubElement(worksheet, f"{{{MAIN_NS}}}cols")
        for index, width in enumerate(widths, 1):
            ET.SubElement(cols, f"{{{MAIN_NS}}}col", {"min": str(index), "max": str(index), "width": str(width), "customWidth": "1"})
        sheet_data = ET.SubElement(worksheet, f"{{{MAIN_NS}}}sheetData")
        sheet_data.append(self._new_row(1, columns))
        ET.SubElement(worksheet, f"{{{MAIN_NS}}}pageMargins", {
            "left": "0.7", "right": "0.7", "top": "0.75", "bottom": "0.75", "header": "0.3", "footer": "0.3",
        })
        table_parts = ET.SubElement(worksheet, f"{{{MAIN_NS}}}tableParts", {"count": "1"})
        ET.SubElement(table_parts, f"{{{MAIN_NS}}}tablePart", {f"{{{DOC_REL_NS}}}id": "rId1"})
        self.entries[sheet_part] = xml_bytes(worksheet)

        sheet_relationships = ET.Element(f"{{{PKG_REL_NS}}}Relationships")
        ET.SubElement(sheet_relationships, f"{{{PKG_REL_NS}}}Relationship", {
            "Id": "rId1", "Type": f"{DOC_REL_NS}/table", "Target": f"../tables/table{table_number}.xml",
        })
        self.entries[relationship_part(sheet_part)] = xml_bytes_with_default(sheet_relationships, PKG_REL_NS)

        display_name = self._unique_table_name(table_name)
        table = ET.Element(f"{{{MAIN_NS}}}table", {
            "id": str(table_id), "name": display_name, "displayName": display_name,
            "ref": table_ref, "headerRowCount": "1", "totalsRowCount": "0", "totalsRowShown": "0",
        })
        ET.SubElement(table, f"{{{MAIN_NS}}}autoFilter", {"ref": table_ref})
        table_columns = ET.SubElement(table, f"{{{MAIN_NS}}}tableColumns", {"count": str(len(columns))})
        for index, column in enumerate(columns, 1):
            ET.SubElement(table_columns, f"{{{MAIN_NS}}}tableColumn", {"id": str(index), "name": column})
        ET.SubElement(table, f"{{{MAIN_NS}}}tableStyleInfo", {
            "name": "TableStyleMedium12", "showFirstColumn": "0", "showLastColumn": "0",
            "showRowStripes": "1", "showColumnStripes": "0",
        })
        self.entries[table_part] = xml_bytes(table)
        self._ensure_override(
            f"/{sheet_part}",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
        )
        self._ensure_override(
            f"/{table_part}",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml",
        )
        return sheet_part

    def _table_for_sheet(self, sheet_part: str) -> tuple[str, ET.Element]:
        rel_part = relationship_part(sheet_part)
        if rel_part not in self.entries:
            raise EvaluationError(f"Worksheet {sheet_part} has no relationships part for its table.")
        relationships = ET.fromstring(self.entries[rel_part])
        relation = next((node for node in relationships if (node.get("Type") or "").endswith("/table")), None)
        if relation is None:
            raise EvaluationError(f"Worksheet {sheet_part} is missing its Excel table relationship.")
        table_part = resolve_part(sheet_part, relation.get("Target", ""))
        if table_part not in self.entries:
            raise EvaluationError(f"Worksheet {sheet_part} points to missing table part {table_part}.")
        return table_part, ET.fromstring(self.entries[table_part])

    def append_rows(self, sheet_name: str, columns: Sequence[str], rows: Sequence[dict[str, Any]], widths: Sequence[int], table_name: str, tab_color: str) -> None:
        item = self.sheet_map().get(sheet_name.casefold())
        sheet_part = item[1] if item else self._add_sheet(sheet_name, columns, widths, table_name, tab_color)
        worksheet = ET.fromstring(self.entries[sheet_part])
        sheet_data = worksheet.find(f"{{{MAIN_NS}}}sheetData")
        if sheet_data is None:
            sheet_data = ET.SubElement(worksheet, f"{{{MAIN_NS}}}sheetData")
        existing_rows = sheet_data.findall(f"{{{MAIN_NS}}}row")
        if not existing_rows:
            sheet_data.append(self._new_row(1, columns))
            existing_rows = sheet_data.findall(f"{{{MAIN_NS}}}row")
        header_values = []
        for cell in existing_rows[0].findall(f"{{{MAIN_NS}}}c"):
            header_values.append(self.cell_value(cell))
        if tuple(header_values) != tuple(columns):
            raise EvaluationError(
                f"Existing worksheet {sheet_name!r} has an incompatible header. "
                "Rename it or restore the evaluator-created columns before appending."
            )
        last_row = max((int(row.get("r", "0")) for row in existing_rows), default=1)
        for record in rows:
            last_row += 1
            sheet_data.append(self._new_row(last_row, [record.get(column, "") for column in columns]))
        last_column = xlsx_column_name(len(columns))
        reference = f"A1:{last_column}{max(1, last_row)}"
        dimension = worksheet.find(f"{{{MAIN_NS}}}dimension")
        if dimension is None:
            dimension = ET.Element(f"{{{MAIN_NS}}}dimension")
            worksheet.insert(0, dimension)
        dimension.set("ref", reference)
        table_part, table = self._table_for_sheet(sheet_part)
        table.set("ref", reference)
        auto_filter = table.find(f"{{{MAIN_NS}}}autoFilter")
        if auto_filter is not None:
            auto_filter.set("ref", reference)
        self.entries[sheet_part] = xml_bytes(worksheet)
        self.entries[table_part] = xml_bytes(table)

    def save_atomic(self, destination: Path, backup_source: Path | None = None) -> Path | None:
        destination.parent.mkdir(parents=True, exist_ok=True)
        temporary = destination.with_name(f"{destination.name}.{os.getpid()}.{threading.get_ident()}.tmp")
        backup: Path | None = None
        try:
            with zipfile.ZipFile(temporary, "w") as archive:
                for name, payload in self.entries.items():
                    if name in self.infos:
                        info = self.infos[name]
                        info.compress_type = zipfile.ZIP_DEFLATED
                        archive.writestr(info, payload)
                    else:
                        archive.writestr(name, payload, compress_type=zipfile.ZIP_DEFLATED)
            with zipfile.ZipFile(temporary) as validation:
                damaged = validation.testzip()
                if damaged is not None:
                    raise EvaluationError(f"Generated workbook is damaged at {damaged}.")
                ET.fromstring(validation.read("xl/workbook.xml"))
            if backup_source is not None and backup_source.exists():
                stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
                backup = backup_source.with_name(f"{backup_source.stem}.backup-{stamp}{backup_source.suffix}")
                counter = 2
                while backup.exists():
                    backup = backup_source.with_name(f"{backup_source.stem}.backup-{stamp}-{counter}{backup_source.suffix}")
                    counter += 1
                shutil.copy2(backup_source, backup)
            os.replace(temporary, destination)
            return backup
        except PermissionError as exc:
            raise EvaluationError(
                f"Could not write {destination}. Close the workbook in Excel and try again."
            ) from exc
        finally:
            try:
                temporary.unlink(missing_ok=True)
            except OSError:
                pass


def record_value(record: dict[str, str], column: str | None) -> str:
    return record.get(column, "") if column else ""


def workbook_column(record: dict[str, str], role: str) -> str:
    aliases = HEADER_ALIASES[role]
    for key, value in record.items():
        if normalise_header(key) in aliases:
            return safe_text(value)
    return ""


def seeded_errors_from_register(register: XlsxRegister) -> list[SeededError]:
    rows = register.sheet_records("Seeded Errors")
    if not rows:
        raise EvaluationError("The register does not contain a populated 'Seeded Errors' worksheet.")
    result = []
    for record in rows:
        seeded = SeededError(
            error_id=workbook_column(record, "error_id"),
            tester=workbook_column(record, "tester"),
            filename=workbook_column(record, "filename"),
            segment=workbook_column(record, "segment"),
            ori=workbook_column(record, "ori"),
            correct_tra=workbook_column(record, "correct_tra"),
            modified_tra=workbook_column(record, "modified_tra"),
            comment=workbook_column(record, "comment"),
            changed_words=workbook_column(record, "changed_words"),
            similar_segments=workbook_column(record, "similar_segments"),
        )
        if seeded.substantive:
            result.append(seeded)
    return result


def join_dataset_values(lookup: dict[str, list[DatasetRow]], segments: set[str], attribute: str) -> str:
    values: list[str] = []
    for segment in sorted(segments, key=lambda value: (not value.isdigit(), int(value) if value.isdigit() else value)):
        for row in lookup.get(segment, []):
            value = safe_text(getattr(row, attribute))
            if value and value not in values:
                values.append(value)
    return "\n---\n".join(values)


def changed_segment_keys(golden: DatasetFile, modified: DatasetFile) -> set[str]:
    golden_lookup, modified_lookup = golden.by_segment(), modified.by_segment()
    result: set[str] = set()
    for segment in set(golden_lookup) | set(modified_lookup):
        golden_rows = golden_lookup.get(segment, [])
        modified_rows = modified_lookup.get(segment, [])
        golden_value = "\n".join(normalise_text(row.tra) for row in golden_rows)
        modified_value = "\n".join(normalise_text(row.tra) for row in modified_rows)
        if golden_value != modified_value:
            result.add(segment)
    return result


def row_text_matches_seed(seed: SeededError, segments: set[str], modified_lookup: dict[str, list[DatasetRow]], ori: str, modified_tra: str) -> tuple[bool, bool]:
    candidate_modified = normalise_text(modified_tra)
    candidate_ori = normalise_text(ori)
    seed_modified = normalise_text(seed.modified_tra)
    seed_ori = normalise_text(seed.ori)
    modified_match = bool(candidate_modified and seed_modified and candidate_modified == seed_modified)
    ori_match = bool(candidate_ori and seed_ori and candidate_ori == seed_ori)
    if not modified_match and segments and seed_modified:
        modified_match = any(
            normalise_text(row.tra) == seed_modified
            for segment in segments
            for row in modified_lookup.get(segment, [])
        )
    if not ori_match and segments and seed_ori:
        ori_match = any(
            normalise_text(row.ori) == seed_ori
            for segment in segments
            for row in modified_lookup.get(segment, [])
        )
    return modified_match, ori_match


def match_finding(
    seed: SeededError,
    *,
    finding_filename: str,
    default_filename: str,
    segments: set[str],
    ori: str,
    modified_tra: str,
    modified_lookup: dict[str, list[DatasetRow]],
) -> str | None:
    file_match = filename_matches(finding_filename or default_filename, seed.filename) or filename_matches(default_filename, seed.filename)
    segment_match = bool(segments and seed.segments and segments.intersection(seed.segments))
    modified_match, ori_match = row_text_matches_seed(seed, segments, modified_lookup, ori, modified_tra)
    if segment_match and file_match:
        return "filename+segment"
    if segment_match and modified_match:
        return "segment+modified_text"
    if not seed.segments and modified_match and (file_match or ori_match):
        return "modified_text" + ("+filename" if file_match else "+ori")
    if not segments and modified_match and file_match:
        return "filename+modified_text"
    return None


def seeded_scope(
    seeds: Sequence[SeededError],
    modified: DatasetFile,
) -> list[SeededError]:
    lookup = modified.by_segment()
    result: list[SeededError] = []
    for seed in seeds:
        if filename_matches(modified.path.name, seed.filename):
            result.append(seed)
            continue
        if not seed.segments:
            seed_modified = normalise_text(seed.modified_tra)
            seed_ori = normalise_text(seed.ori)
            if any(
                (seed_modified and normalise_text(row.tra) == seed_modified)
                or (seed_ori and normalise_text(row.ori) == seed_ori)
                for rows in lookup.values() for row in rows
            ):
                result.append(seed)
            continue
        for segment in seed.segments:
            rows = lookup.get(segment, [])
            if any(
                (seed.modified_tra and normalise_text(row.tra) == normalise_text(seed.modified_tra))
                or (seed.ori and normalise_text(row.ori) == normalise_text(seed.ori))
                for row in rows
            ):
                result.append(seed)
                break
    return result


def resolved_mapping(data: FindingsData, requested: ColumnMapping) -> ColumnMapping:
    headers = data.headers
    mapping = ColumnMapping(
        segment=resolve_column(headers, "segment", requested.segment),
        filename=resolve_column(headers, "filename", requested.filename),
        message=resolve_column(headers, "comment", requested.message),
        severity=resolve_column(headers, "severity", requested.severity),
        finding_id=resolve_column(headers, "finding_id", requested.finding_id),
        ori=resolve_column(headers, "ori", requested.ori),
        correct_tra=resolve_column(headers, "correct_tra", requested.correct_tra),
        modified_tra=resolve_column(headers, "modified_tra", requested.modified_tra),
    )
    if data.records and mapping.segment is None and mapping.modified_tra is None and mapping.ori is None:
        raise EvaluationError(
            "Could not identify a segment, source-text, or target-text column in the findings file. "
            f"Detected columns: {', '.join(headers) or '(none)'}. Use the advanced column mappings."
        )
    return mapping


def mapping_note(mapping: ColumnMapping) -> str:
    parts = []
    for label in ("segment", "filename", "message", "severity", "finding_id", "ori", "correct_tra", "modified_tra"):
        value = getattr(mapping, label)
        if value:
            parts.append(f"{label}={value}")
    return ", ".join(parts) or "no columns mapped (empty findings file)"


def validate_options(options: EvaluationOptions) -> None:
    for label, path in (
        ("golden CSV", options.golden),
        ("modified CSV", options.modified),
        ("register", options.register),
        ("findings file", options.findings),
    ):
        if not path.exists() or not path.is_file():
            raise EvaluationError(f"Selected {label} does not exist: {path}")
    if options.golden.suffix.casefold() not in {".csv", ".tsv"}:
        raise EvaluationError("The golden-standard input must be CSV or TSV.")
    if options.modified.suffix.casefold() not in {".csv", ".tsv"}:
        raise EvaluationError("The modified input must be CSV or TSV.")
    if options.register.suffix.casefold() != ".xlsx":
        raise EvaluationError("The dataset register must be an .xlsx workbook.")
    if options.findings.suffix.casefold() not in SUPPORTED_FINDINGS_EXTENSIONS:
        raise EvaluationError("The findings file must be CSV, TSV, JSON, or XML.")
    if options.output_register is not None and options.output_register.suffix.casefold() != ".xlsx":
        raise EvaluationError("The output register must use the .xlsx extension.")


def make_run_id(processed_at: str, findings_hash: str) -> str:
    compact = re.sub(r"\D", "", processed_at)[:17]
    return f"run-{compact}-{findings_hash[:8]}"


def default_output_path(register: Path) -> Path:
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    return register.with_name(f"{register.stem}-evaluated-{stamp}{register.suffix}")


def evaluate(options: EvaluationOptions) -> EvaluationResult:
    validate_options(options)
    options.golden = options.golden.resolve()
    options.modified = options.modified.resolve()
    options.register = options.register.resolve()
    options.findings = options.findings.resolve()
    if options.output_register is not None:
        options.output_register = options.output_register.resolve()

    golden = read_bilingual_csv(options.golden)
    modified = read_bilingual_csv(options.modified)
    findings = read_findings(options.findings)
    mapping = resolved_mapping(findings, options.mappings)
    register = XlsxRegister(options.register)
    seeds = seeded_errors_from_register(register)
    findings_hash = file_sha256(options.findings)

    prior_evaluations = register.sheet_records(EVALUATION_SHEET)
    prior_hashes: set[str] = set()
    for record in prior_evaluations:
        for key, value in record.items():
            if normalise_header(key) == normalise_header("Findings_SHA256") and value:
                prior_hashes.add(value)
    if findings_hash in prior_hashes and not options.allow_duplicate:
        raise DuplicateRunError(
            "This findings file content is already recorded in the evaluation sheet. "
            "Use --allow-duplicate only if this is an intentional repeated run."
        )

    processed_at = datetime.now().astimezone().isoformat(timespec="milliseconds")
    run_id = make_run_id(processed_at, findings_hash)
    golden_lookup = golden.by_segment()
    modified_lookup = modified.by_segment()
    changed_segments = changed_segment_keys(golden, modified)
    scope = seeded_scope(seeds, modified)
    scope_ids = {seed.error_id or f"row-{index}" for index, seed in enumerate(scope, 1)}
    warnings: list[str] = []
    if not seeds:
        warnings.append(
            "The register contains no substantive seeded-error rows; findings cannot match ground truth and are counted as false positives."
        )
    if golden.implicit_segments != modified.implicit_segments:
        warnings.append("Only one input CSV has an explicit segment column; alignment uses each file's available segment keys.")
    if not scope:
        warnings.append("No seeded errors could be placed in scope for the selected modified file by filename or exact segment text.")
    if not findings.records:
        warnings.append("The findings file contains no data rows.")

    detailed_rows: list[dict[str, Any]] = []
    matched_scope_ids: set[str] = set()
    true_positives = 0
    tester = options.tester.strip() or default_username()
    tool_name = options.tool_name.strip() or options.findings.stem

    for index, record in enumerate(findings.records, 1):
        segment_raw = record_value(record, mapping.segment)
        segments = parse_segment_tokens(segment_raw, options.segment_offset)
        adjusted_segment = ", ".join(sorted(segments, key=lambda value: (not value.isdigit(), int(value) if value.isdigit() else value)))
        if not adjusted_segment:
            adjusted_segment = segment_raw
        finding_filename = record_value(record, mapping.filename) or options.modified.name
        ori = join_dataset_values(modified_lookup, segments, "ori") or join_dataset_values(golden_lookup, segments, "ori") or record_value(record, mapping.ori)
        correct_tra = join_dataset_values(golden_lookup, segments, "tra") or record_value(record, mapping.correct_tra)
        modified_tra = join_dataset_values(modified_lookup, segments, "tra") or record_value(record, mapping.modified_tra)
        matches: list[tuple[SeededError, str]] = []
        for seed in seeds:
            method = match_finding(
                seed,
                finding_filename=finding_filename,
                default_filename=options.modified.name,
                segments=segments,
                ori=ori,
                modified_tra=modified_tra,
                modified_lookup=modified_lookup,
            )
            if method:
                matches.append((seed, method))
        if matches:
            true_positives += 1
        for seed, _method in matches:
            identity = seed.error_id or f"row-{seeds.index(seed) + 1}"
            if identity in scope_ids:
                matched_scope_ids.add(identity)

        def joined(attribute: str) -> str:
            values: list[str] = []
            for seed, _method in matches:
                value = safe_text(getattr(seed, attribute))
                if value and value not in values:
                    values.append(value)
            return " | ".join(values)

        match_ids = [seed.error_id or f"row-{seeds.index(seed) + 1}" for seed, _method in matches]
        methods = list(dict.fromkeys(method for _seed, method in matches))
        finding_message = record_value(record, mapping.message)
        if not finding_message:
            finding_message = next((safe_text(value) for key, value in record.items() if key not in {mapping.segment, mapping.filename} and safe_text(value).strip()), "")
        detailed_rows.append({
            "Run_ID": run_id,
            "Finding_ID": record_value(record, mapping.finding_id) or index,
            "Tester": tester,
            "Tool": tool_name,
            "Filename": finding_filename,
            "Segment_number": adjusted_segment,
            "ORI": ori,
            "Correct_TRA": correct_tra,
            "Modified_TRA": modified_tra,
            "Finding": finding_message,
            "Severity": record_value(record, mapping.severity),
            "Source_details": json.dumps(record, ensure_ascii=False, sort_keys=True),
            "Matches_seeded_error": "Yes" if matches else "No",
            "Matched_Error_IDs": ", ".join(dict.fromkeys(match_ids)),
            "Match_method": ", ".join(methods),
            "Seeded_comment": joined("comment"),
            "Seeded_changed_words": joined("changed_words"),
            "Seeded_similar_segments": joined("similar_segments"),
            "Findings_file": options.findings.name,
            "Processed_at": processed_at,
        })

    total = len(detailed_rows)
    false_positives = total - true_positives
    false_negatives = max(len(scope_ids) - len(matched_scope_ids), 0)
    precision = true_positives / total if total else None
    recall = len(matched_scope_ids) / len(scope_ids) if scope_ids else None
    notes = [
        f"Findings format: {findings.format_name}",
        f"Mapping: {mapping_note(mapping)}",
        f"Golden rows: {len(golden.rows)}",
        f"Modified rows: {len(modified.rows)}",
        "Matching is exact after Unicode, quote, case, and whitespace normalisation; segment ranges are expanded.",
    ]
    notes.extend(warnings)
    evaluation_row: dict[str, Any] = {
        "Run_ID": run_id,
        "Processed_at": processed_at,
        "Findings_file": options.findings.name,
        "Findings_SHA256": findings_hash,
        "Tool": tool_name,
        "Golden_file": options.golden.name,
        "Modified_file": options.modified.name,
        "Total_findings": total,
        "True_positives": true_positives,
        "False_positives": false_positives,
        "Unique_seeded_errors_matched": len(matched_scope_ids),
        "Seeded_errors_in_scope": len(scope_ids),
        "False_negatives": false_negatives,
        "Precision": precision,
        "Recall": recall,
        "Changed_segments_in_CSVs": len(changed_segments),
        "Notes": "\n".join(notes),
    }

    destination = options.register if options.update_in_place and options.output_register is None else (options.output_register or default_output_path(options.register))
    destination = destination.resolve()
    backup: Path | None = None
    if not options.dry_run:
        register.append_rows(
            FINDINGS_SHEET, FINDINGS_COLUMNS, detailed_rows,
            (20, 12, 16, 24, 48, 18, 50, 50, 50, 48, 14, 55, 22, 24, 24, 45, 32, 32, 42, 25),
            "FindingsTable", "0F6B78",
        )
        register.append_rows(
            EVALUATION_SHEET, EVALUATION_COLUMNS, [evaluation_row],
            (20, 25, 42, 68, 24, 42, 42, 16, 16, 16, 24, 22, 16, 14, 14, 24, 70),
            "EvaluationTable", "1F4E78",
        )
        backup_source = options.register if destination == options.register and options.create_backup else None
        backup = register.save_atomic(destination, backup_source)

    return EvaluationResult(
        output_register=destination,
        backup=backup,
        run_id=run_id,
        total_findings=total,
        true_positives=true_positives,
        false_positives=false_positives,
        unique_seeded_errors_matched=len(matched_scope_ids),
        seeded_errors_in_scope=len(scope_ids),
        false_negatives=false_negatives,
        precision=precision,
        recall=recall,
        changed_segments=len(changed_segments),
        findings_rows=detailed_rows,
        evaluation_row=evaluation_row,
        warnings=warnings,
        written=not options.dry_run,
    )


def load_settings() -> dict[str, Any]:
    try:
        payload = json.loads(SETTINGS_PATH.read_text(encoding="utf-8"))
        return payload if payload.get("schema_version") == SETTINGS_SCHEMA_VERSION else {}
    except (OSError, json.JSONDecodeError, AttributeError):
        return {}


def save_settings(settings: dict[str, Any]) -> None:
    temporary = SETTINGS_PATH.with_name(f"{SETTINGS_PATH.name}.{os.getpid()}.tmp")
    try:
        temporary.write_text(json.dumps(settings, ensure_ascii=False, indent=2), encoding="utf-8")
        os.replace(temporary, SETTINGS_PATH)
    except OSError:
        pass
    finally:
        try:
            temporary.unlink(missing_ok=True)
        except OSError:
            pass


class EvaluatorApp:
    def __init__(self, root: "tk.Tk"):
        self.root = root
        self.settings = load_settings()
        self.events: queue.Queue[tuple[str, Any]] = queue.Queue()
        self.worker: threading.Thread | None = None
        root.title(f"{APP_NAME} {APP_VERSION}")
        root.geometry(self.settings.get("geometry", "980x820"))
        root.minsize(820, 680)
        root.protocol("WM_DELETE_WINDOW", self.on_close)

        self.path_vars = {
            "golden": tk.StringVar(value=""),
            "modified": tk.StringVar(value=""),
            "register": tk.StringVar(value=""),
            "findings": tk.StringVar(value=""),
            "output": tk.StringVar(value=""),
        }
        self.update_in_place = tk.BooleanVar(value=self.settings.get("update_in_place", True))
        self.create_backup = tk.BooleanVar(value=self.settings.get("create_backup", True))
        self.tester = tk.StringVar(value=self.settings.get("tester", default_username()))
        self.tool_name = tk.StringVar(value="")
        self.segment_offset = tk.StringVar(value="0")
        self.mapping_vars = {
            "segment": tk.StringVar(value="auto"),
            "filename": tk.StringVar(value="auto"),
            "message": tk.StringVar(value="auto"),
            "severity": tk.StringVar(value="auto"),
            "finding_id": tk.StringVar(value="auto"),
            "ori": tk.StringVar(value="auto"),
            "correct_tra": tk.StringVar(value="auto"),
            "modified_tra": tk.StringVar(value="auto"),
        }
        self.status = tk.StringVar(value="Ready — select the four input files.")
        self.build_ui()
        self.toggle_output_state()
        root.after(100, self.poll_events)

    def build_ui(self) -> None:
        menubar = tk.Menu(self.root)
        help_menu = tk.Menu(menubar, tearoff=False)
        help_menu.add_command(label="About and matching details", command=self.show_about)
        menubar.add_cascade(label="Help", menu=help_menu)
        self.root.configure(menu=menubar)

        main = ttk.Frame(self.root, padding=12)
        main.pack(fill="both", expand=True)
        main.columnconfigure(0, weight=1)
        main.rowconfigure(3, weight=1)

        files = ttk.LabelFrame(main, text="1. Input files", padding=10)
        files.grid(row=0, column=0, sticky="ew")
        files.columnconfigure(1, weight=1)
        definitions = (
            ("Golden-standard CSV", "golden", (("CSV files", "*.csv *.tsv"), ("All files", "*.*"))),
            ("Modified CSV", "modified", (("CSV files", "*.csv *.tsv"), ("All files", "*.*"))),
            ("Dataset register", "register", (("Excel workbooks", "*.xlsx"),)),
            ("Tool findings", "findings", (("Findings files", "*.csv *.tsv *.json *.xml"), ("All files", "*.*"))),
        )
        for row, (label, key, filetypes) in enumerate(definitions):
            ttk.Label(files, text=label).grid(row=row, column=0, sticky="w", padx=(0, 8), pady=4)
            ttk.Entry(files, textvariable=self.path_vars[key]).grid(row=row, column=1, sticky="ew", pady=4)
            ttk.Button(files, text="Browse…", command=lambda k=key, f=filetypes: self.browse_file(k, f)).grid(row=row, column=2, padx=(8, 0), pady=4)

        options = ttk.LabelFrame(main, text="2. Run details and output", padding=10)
        options.grid(row=1, column=0, sticky="ew", pady=(10, 0))
        options.columnconfigure(1, weight=1)
        ttk.Label(options, text="Tester").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=4)
        ttk.Entry(options, textvariable=self.tester, width=28).grid(row=0, column=1, sticky="w", pady=4)
        ttk.Label(options, text="Tool name").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=4)
        ttk.Entry(options, textvariable=self.tool_name).grid(row=1, column=1, sticky="ew", pady=4)
        ttk.Label(options, text="Leave blank to use the findings filename.").grid(row=1, column=2, sticky="w", padx=(8, 0))
        ttk.Checkbutton(
            options, text="Update selected register in place", variable=self.update_in_place,
            command=self.toggle_output_state,
        ).grid(row=2, column=0, columnspan=2, sticky="w", pady=(7, 2))
        ttk.Checkbutton(
            options, text="Create a timestamped backup before replacing it", variable=self.create_backup,
        ).grid(row=3, column=0, columnspan=2, sticky="w", pady=2)
        ttk.Label(options, text="Output register").grid(row=4, column=0, sticky="w", padx=(0, 8), pady=4)
        self.output_entry = ttk.Entry(options, textvariable=self.path_vars["output"])
        self.output_entry.grid(row=4, column=1, sticky="ew", pady=4)
        self.output_button = ttk.Button(options, text="Browse…", command=self.browse_output)
        self.output_button.grid(row=4, column=2, padx=(8, 0), pady=4)

        advanced = ttk.LabelFrame(main, text="3. Advanced findings-column mapping", padding=10)
        advanced.grid(row=2, column=0, sticky="ew", pady=(10, 0))
        ttk.Label(
            advanced,
            text="Use auto unless the findings headers are unusual. Enter the exact header, or none to disable a mapping.",
        ).grid(row=0, column=0, columnspan=8, sticky="w", pady=(0, 8))
        labels = (
            ("Segment", "segment"), ("Filename", "filename"), ("Finding/message", "message"),
            ("Severity", "severity"), ("Finding ID", "finding_id"), ("ORI", "ori"),
            ("Correct TRA", "correct_tra"), ("Modified TRA", "modified_tra"),
        )
        for index, (label, key) in enumerate(labels):
            row, pair = divmod(index, 4)
            column = pair * 2
            ttk.Label(advanced, text=label).grid(row=row + 1, column=column, sticky="w", padx=(0, 5), pady=3)
            ttk.Entry(advanced, textvariable=self.mapping_vars[key], width=17).grid(row=row + 1, column=column + 1, sticky="ew", padx=(0, 12), pady=3)
            advanced.columnconfigure(column + 1, weight=1)
        ttk.Label(advanced, text="Segment number offset").grid(row=3, column=0, sticky="w", pady=(7, 3))
        ttk.Entry(advanced, textvariable=self.segment_offset, width=8).grid(row=3, column=1, sticky="w", pady=(7, 3))
        ttk.Label(advanced, text="Example: -1 if a tool reports CSV line numbers including the header.").grid(row=3, column=2, columnspan=6, sticky="w", pady=(7, 3))

        activity = ttk.LabelFrame(main, text="4. Evaluation", padding=10)
        activity.grid(row=3, column=0, sticky="nsew", pady=(10, 0))
        activity.columnconfigure(0, weight=1)
        activity.rowconfigure(1, weight=1)
        controls = ttk.Frame(activity)
        controls.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        self.run_button = ttk.Button(controls, text="Evaluate and update register", command=self.start)
        self.run_button.pack(side="left")
        ttk.Label(controls, textvariable=self.status).pack(side="left", padx=12)
        self.log = ScrolledText(activity, height=12, wrap="word", state="disabled")
        self.log.grid(row=1, column=0, sticky="nsew")

    def browse_file(self, key: str, filetypes: Sequence[tuple[str, str]]) -> None:
        selected = filedialog.askopenfilename(title=f"Select {key}", filetypes=filetypes)
        if selected:
            self.path_vars[key].set(selected)

    def browse_output(self) -> None:
        selected = filedialog.asksaveasfilename(
            title="Save updated register as", defaultextension=".xlsx",
            filetypes=(("Excel workbooks", "*.xlsx"),),
        )
        if selected:
            self.path_vars["output"].set(selected)

    def toggle_output_state(self) -> None:
        state = "disabled" if self.update_in_place.get() else "normal"
        self.output_entry.configure(state=state)
        self.output_button.configure(state=state)

    def log_line(self, message: str) -> None:
        self.log.configure(state="normal")
        self.log.insert("end", message.rstrip() + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def collect_options(self, allow_duplicate: bool = False) -> EvaluationOptions:
        try:
            offset = int(self.segment_offset.get().strip() or "0")
        except ValueError as exc:
            raise EvaluationError("Segment number offset must be a whole number.") from exc
        mapping = ColumnMapping(**{
            key: (value.get().strip() or "auto")
            for key, value in self.mapping_vars.items()
        })
        output_text = self.path_vars["output"].get().strip()
        return EvaluationOptions(
            golden=Path(self.path_vars["golden"].get().strip()),
            modified=Path(self.path_vars["modified"].get().strip()),
            register=Path(self.path_vars["register"].get().strip()),
            findings=Path(self.path_vars["findings"].get().strip()),
            output_register=Path(output_text) if output_text else None,
            update_in_place=self.update_in_place.get(),
            create_backup=self.create_backup.get(),
            allow_duplicate=allow_duplicate,
            tester=self.tester.get(),
            tool_name=self.tool_name.get(),
            segment_offset=offset,
            mappings=mapping,
        )

    def start(self) -> None:
        if self.worker and self.worker.is_alive():
            return
        try:
            options = self.collect_options()
            validate_options(options)
            if not options.update_in_place and options.output_register is None:
                raise EvaluationError("Choose an output register or enable in-place updating.")
        except EvaluationError as exc:
            messagebox.showerror(APP_NAME, str(exc), parent=self.root)
            return
        self.run_button.configure(state="disabled")
        self.status.set("Evaluating…")
        self.log_line(f"Reading findings: {options.findings}")
        self.worker = threading.Thread(target=self.run_worker, args=(options,), daemon=True)
        self.worker.start()

    def run_worker(self, options: EvaluationOptions) -> None:
        try:
            result = evaluate(options)
            self.events.put(("done", result))
        except DuplicateRunError as exc:
            self.events.put(("duplicate", (options, str(exc))))
        except Exception as exc:
            self.events.put(("error", (str(exc), traceback.format_exc())))

    def poll_events(self) -> None:
        try:
            while True:
                kind, payload = self.events.get_nowait()
                if kind == "done":
                    result: EvaluationResult = payload
                    self.run_button.configure(state="normal")
                    self.status.set("Completed")
                    self.log_line(
                        f"Completed {result.run_id}: {result.true_positives} true positives, "
                        f"{result.false_positives} false positives, {result.false_negatives} false negatives."
                    )
                    self.log_line(f"Saved register: {result.output_register}")
                    if result.backup:
                        self.log_line(f"Backup: {result.backup}")
                    for warning in result.warnings:
                        self.log_line(f"Warning: {warning}")
                    messagebox.showinfo(
                        APP_NAME,
                        f"Evaluation saved.\n\nTrue positives: {result.true_positives}\n"
                        f"False positives: {result.false_positives}\nFalse negatives: {result.false_negatives}\n\n"
                        f"{result.output_register}",
                        parent=self.root,
                    )
                elif kind == "duplicate":
                    options, message = payload
                    self.run_button.configure(state="normal")
                    self.status.set("Duplicate findings detected")
                    if messagebox.askyesno(
                        APP_NAME,
                        message + "\n\nAppend it again anyway?",
                        parent=self.root,
                    ):
                        options.allow_duplicate = True
                        self.run_button.configure(state="disabled")
                        self.status.set("Evaluating repeated run…")
                        self.worker = threading.Thread(target=self.run_worker, args=(options,), daemon=True)
                        self.worker.start()
                elif kind == "error":
                    message, details = payload
                    self.run_button.configure(state="normal")
                    self.status.set("Evaluation failed")
                    self.log_line(f"Error: {message}")
                    self.log_line(details)
                    messagebox.showerror(APP_NAME, message, parent=self.root)
        except queue.Empty:
            pass
        self.root.after(100, self.poll_events)

    def show_about(self) -> None:
        messagebox.showinfo(
            f"About {APP_NAME}",
            f"{APP_NAME} {APP_VERSION}\n\n"
            "Inputs: golden CSV/TSV, modified CSV/TSV, dataset-register XLSX, and findings CSV/TSV/JSON/XML.\n\n"
            "A finding is a true positive when it matches a seeded error by filename and segment, "
            "or by segment plus exact modified text. Seeded rows without segment numbers are matched "
            "by exact modified text plus filename or source text. Text comparison normalises Unicode, "
            "quotes, case, non-breaking spaces, and repeated whitespace.\n\n"
            "The findings and evaluation worksheets are appended. Reprocessing identical findings is "
            "blocked unless explicitly confirmed. Source data files are never changed.",
            parent=self.root,
        )

    def on_close(self) -> None:
        if self.worker and self.worker.is_alive():
            if not messagebox.askyesno(APP_NAME, "An evaluation is running. Close the application anyway?", parent=self.root):
                return
        save_settings({
            "schema_version": SETTINGS_SCHEMA_VERSION,
            "geometry": self.root.geometry(),
            "update_in_place": self.update_in_place.get(),
            "create_backup": self.create_backup.get(),
            "tester": self.tester.get(),
        })
        self.root.destroy()


def parse_cli(argv: Sequence[str] | None = None) -> argparse.Namespace:
    raw_arguments = list(sys.argv[1:] if argv is None else argv)
    script_path = Path(__file__).resolve()
    arguments: list[str] = []
    for value in raw_arguments:
        if not value:
            continue
        if not value.startswith("-"):
            try:
                if Path(value).resolve() == script_path:
                    continue
            except OSError:
                pass
        arguments.append(value)

    parser = argparse.ArgumentParser(
        description=f"{APP_NAME} {APP_VERSION}: append findings and run metrics to an XLSX dataset register."
    )
    parser.add_argument("--headless", action="store_true", help="Run without the desktop UI")
    parser.add_argument("--golden", help="Golden-standard CSV or TSV")
    parser.add_argument("--modified", help="Modified CSV or TSV")
    parser.add_argument("--register", help="Dataset-register XLSX")
    parser.add_argument("--findings", help="Tool findings in CSV, TSV, JSON, or XML")
    parser.add_argument("--output-register", help="Write an updated copy here instead of replacing the selected register")
    parser.add_argument("--no-backup", action="store_true", help="Do not create a backup when updating in place")
    parser.add_argument("--allow-duplicate", action="store_true", help="Allow the same findings content to be appended again")
    parser.add_argument("--dry-run", action="store_true", help="Evaluate and print metrics without writing a workbook")
    parser.add_argument("--tester", default=default_username(), help="Tester name stored with each finding")
    parser.add_argument("--tool-name", default="", help="Name of the tested checker; defaults to the findings filename")
    parser.add_argument(
        "--segment-offset", type=int, default=0,
        help="Add this integer to every numeric finding segment (for example -1 for header-inclusive line numbers)",
    )
    for option, label in (
        ("segment-column", "segment"),
        ("filename-column", "filename"),
        ("message-column", "finding/message"),
        ("severity-column", "severity"),
        ("id-column", "finding ID"),
        ("ori-column", "ORI/source text"),
        ("correct-tra-column", "correct TRA"),
        ("modified-tra-column", "modified TRA/target text"),
    ):
        parser.add_argument(
            f"--{option}", default="auto", metavar="HEADER",
            help=f"Exact {label} header, auto (default), or none",
        )
    return parser.parse_args(arguments)


def options_from_args(args: argparse.Namespace) -> EvaluationOptions:
    missing = [name for name in ("golden", "modified", "register", "findings") if not getattr(args, name)]
    if missing:
        raise EvaluationError("Headless mode requires: " + ", ".join(f"--{name}" for name in missing))
    mapping = ColumnMapping(
        segment=args.segment_column,
        filename=args.filename_column,
        message=args.message_column,
        severity=args.severity_column,
        finding_id=args.id_column,
        ori=args.ori_column,
        correct_tra=args.correct_tra_column,
        modified_tra=args.modified_tra_column,
    )
    return EvaluationOptions(
        golden=Path(args.golden),
        modified=Path(args.modified),
        register=Path(args.register),
        findings=Path(args.findings),
        output_register=Path(args.output_register) if args.output_register else None,
        update_in_place=not bool(args.output_register),
        create_backup=not args.no_backup,
        allow_duplicate=args.allow_duplicate,
        tester=args.tester,
        tool_name=args.tool_name,
        segment_offset=args.segment_offset,
        mappings=mapping,
        dry_run=args.dry_run,
    )


def report_startup_failure(message: str, details: str = "") -> None:
    if tk is not None:
        try:
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror(APP_NAME, message + (f"\n\n{details}" if details else ""), parent=root)
            root.destroy()
            return
        except Exception:
            pass
    print(message, file=sys.stderr)
    if details:
        print(details, file=sys.stderr)


def headless_main(args: argparse.Namespace) -> int:
    try:
        result = evaluate(options_from_args(args))
    except EvaluationError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2
    print(json.dumps(result.summary(), ensure_ascii=False, indent=2))
    return 0


def main(argv: Sequence[str] | None = None) -> int:
    args = parse_cli(argv)
    if args.headless:
        return headless_main(args)
    if tk is None:
        report_startup_failure("This Python installation does not include Tkinter, so the desktop interface cannot open.")
        return 2
    try:
        os.chdir(Path(__file__).resolve().parent)
    except OSError:
        pass
    try:
        root = tk.Tk()
    except Exception as exc:
        report_startup_failure(
            "The desktop interface could not be created.",
            "".join(traceback.format_exception(type(exc), exc, exc.__traceback__)),
        )
        return 2
    try:
        style = ttk.Style(root)
        if sys.platform.startswith("win") and "vista" in style.theme_names():
            style.theme_use("vista")
    except Exception:
        pass
    app = EvaluatorApp(root)
    for key in ("golden", "modified", "register", "findings"):
        value = getattr(args, key, None)
        if value:
            app.path_vars[key].set(str(Path(value).resolve()))
    if args.output_register:
        app.update_in_place.set(False)
        app.path_vars["output"].set(str(Path(args.output_register).resolve()))
        app.toggle_output_state()
    root.mainloop()
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except SystemExit:
        raise
    except Exception as exc:
        report_startup_failure(
            "The evaluator encountered an unexpected error while starting.",
            "".join(traceback.format_exception(type(exc), exc, exc.__traceback__)),
        )
        raise SystemExit(2)
