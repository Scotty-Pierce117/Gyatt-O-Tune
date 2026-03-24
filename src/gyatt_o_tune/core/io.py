from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
import re
from typing import Any
import xml.etree.ElementTree as ET

@dataclass
class TuneData:
    file_path: Path
    raw_text: str
    tables: dict[str, "TableData"] = field(default_factory=dict)
    vectors: dict[str, "AxisVector"] = field(default_factory=dict)

    def resolve_table_axes(self, table: "TableData") -> tuple["AxisVector | None", "AxisVector | None"]:
        if table.name == "knock_thresholds":
            knock_rpm_axis = self.vectors.get("knock_rpms")
            if knock_rpm_axis is not None:
                if knock_rpm_axis.length == table.cols:
                    return knock_rpm_axis, None
                if knock_rpm_axis.length == table.rows:
                    return None, knock_rpm_axis

        table_number = self._extract_table_number(table.name)
        preferred_prefix = self._preferred_axis_prefix(table.name)

        candidate_x_names: list[str] = []
        candidate_y_names: list[str] = []
        if table_number:
            for prefix in [preferred_prefix, "f", "s", "a", "k", "e", "v"]:
                if not prefix:
                    continue
                candidate_x_names.append(f"{prefix}rpm_table{table_number}")
                candidate_y_names.append(f"{prefix}map_table{table_number}")

        candidate_x_names.extend([f"X{table.name}", f"{table.name}X", f"x_{table.name}", f"{table.name}_x"])
        candidate_y_names.extend([f"Y{table.name}", f"{table.name}Y", f"y_{table.name}", f"{table.name}_y"])

        x_axis = self._pick_vector(candidate_x_names, target_length=table.cols, preferred_unit="RPM")
        y_axis = self._pick_vector(candidate_y_names, target_length=table.rows, preferred_unit=None)
        return x_axis, y_axis

    def _pick_vector(
        self,
        candidate_names: list[str],
        target_length: int,
        preferred_unit: str | None,
    ) -> "AxisVector | None":
        for name in candidate_names:
            vector = self.vectors.get(name)
            if vector and vector.length == target_length:
                return vector

        fallback: AxisVector | None = None
        for vector in self.vectors.values():
            if vector.length != target_length:
                continue
            if preferred_unit and vector.units and vector.units.upper() == preferred_unit.upper():
                return vector
            if fallback is None:
                fallback = vector
        return fallback

    @staticmethod
    def _extract_table_number(table_name: str) -> str | None:
        match = re.search(r"(\d+)$", table_name)
        return match.group(1) if match else None

    @staticmethod
    def _preferred_axis_prefix(table_name: str) -> str | None:
        lowered = table_name.lower()
        if lowered.startswith("ve"):
            return "f"
        if lowered.startswith("advance") or lowered.startswith("spk"):
            return "s"
        if lowered.startswith("afr"):
            return "a"
        if lowered.startswith("knock") or lowered.startswith("knk"):
            return "k"
        return None


@dataclass
class TableData:
    name: str
    source_tag: str
    rows: int
    cols: int
    units: str | None
    digits: str | None
    values: list[list[float]]

    @property
    def min_value(self) -> float:
        return min(min(row) for row in self.values)

    @property
    def max_value(self) -> float:
        return max(max(row) for row in self.values)


@dataclass
class AxisVector:
    name: str
    source_tag: str
    length: int
    orientation: str
    units: str | None
    digits: str | None
    values: list[float]


@dataclass
class LogParseResult:
    dataframe: Any
    parser_used: str
    encoding: str
    notes: str = ""


class TuneLoader:
    """Extract matrix tables from MegaSquirt MSQ XML."""

    def load(self, file_path: Path) -> TuneData:
        raw_text = file_path.read_text(encoding="utf-8", errors="ignore")
        tables, vectors = self._parse_tables(raw_text)
        return TuneData(file_path=file_path, raw_text=raw_text, tables=tables, vectors=vectors)

    def _parse_tables(self, raw_text: str) -> tuple[dict[str, TableData], dict[str, AxisVector]]:
        root = ET.fromstring(raw_text)
        namespace = self._extract_namespace(root.tag)
        tag_candidates = [f".//{namespace}constant", f".//{namespace}pcVariable"] if namespace else [".//constant", ".//pcVariable"]

        tables: dict[str, TableData] = {}
        vectors: dict[str, AxisVector] = {}
        for selector in tag_candidates:
            for element in root.findall(selector):
                rows_text = element.attrib.get("rows")
                cols_text = element.attrib.get("cols")
                if not rows_text or not cols_text:
                    continue

                try:
                    rows = int(float(rows_text))
                    cols = int(float(cols_text))
                except ValueError:
                    continue

                values = self._extract_numeric_values(element.text or "")
                cell_count = rows * cols
                if len(values) < cell_count:
                    continue

                name = element.attrib.get("name", f"unnamed_{len(tables) + 1}")
                if rows >= 2 and cols >= 2:
                    matrix = [values[idx * cols : (idx + 1) * cols] for idx in range(rows)]
                    tables[name] = TableData(
                        name=name,
                        source_tag=self._strip_namespace(element.tag),
                        rows=rows,
                        cols=cols,
                        units=element.attrib.get("units"),
                        digits=element.attrib.get("digits"),
                        values=matrix,
                    )
                    continue

                if rows == 1 and cols >= 2:
                    vectors[name] = AxisVector(
                        name=name,
                        source_tag=self._strip_namespace(element.tag),
                        length=cols,
                        orientation="row",
                        units=element.attrib.get("units"),
                        digits=element.attrib.get("digits"),
                        values=values[:cols],
                    )
                elif cols == 1 and rows >= 2:
                    vectors[name] = AxisVector(
                        name=name,
                        source_tag=self._strip_namespace(element.tag),
                        length=rows,
                        orientation="column",
                        units=element.attrib.get("units"),
                        digits=element.attrib.get("digits"),
                        values=values[:rows],
                    )

        # Promote all 1D vectors to selectable tables for the All Tables view.
        for vector_name, vector in vectors.items():
            if vector_name in tables:
                continue

            if vector.orientation == "row":
                table_values = [[float(v) for v in vector.values]]
                rows = 1
                cols = vector.length
            else:
                table_values = [[float(v)] for v in vector.values]
                rows = vector.length
                cols = 1

            tables[vector_name] = TableData(
                name=vector_name,
                source_tag=vector.source_tag,
                rows=rows,
                cols=cols,
                units=vector.units,
                digits=vector.digits,
                values=table_values,
            )

        return tables, vectors

    @staticmethod
    def _extract_namespace(tag: str) -> str:
        if tag.startswith("{") and "}" in tag:
            return tag[: tag.index("}") + 1]
        return ""

    @staticmethod
    def _strip_namespace(tag: str) -> str:
        if "}" in tag:
            return tag.split("}", maxsplit=1)[1]
        return tag

    @staticmethod
    def _extract_numeric_values(text: str) -> list[float]:
        number_tokens = re.findall(r"[-+]?\d*\.?\d+(?:[eE][-+]?\d+)?", text)
        return [float(token) for token in number_tokens]


class LogLoader:
    """Load log files into a DataFrame for plotting and table synthesis."""

    def load_log(self, file_path: Path) -> Any:
        return self.load_log_with_report(file_path).dataframe

    def load_log_with_report(self, file_path: Path) -> LogParseResult:
        import pandas as pd

        encodings = ["utf-8", "cp1252", "latin-1"]

        # First: attempt explicit MegaSquirt tab-log parsing with metadata/header detection.
        for encoding in encodings:
            try:
                preview = self._read_preview_lines(file_path, encoding=encoding, max_lines=300)
                header_idx = self._detect_megasquirt_header_row(preview)
                if header_idx is None:
                    continue

                dataframe = pd.read_csv(
                    file_path,
                    sep="\t",
                    encoding=encoding,
                    skiprows=header_idx,
                    header=0,
                    engine="python",
                    on_bad_lines="skip",
                )
                dataframe = self._drop_units_row_if_present(dataframe)
                dataframe = dataframe.dropna(axis=1, how="all")
                return LogParseResult(
                    dataframe=dataframe,
                    parser_used="megasquirt-tab",
                    encoding=encoding,
                    notes=f"Detected header row at line {header_idx + 1}.",
                )
            except Exception:
                continue

        # Second: generic CSV/TSV fallback parsing.
        for encoding in encodings:
            try:
                dataframe = pd.read_csv(
                    file_path,
                    sep=None,
                    engine="python",
                    encoding=encoding,
                    on_bad_lines="skip",
                )
                dataframe = dataframe.dropna(axis=1, how="all")
                return LogParseResult(dataframe=dataframe, parser_used="generic-auto-sep", encoding=encoding)
            except Exception:
                pass

            try:
                dataframe = pd.read_csv(
                    file_path,
                    encoding=encoding,
                    on_bad_lines="skip",
                )
                dataframe = dataframe.dropna(axis=1, how="all")
                return LogParseResult(dataframe=dataframe, parser_used="generic-csv", encoding=encoding)
            except Exception:
                pass

        dataframe = pd.read_csv(
            file_path,
            sep=None,
            engine="python",
            encoding="utf-8",
            encoding_errors="replace",
            on_bad_lines="skip",
        )
        dataframe = dataframe.dropna(axis=1, how="all")
        return LogParseResult(
            dataframe=dataframe,
            parser_used="generic-replacement",
            encoding="utf-8(replace)",
            notes="Used replacement characters to recover undecodable bytes.",
        )

    @staticmethod
    def _read_preview_lines(file_path: Path, encoding: str, max_lines: int) -> list[str]:
        lines: list[str] = []
        with file_path.open("r", encoding=encoding, errors="strict") as handle:
            for _ in range(max_lines):
                line = handle.readline()
                if not line:
                    break
                lines.append(line.rstrip("\n").rstrip("\r"))
        return lines

    @staticmethod
    def _detect_megasquirt_header_row(lines: list[str]) -> int | None:
        for idx, line in enumerate(lines):
            lowered = line.lower()
            if "\t" not in line:
                continue
            # Typical MegaSquirt header line contains many fields including Time + RPM.
            if "time" in lowered and "rpm" in lowered and line.count("\t") >= 10:
                return idx
        return None

    @staticmethod
    def _drop_units_row_if_present(dataframe: Any) -> Any:
        if dataframe is None or dataframe.empty:
            return dataframe
        first_row = dataframe.iloc[0]
        non_numeric = 0
        total = 0
        for value in first_row:
            text = str(value).strip()
            if text == "" or text.lower() == "nan":
                continue
            total += 1
            try:
                float(text)
            except ValueError:
                non_numeric += 1
        if total > 0 and (non_numeric / total) >= 0.6:
            return dataframe.iloc[1:].reset_index(drop=True)
        return dataframe

    def load_csv(self, file_path: Path) -> Any:
        return self.load_log(file_path)
