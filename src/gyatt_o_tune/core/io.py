from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
import re
import struct
from typing import Any
import xml.etree.ElementTree as ET

@dataclass
class TuneData:
    file_path: Path
    raw_text: str
    text_encoding: str
    tables: dict[str, "TableData"] = field(default_factory=dict)
    vectors: dict[str, "AxisVector"] = field(default_factory=dict)

    def resolve_table_axes(self, table: "TableData") -> tuple["AxisVector | None", "AxisVector | None"]:
        table_name_lower = table.name.lower()

        if table_name_lower.startswith("vvt_timing"):
            x_axis = self._pick_vector(
                ["vvt_timing_rpm", "vvt_onoff_rpms", f"{table.name}_rpm", "vvt_rpm"],
                target_length=table.cols,
                preferred_unit="RPM",
            )
            y_axis = self._pick_vector(
                ["vvt_timing_load", "vvt_onoff_loads", f"{table.name}_load", "vvt_load"],
                target_length=table.rows,
                preferred_unit="kPa",
            )
            return x_axis, y_axis

        if table_name_lower.startswith("vvt_onoff"):
            x_axis = self._pick_vector(
                ["vvt_onoff_rpms", "vvt_timing_rpm", f"{table.name}_rpms", f"{table.name}_rpm"],
                target_length=table.cols,
                preferred_unit="RPM",
            )
            y_axis = self._pick_vector(
                ["vvt_onoff_loads", "vvt_timing_load", f"{table.name}_loads", f"{table.name}_load"],
                target_length=table.rows,
                preferred_unit="kPa",
            )
            return x_axis, y_axis

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
                candidate_y_names.append(f"{prefix}load_table{table_number}")
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
        raw_text, text_encoding = self._read_tune_text(file_path)
        tables, vectors = self._parse_tables(raw_text)
        return TuneData(
            file_path=file_path,
            raw_text=raw_text,
            text_encoding=text_encoding,
            tables=tables,
            vectors=vectors,
        )

    def _read_tune_text(self, file_path: Path) -> tuple[str, str]:
        raw_bytes = file_path.read_bytes()
        candidate_encodings: list[str] = []

        declared_encoding = self._extract_declared_encoding(raw_bytes)
        if declared_encoding:
            candidate_encodings.append(declared_encoding)

        for encoding in ["utf-8", "cp1252", "latin-1"]:
            if encoding.lower() not in {candidate.lower() for candidate in candidate_encodings}:
                candidate_encodings.append(encoding)

        for encoding in candidate_encodings:
            try:
                return raw_bytes.decode(encoding), encoding
            except UnicodeDecodeError:
                continue

        return raw_bytes.decode("latin-1"), "latin-1"

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

    @staticmethod
    def _extract_declared_encoding(raw_bytes: bytes) -> str | None:
        header = raw_bytes[:256].decode("ascii", errors="ignore")
        match = re.search(r'encoding=["\']([^"\']+)["\']', header, re.IGNORECASE)
        return match.group(1) if match else None

    def save(self, tune_data: "TuneData", dest_path: Path) -> None:
        """Write tune data by replacing only the changed table sections in the original XML text."""
        updated_text = tune_data.raw_text

        replacements: list[tuple[int, int, str]] = []
        for table in tune_data.tables.values():
            element_span = self._find_table_text_span(updated_text, table)
            if element_span is None:
                continue

            text_start, text_end = element_span
            original_text = updated_text[text_start:text_end]
            if not self._table_text_differs(original_text, table):
                continue

            replacements.append((text_start, text_end, self._format_table_text(table, original_text)))

        for text_start, text_end, replacement_text in reversed(replacements):
            updated_text = updated_text[:text_start] + replacement_text + updated_text[text_end:]

        with dest_path.open("w", encoding=tune_data.text_encoding, newline="") as handle:
            handle.write(updated_text)
        tune_data.raw_text = updated_text
        tune_data.file_path = dest_path

    @staticmethod
    def _find_table_text_span(raw_text: str, table: "TableData") -> tuple[int, int] | None:
        quoted_name = re.escape(table.name)
        quoted_tag = re.escape(table.source_tag)
        pattern = re.compile(
            rf"<(?:\w+:)?{quoted_tag}\b(?P<attrs>[^>]*)\bname\s*=\s*(['\"])({quoted_name})\2(?P<attrs_after>[^>]*)>(?P<text>.*?)</(?:\w+:)?{quoted_tag}>",
            re.DOTALL,
        )
        match = pattern.search(raw_text)
        if match is None:
            return None
        return match.start("text"), match.end("text")

    def _table_text_differs(self, original_text: str, table: "TableData") -> bool:
        original_values = self._extract_numeric_values(original_text)
        current_values = [float(value) for row in table.values for value in row]
        if len(original_values) < len(current_values):
            return True
        return original_values[: len(current_values)] != current_values

    @staticmethod
    def _format_table_text(table: "TableData", original_text: str) -> str:
        try:
            digits = int(table.digits) if table.digits is not None else 2
        except ValueError:
            digits = 2

        newline = "\r\n" if "\r\n" in original_text else "\n"
        indent_match = re.search(r"(?:^|\r?\n)([ \t]*)\S", original_text)
        indent = indent_match.group(1) if indent_match else "  "

        lines: list[str] = []
        for row in table.values:
            if digits == 0:
                row_str = " ".join(str(int(round(v))) for v in row)
            else:
                row_str = " ".join(f"{v:.{digits}f}" for v in row)
            lines.append(indent + row_str)
        return newline + newline.join(lines) + newline


class LogLoader:
    """Load log files into a DataFrame for plotting and table synthesis."""

    def load_log(self, file_path: Path) -> Any:
        return self.load_log_with_report(file_path).dataframe

    def load_log_with_report(self, file_path: Path) -> LogParseResult:
        import pandas as pd

        encodings = ["utf-8", "cp1252", "latin-1"]
        parse_path = file_path

        if file_path.suffix.lower() == ".mlg" and self._is_probably_binary_log(file_path):
            native_result = self._load_binary_mlvlg_with_report(file_path)
            if native_result is not None:
                return native_result

        # First: attempt explicit MegaSquirt tab-log parsing with metadata/header detection.
        for encoding in encodings:
            try:
                preview = self._read_preview_lines(parse_path, encoding=encoding, max_lines=300)
                header_idx = self._detect_megasquirt_header_row(preview)
                if header_idx is None:
                    continue

                dataframe = pd.read_csv(
                    parse_path,
                    sep="\t",
                    encoding=encoding,
                    skiprows=header_idx,
                    header=0,
                    engine="python",
                    on_bad_lines="skip",
                )
                dataframe = self._drop_units_row_if_present(dataframe)
                dataframe = dataframe.dropna(axis=1, how="all")
                note = f"Detected header row at line {header_idx + 1}."
                return LogParseResult(
                    dataframe=dataframe,
                    parser_used="megasquirt-tab",
                    encoding=encoding,
                    notes=note,
                )
            except Exception:
                continue

        # Second: generic CSV/TSV fallback parsing.
        for encoding in encodings:
            try:
                dataframe = pd.read_csv(
                    parse_path,
                    sep=None,
                    engine="python",
                    encoding=encoding,
                    on_bad_lines="skip",
                )
                dataframe = dataframe.dropna(axis=1, how="all")
                return LogParseResult(
                    dataframe=dataframe,
                    parser_used="generic-auto-sep",
                    encoding=encoding,
                    notes="",
                )
            except Exception:
                pass

            try:
                dataframe = pd.read_csv(
                    parse_path,
                    encoding=encoding,
                    on_bad_lines="skip",
                )
                dataframe = dataframe.dropna(axis=1, how="all")
                return LogParseResult(
                    dataframe=dataframe,
                    parser_used="generic-csv",
                    encoding=encoding,
                    notes="",
                )
            except Exception:
                pass

        dataframe = pd.read_csv(
            parse_path,
            sep=None,
            engine="python",
            encoding="utf-8",
            encoding_errors="replace",
            on_bad_lines="skip",
        )
        dataframe = dataframe.dropna(axis=1, how="all")
        notes = "Used replacement characters to recover undecodable bytes."
        return LogParseResult(
            dataframe=dataframe,
            parser_used="generic-replacement",
            encoding="utf-8(replace)",
            notes=notes,
        )

    @staticmethod
    def _is_probably_binary_log(file_path: Path) -> bool:
        try:
            head = file_path.read_bytes()[:4096]
        except Exception:
            return False
        if head.startswith(b"MLVLG"):
            return True
        if b"\x00" in head:
            return True
        try:
            head.decode("utf-8")
        except UnicodeDecodeError:
            return True
        return False

    @staticmethod
    def _load_binary_mlvlg_with_report(file_path: Path) -> LogParseResult | None:
        import pandas as pd

        try:
            payload = file_path.read_bytes()
            if len(payload) < 24:
                return None

            offset = 0

            def take(count: int) -> bytes:
                nonlocal offset
                if offset + count > len(payload):
                    raise ValueError("Unexpected end of binary .mlg payload")
                chunk = payload[offset : offset + count]
                offset += count
                return chunk

            def u8() -> int:
                return struct.unpack(">B", take(1))[0]

            def i8() -> int:
                return struct.unpack(">b", take(1))[0]

            def u16() -> int:
                return struct.unpack(">H", take(2))[0]

            def i16() -> int:
                return struct.unpack(">h", take(2))[0]

            def u32() -> int:
                return struct.unpack(">I", take(4))[0]

            def i32() -> int:
                return struct.unpack(">i", take(4))[0]

            def i64() -> int:
                return struct.unpack(">q", take(8))[0]

            def f32() -> float:
                return struct.unpack(">f", take(4))[0]

            def clean_string(raw: bytes) -> str:
                return raw.decode("utf-8", errors="ignore").replace("\x00", "").replace('"', "").strip()

            file_format = clean_string(take(6))
            if file_format != "MLVLG":
                return None

            format_version = i16()
            if format_version not in (1, 2):
                return None

            _timestamp = i32()
            info_data_start = i32() if format_version == 2 else i16()
            data_begin_index = i32()
            _record_length = i16()
            num_fields = i16()
            field_length = 89 if format_version == 2 else 55

            fields: list[dict[str, Any]] = []
            for _ in range(num_fields):
                field_type = i8()
                field_name = clean_string(take(34))
                _field_units = clean_string(take(10))
                _display_style = i8()

                field_info: dict[str, Any] = {
                    "type": field_type,
                    "name": field_name,
                }

                if field_type < 10:
                    field_info["scale"] = f32()
                    field_info["transform"] = f32()
                    field_info["digits"] = i8()
                    if format_version == 2:
                        take(34)  # category
                else:
                    take(1)  # bit field style
                    take(4)  # bit field names index
                    take(1)  # bit count
                    take(3)  # reserved
                    if format_version == 2:
                        take(34)  # category

                fields.append(field_info)

            if info_data_start > offset:
                take(info_data_start - offset)

            if data_begin_index > offset:
                take(data_begin_index - offset)

            def read_field_value(field_type: int) -> int | float:
                if field_type in (0, 10):
                    return u8()
                if field_type == 1:
                    return i8()
                if field_type in (2, 11):
                    return u16()
                if field_type == 3:
                    return struct.unpack(">h", take(2))[0]
                if field_type in (4, 12):
                    return u32()
                if field_type == 5:
                    return i32()
                if field_type == 6:
                    return i64()
                if field_type == 7:
                    return f32()
                return u8()

            records: list[dict[str, Any]] = []
            while offset < len(payload):
                if offset + 4 > len(payload):
                    break

                block_type = u8()
                take(1)  # counter
                take(2)  # block timestamp

                if block_type == 0:
                    row: dict[str, Any] = {}
                    for field in fields:
                        raw_value = read_field_value(int(field["type"]))
                        if int(field["type"]) < 10:
                            scaled = (float(raw_value) + float(field.get("transform", 0.0))) * float(field.get("scale", 1.0))
                            row[str(field["name"])] = scaled
                        else:
                            row[str(field["name"])] = raw_value
                    if offset < len(payload):
                        take(1)  # crc
                    records.append(row)
                    continue

                if block_type == 1:
                    if offset + 50 > len(payload):
                        break
                    take(50)  # marker message
                    continue

                break

            if not records:
                return None

            dataframe = pd.DataFrame.from_records(records)
            dataframe = dataframe.dropna(axis=1, how="all")
            return LogParseResult(
                dataframe=dataframe,
                parser_used="mlvlg-binary-native",
                encoding="binary",
                notes="Parsed binary .mlg directly using integrated mlg-converter-compatible logic.",
            )
        except Exception:
            return None

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
