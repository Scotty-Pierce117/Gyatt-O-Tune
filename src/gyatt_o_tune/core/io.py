from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
import re
import struct
from typing import Any
import xml.etree.ElementTree as ET

# ---------------------------------------------------------------------------
# Explicit axis lookup for well-known MS3/MegaSquirt table names.
# Keys are exact table names from the MSQ; values are (x_candidates, y_candidates).
# ---------------------------------------------------------------------------
_EXPLICIT_AXIS_MAP: dict[str, tuple[list[str], list[str]]] = {
    # Rotary split table
    "RotarySplitTable":              (["RotarySplitRPM"],                    ["RotarySplitLoad"]),
    # Alpha-N MAP table
    "alphaMAPtable":                 (["amap_rpm"],                           ["amap_tps"]),
    # Anti-lag system
    "als_addfuel":                   (["als_rpms"],                           ["als_tpss"]),
    "als_fuelcut":                   (["als_rpms"],                           ["als_tpss"]),
    "als_rifuelcut":                 (["als_rirpms"],                         ["als_ritpss"]),
    "als_sparkcut":                  (["als_rpms"],                           ["als_tpss"]),
    "als_timing":                    (["als_rpms"],                           ["als_tpss"]),
    # Closed-loop boost PWM targets (numbered 1-6, Xboost variant 1-2)
    "boost_ctl_cl_pwm_targs1":       (["Xboost_ctl_cl_pwm_rpms1",       "boost_ctl_cl_pwm_rpms"],        ["Xboost_ctl_cl_pwm_targboosts1", "boost_ctl_cl_pwm_targboosts"]),
    "boost_ctl_cl_pwm_targs2":       (["Xboost_ctl_cl_pwm_rpms2",       "boost_ctl_cl_pwm_rpms"],        ["Xboost_ctl_cl_pwm_targboosts2", "boost_ctl_cl_pwm_targboosts"]),
    "boost_ctl_cl_pwm_targs3":       (["boost_ctl_cl_pwm_rpms"],                                         ["boost_ctl_cl_pwm_targboosts"]),
    "boost_ctl_cl_pwm_targs4":       (["boost_ctl_cl_pwm_rpms"],                                         ["boost_ctl_cl_pwm_targboosts"]),
    "boost_ctl_cl_pwm_targs5":       (["boost_ctl_cl_pwm_rpms"],                                         ["boost_ctl_cl_pwm_targboosts"]),
    "boost_ctl_cl_pwm_targs6":       (["boost_ctl_cl_pwm_rpms"],                                         ["boost_ctl_cl_pwm_targboosts"]),
    "Xboost_ctl_cl_pwm_targs1":      (["Xboost_ctl_cl_pwm_rpms1",       "boost_ctl_cl_pwm_rpms"],        ["Xboost_ctl_cl_pwm_targboosts1", "boost_ctl_cl_pwm_targboosts"]),
    "Xboost_ctl_cl_pwm_targs2":      (["Xboost_ctl_cl_pwm_rpms2",       "boost_ctl_cl_pwm_rpms"],        ["Xboost_ctl_cl_pwm_targboosts2", "boost_ctl_cl_pwm_targboosts"]),
    # Boost load targets (open-loop and closed-loop)
    "boost_ctl_load_targets":        (["boost_ctl_loadtarg_rpm_bins"],                                   ["boost_ctl_loadtarg_tps_bins"]),
    "boost_ctl_load_targets2":       (["boost_ctl_loadtarg_rpm_bins2"],                                  ["boost_ctl_loadtarg_tps_bins2"]),
    "boost_ctl_pwm_targets":         (["boost_ctl_pwmtarg_rpm_bins"],                                    ["boost_ctl_pwmtarg_tps_bins"]),
    "boost_ctl_pwm_targets2":        (["boost_ctl_pwmtarg_rpm_bins2"],                                   ["boost_ctl_pwmtarg_tps_bins2"]),
    # Boost dome pressure targets
    "boost_dome_targets1":           (["boost_dome_target_rpms1"],                                       ["boost_dome_target_kpas1"]),
    "boost_dome_targets2":           (["boost_dome_target_rpms2"],                                       ["boost_dome_target_kpas2"]),
    # Dwell
    "dwell_table_values":            (["dwell_table_rpms"],                                              ["dwell_table_loads"]),
    # EGO authority and delay
    "ego_auth_table":                (["ego_auth_rpms"],                                                  ["ego_auth_loads"]),
    "ego_auth_table2":               (["ego_auth_rpms"],                                                  ["ego_auth_loads"]),
    "ego_delay_table":               (["ego_delay_rpms"],                                                 ["ego_delay_loads"]),
    # Electronic throttle control
    "etc_targ_pos":                  (["etc_rpms"],                                                       ["etc_pedal_pos"]),
    # Fuel pump duty
    "fpd_duty":                      (["fpd_rpm"],                                                        ["fpd_load"]),
    # Generic PID
    "generic_pid_targets_a":         (["generic_pid_rpms_a"],                                             ["generic_pid_loadvals_a"]),
    "generic_pid_targets_b":         (["generic_pid_rpms_b"],                                             ["generic_pid_loadvals_b"]),
    # Idle VE tables
    "idleve_table1":                 (["idleve_rpms1"],                                                   ["idleve_loads1"]),
    "idleve_table2":                 (["idleve_rpms2"],                                                   ["idleve_loads2"]),
    # Injector deadtime (X=volts, Y=fuel pressure differential)
    "inj_deadtime_table1":           (["inj_deadtime_volts1"],                                            ["inj_deadtime_pressure1"]),
    "inj_deadtime_table2":           (["inj_deadtime_volts2"],                                            ["inj_deadtime_pressure2"]),
    "inj_deadtime_table3":           (["inj_deadtime_volts3"],                                            ["inj_deadtime_pressure3"]),
    "inj_deadtime_table4":           (["inj_deadtime_volts4"],                                            ["inj_deadtime_pressure4"]),
    # Injection timing
    "inj_timing":                    (["inj_timing_rpm"],                                                  ["inj_timing_load"]),
    "inj_timing_sec":                (["inj_timing_sec_rpm"],                                              ["inj_timing_sec_load"]),
    # Long-term trim table
    "ltt_table1":                    (["ltt_rpms"],                                                        ["ltt_loads"]),
    # MAP prediction
    "map_predict_lookup_table":      (["map_predict_rpm"],                                                 ["map_predict_tps"]),
    "map_predict_lookup_table2":     (["map_predict_rpm2"],                                                ["map_predict_tps2"]),
    # Max AFR differential
    "maxafr1_diff":                  (["maxafr1_rpm"],                                                     ["maxafr1_load"]),
    # Narrowband targets
    "narrowband_tgts":               (["narrowband_tgts_rpms"],                                           ["narrowband_tgts_loads"]),
    # Per-channel PWM output duty tables (a-f)
    "pwm_duties_a":                  (["pwm_rpms_a"],                                                     ["pwm_loadvals_a"]),
    "pwm_duties_b":                  (["pwm_rpms_b"],                                                     ["pwm_loadvals_b"]),
    "pwm_duties_c":                  (["pwm_rpms_c"],                                                     ["pwm_loadvals_c"]),
    "pwm_duties_d":                  (["pwm_rpms_d"],                                                     ["pwm_loadvals_d"]),
    "pwm_duties_e":                  (["pwm_rpms_e"],                                                     ["pwm_loadvals_e"]),
    "pwm_duties_f":                  (["pwm_rpms_f"],                                                     ["pwm_loadvals_f"]),
    # PWM idle CL initial value tables (X=RPM, Y=coolant/MAT temperature)
    "pwmidle_cl_initialvalues_duties": (["pwmidle_cl_initialvalue_rpms"],                                 ["pwmidle_cl_initialvalue_matorclt"]),
    "pwmidle_cl_initialvalues_steps":  (["pwmidle_cl_initialvalue_rpms"],                                 ["pwmidle_cl_initialvalue_matorclt"]),
    # Staged injection
    "staged_percents":               (["staged_rpms"],                                                    ["staged_loads"]),
    # VSS differential (two wheel-speed inputs)
    "vss_diff_table":                (["vss_diff_vss1"],                                                  ["vss_diff_vss2"]),
}

# Cylinder-trim tables use a single undelimited letter suffix (a-p).
# Build these programmatically to avoid repetition.
for _ch in "abcdefghijklmnop":
    _EXPLICIT_AXIS_MAP[f"inj_trim{_ch}"] = (["inj_trim_rpm"], ["inj_trim_load"])
    _EXPLICIT_AXIS_MAP[f"spk_trim{_ch}"] = (["spk_trim_rpm"], ["spk_trim_load"])


@dataclass
class TuneData:
    file_path: Path
    raw_text: str
    text_encoding: str
    tables: dict[str, "TableData"] = field(default_factory=dict)
    vectors: dict[str, "AxisVector"] = field(default_factory=dict)

    def resolve_table_axes(self, table: "TableData") -> tuple["AxisVector | None", "AxisVector | None"]:
        table_name_lower = table.name.lower()

        # --- 1. VVT timing / on-off tables ---
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
                exclude_name=x_axis.name if x_axis else None,
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
                exclude_name=x_axis.name if x_axis else None,
            )
            return x_axis, y_axis

        # --- 2. Knock threshold (single axis) ---
        if table.name == "knock_thresholds":
            knock_rpm_axis = self.vectors.get("knock_rpms")
            if knock_rpm_axis is not None:
                if knock_rpm_axis.length == table.cols:
                    return knock_rpm_axis, None
                if knock_rpm_axis.length == table.rows:
                    return None, knock_rpm_axis

        # --- 3. Explicit per-table-name lookup ---
        if table.name in _EXPLICIT_AXIS_MAP:
            x_cands, y_cands = _EXPLICIT_AXIS_MAP[table.name]
            x_axis = self._pick_vector(x_cands, target_length=table.cols, preferred_unit="RPM")
            y_axis = self._pick_vector(
                y_cands,
                target_length=table.rows,
                preferred_unit=None,
                exclude_name=x_axis.name if x_axis else None,
            )
            return x_axis, y_axis

        # --- 4. Pattern-based: derive candidate names from the table name ---
        x_cands, y_cands = self._derive_axis_candidates(table.name)
        x_axis = self._pick_vector(x_cands, target_length=table.cols, preferred_unit="RPM", names_only=True)
        y_axis = self._pick_vector(
            y_cands,
            target_length=table.rows,
            preferred_unit=None,
            names_only=True,
            exclude_name=x_axis.name if x_axis else None,
        )
        if x_axis is not None or y_axis is not None:
            # Fill any missing side using the safe fallback (excludes UNALLOCATED/RAW).
            if x_axis is None:
                x_axis = self._pick_vector(
                    [], table.cols, "RPM",
                    exclude_name=y_axis.name if y_axis else None,
                )
            if y_axis is None:
                y_axis = self._pick_vector(
                    [], table.rows, None,
                    exclude_name=x_axis.name if x_axis else None,
                )
            return x_axis, y_axis

        # --- 5. Number-suffix approach (veTable1, advanceTable1, afrTable1, etc.) ---
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
        y_axis = self._pick_vector(
            candidate_y_names,
            target_length=table.rows,
            preferred_unit=None,
            exclude_name=x_axis.name if x_axis else None,
        )
        return x_axis, y_axis

    def _pick_vector(
        self,
        candidate_names: list[str],
        target_length: int,
        preferred_unit: str | None,
        exclude_name: str | None = None,
        names_only: bool = False,
    ) -> "AxisVector | None":
        """Return the best matching AxisVector for the given candidate names and length.

        When *names_only* is True the fallback scan is skipped entirely so the
        caller can distinguish "no named match" from "fallback match".
        The fallback always excludes UNALLOCATED_SPACE / RAW vectors and any
        vector identified by *exclude_name* (used to prevent X and Y resolving
        to the same vector).
        """
        for name in candidate_names:
            vector = self.vectors.get(name)
            if vector and vector.length == target_length:
                if exclude_name is None or vector.name != exclude_name:
                    return vector

        if names_only:
            return None

        # Fallback: scan all vectors by unit preference, skipping bad candidates.
        fallback: AxisVector | None = None
        for vector in self.vectors.values():
            if vector.length != target_length:
                continue
            if exclude_name and vector.name == exclude_name:
                continue
            if self._is_unallocated_vector(vector):
                continue
            if preferred_unit and vector.units and vector.units.upper() == preferred_unit.upper():
                return vector
            if fallback is None:
                fallback = vector
        return fallback

    @staticmethod
    def _is_unallocated_vector(vector: "AxisVector") -> bool:
        """Return True for placeholder / unallocated vectors that should never be used as axes."""
        name_upper = vector.name.upper()
        if "UNALLOCATED" in name_upper:
            return True
        if vector.units and vector.units.upper() == "RAW":
            return True
        return False

    @staticmethod
    def _derive_axis_candidates(table_name: str) -> tuple[list[str], list[str]]:
        """Generate RPM-axis and load/map-axis candidate names from table-name patterns.

        This covers tables whose axis vectors share a common name stem with the
        table itself (e.g. ``dwell_table_values`` → ``dwell_table_rpms`` /
        ``dwell_table_loads``), and cylinder-trim tables that use an undelimited
        single-letter suffix (``inj_trima`` → ``inj_trim_rpm`` / ``inj_trim_load``).
        """
        x_cands: list[str] = []
        y_cands: list[str] = []

        def add_rpm_variants(stem: str) -> None:
            for sfx in ("_rpm", "_rpms", "_rpmbin", "_rpmbins"):
                x_cands.append(stem + sfx)

        def add_load_variants(stem: str) -> None:
            for sfx in ("_load", "_loads", "_kpa", "_kpas", "_tps", "_tpss",
                        "_map", "_pressure", "_volts", "_loadvals"):
                y_cands.append(stem + sfx)

        # Direct: {table_name}_rpm / {table_name}_load
        add_rpm_variants(table_name)
        add_load_variants(table_name)

        # Strip common MS3 table-name suffixes and try stem variants.
        stripped = re.sub(
            r"(_lookup_table\d*|_table_values?\d*|_table\d*|_values?\d*"
            r"|_targs?\d*|_targets?\d*|_percents?\d*|_diff\d*"
            r"|_duties?\d*|_duty\d*)$",
            "",
            table_name,
            flags=re.IGNORECASE,
        )
        if stripped and stripped != table_name:
            add_rpm_variants(stripped)
            add_load_variants(stripped)

        # Cylinder-trim pattern: {stem}_trim{letter a-p} (no underscore before letter).
        m_trim = re.match(r"^(.+_trim)([a-p])$", table_name)
        if m_trim:
            stem = m_trim.group(1)
            add_rpm_variants(stem)
            add_load_variants(stem)

        # Underscore-delimited single-letter suffix: {stem}_{letter a-p}
        m_letter = re.match(r"^(.+)_([a-p])$", table_name)
        if m_letter:
            stem = m_letter.group(1)
            letter = m_letter.group(2)
            add_rpm_variants(stem)
            add_load_variants(stem)
            # Also try per-channel variants: {stem}_rpm_{letter}
            for sfx in ("_rpm", "_rpms"):
                x_cands.append(f"{stem}{sfx}_{letter}")
            for sfx in ("_load", "_loads", "_loadvals"):
                y_cands.append(f"{stem}{sfx}_{letter}")

        # Numbered suffix: try both with and without the trailing number.
        m_num = re.match(r"^(.+?)(\d+)$", table_name)
        if m_num:
            stem, num = m_num.group(1), m_num.group(2)
            add_rpm_variants(stem + num)
            add_load_variants(stem + num)
            add_rpm_variants(stem)
            add_load_variants(stem)

        return x_cands, y_cands

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
