# SPDX-License-Identifier: MIT
#
# Copyright (c) 2025, the xcube team
#               2026, Antonio Valentino
#
# Permission is hereby granted, free of charge, to any person obtaining a
# copy of this software and associated documentation files (the "Software"),
# to deal in the Software without restriction, including without limitation
# the rights to use, copy, modify, merge, publish, distribute, sublicense,
# and/or sell copies of the Software, and to permit persons to whom the
# Software is furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in
# all copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
# FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
# DEALINGS IN THE SOFTWARE.

# /// script
# dependencies = [
#   "openpyxl",
#   "argcomplete",
# ]
# requires-python = ">=3.11"
# ///

"""Generate documentation for the SARCalNet submission template.

The tool analyzes the SARCalNet submission template, extracts
the documentation for each of the fields, and generate
ReStructuredText documentation.

The input template file in XLSX format is assumed to have:
* display title in row 1,
* field name in row 2,
* indication of mandatory/optional fields in row 3,
* literal comment blocks in rows 4-5.

"""

import re
import abc
import types
import pathlib
import argparse
import textwrap
import functools
from typing import NamedTuple
from collections.abc import Mapping, Iterator, Collection

import openpyxl


class FieldInfoRows(NamedTuple):
    title: int = 1
    field_id: int = 2
    mandatory: int = 3
    description: int = 4
    validity_range: int = 5


class FieldInfo(NamedTuple):
    title: str
    field_id: str
    mandatory: str
    description: str
    validity_range: str


MAX_LINE_LEN = 80
DEFAULT_SKIP_SHEETS = frozenset({"README", "definitions"})
DEFAULT_INFO_ROWS = FieldInfoRows()


class DataModel:
    def __init__(
        self,
        filename,
        skip_sheet: Collection[str] = DEFAULT_SKIP_SHEETS,
        field_info_rows: FieldInfoRows = DEFAULT_INFO_ROWS,
    ):
        self._wb = openpyxl.load_workbook(
            filename,
            read_only=True,
            data_only=False,
        )
        self._skip_sheets = {sheet.casefold() for sheet in skip_sheet}
        self._field_info_rows = FieldInfoRows(*field_info_rows)

    def iter_sections(self) -> Iterator[str]:
        for title in self._wb.sheetnames:
            if title.casefold() not in self._skip_sheets:
                yield title

    @functools.cache  # noqa: B019
    def _used_columns(self, section: str, rows: FieldInfoRows) -> list[int]:
        ws = self._wb[section]
        columns = []
        for col in range(1, ws.max_column + 1):
            values = [ws.cell(row=row, column=col).value for row in rows]
            if any(value is not None for value in values):
                columns.append(col)
        return columns

    @staticmethod
    def _cell_text(value):
        return "" if value is None else str(value)

    def iter_fields(
        self,
        section: str,
        info: FieldInfoRows | None = None,
    ) -> Iterator[FieldInfo]:
        info_rows: FieldInfoRows = (
            info if info is not None else self._field_info_rows
        )

        ws = self._wb[section]
        for col in self._used_columns(section, info_rows):
            yield FieldInfo(
                title=self._cell_text(
                    ws.cell(row=info_rows.title, column=col).value
                ),
                field_id=self._cell_text(
                    ws.cell(row=info_rows.field_id, column=col).value
                ),
                mandatory=self._cell_text(
                    ws.cell(row=info_rows.mandatory, column=col).value
                ),
                description=self._cell_text(
                    ws.cell(row=info_rows.description, column=col).value
                ),
                validity_range=self._cell_text(
                    ws.cell(row=info_rows.validity_range, column=col).value
                ),
            )


class BaseFormatter:
    _title_spec: Mapping[int, tuple[str, bool]] = types.MappingProxyType(
        {
            1: ("#", True),
            2: ("*", True),
            3: ("=", False),
            4: ("-", False),
            5: ("^", False),
            6: ('"', False),
        }
    )

    @classmethod
    def make_title(cls, title: str, level: int) -> list[str]:
        """Return lines of a RST formatted title of the specified level."""
        char, overline = cls._title_spec[level]
        underline_str = f"{char * len(title)}"
        lines = [underline_str] if overline else []
        lines.extend([title, f"{char * len(title)}"])
        return lines

    @staticmethod

    @staticmethod
    def literal_block(text, indent: int | str = 4):
        """Generate lines of a RST literal block."""
        prefix = indent if isinstance(indent, str) else " " * indent
        lines = []
        lines.append("::")
        lines.append("")
        lines.extend(textwrap.indent(text, prefix=prefix).splitlines())

        return lines

    @staticmethod
    def field_block(label: str, value) -> list[str]:
        return [f":{label}: {value}"]

    @abc.abstractmethod
    def format_section(
        self, title: str, fields: Iterator[FieldInfo], level: int = 1
    ) -> list[str]:
        """Generate lines of a RST section."""


class SimpleFormatter(BaseFormatter):
    def format_description(self, description: str, level: int = 3) -> list[str]:
        lines = []
        lines.extend(self.make_title("Description", level))
        lines.append("")
        for paragraph in re.split(r"\n\W*\n", description, flags=re.MULTILINE):
            lines.extend(textwrap.wrap(paragraph, MAX_LINE_LEN))
        return lines

    def format_validity_range(
        self, validity_range: str, level: int = 3
    ) -> list[str]:
        lines = []
        lines.extend(self.make_title("Constraints", level))
        lines.append("")
        lines.extend(self.literal_block(validity_range))
        return lines

    def format_field(self, field_info: FieldInfo, level: int = 2) -> list[str]:
        lines = []

        column_title = field_info.title or field_info.field_id
        lines.extend(self.make_title(column_title, level))
        lines.append("")
        lines.extend(self.field_block("Field name", field_info.field_id))
        lines.extend(self.field_block("Required", field_info.mandatory))
        lines.append("")

        if field_info.description:
            lines.append("")
            lines.extend(
                self.format_description(field_info.description, level + 1)
            )
            lines.append("")

        if field_info.validity_range:
            lines.append("")
            lines.extend(
                self.format_validity_range(field_info.validity_range, level + 1)
            )
            lines.append("")

        return lines

    def format_section(
        self, title: str, fields: Iterator[FieldInfo], level: int = 1
    ) -> list[str]:
        """Generate lines of a RST section."""
        lines = []

        lines.extend(self.make_title(title, level))
        lines.append("")

        for field_info in fields:
            lines.extend(self.format_field(field_info, level=level + 1))
            lines.append("")

        return lines


class BaseDocGenerator(abc.ABC):
    def __init__(
        self,
        model: DataModel,
        title: str,
        *,
        preamble: Collection[str] | None = None,
        formatter: BaseFormatter | None = None,
    ):
        self.model = model
        self.title = title
        self.preamble: list[str] = (
            list(preamble) if preamble is not None else []
        )
        self.formatter = (
            formatter if formatter is not None else SimpleFormatter()
        )

    @abc.abstractmethod
    def generate_doc(self, out_path: str) -> None:
        """Generate the entire document."""


class SingleFileDocGenerator(BaseDocGenerator):
    def _format_doc(self) -> list[str]:
        lines = []
        lines.extend(self.formatter.make_title(self.title, 1))
        lines.append("")
        if self.preamble:
            lines.extend(self.preamble)
            lines.append("")

        for section in self.model.iter_sections():
            lines.append("")
            section_lines = self.formatter.format_section(
                section, fields=self.model.iter_fields(section), level=2
            )
            lines.extend(section_lines)

        return lines

    def generate_doc(self, out_path: str) -> None:
        """Generate the string of the entire document."""
        lines = self._format_doc()
        data = "\n".join(lines)
        pathlib.Path(out_path).write_text(data)


def title_from_source(source: pathlib.Path):
    stem = source.stem
    stem = re.sub(r"_v\d+(?:\.\d+)*$", "", stem)
    stem = stem.replace("_submission_template", " Submission Template")
    words = stem.replace("_", " ").replace("-", " ").split()
    return " ".join(
        "SARCalNet" if word.lower() == "sarcalnet" else word.title()
        for word in words
    )


def parse_args():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("source", type=pathlib.Path, help="Input XSLX file.")
    parser.add_argument(
        "--output",
        type=pathlib.Path,
        help="Output .rst file. Defaults to '<source-stem>.rst'.",
    )
    parser.add_argument(
        "--title",
        help=(
            "Document title. Defaults to a readable title from the source "
            "file name."
        ),
    )
    parser.add_argument(
        "--skip-sheet",
        action="append",
        default=list(DEFAULT_SKIP_SHEETS),
        help="Sheet name to skip. Can be repeated. Defaults: %(default)s.",
    )
    parser.add_argument(
        "--title-row",
        type=int,
        default=DEFAULT_INFO_ROWS.title,
        help="Row containing human-readable column titles.",
    )
    parser.add_argument(
        "--field-id-row",
        type=int,
        default=DEFAULT_INFO_ROWS.field_id,
        help="Row containing field IDs.",
    )
    parser.add_argument(
        "--mandatory-row",
        type=int,
        default=DEFAULT_INFO_ROWS.mandatory,
        help="Row in which it is specified if a field is mandatory or not.",
    )
    parser.add_argument(
        "--description-row",
        type=int,
        default=DEFAULT_INFO_ROWS.description,
        help="Row containing the description text.",
    )
    parser.add_argument(
        "--validity-range-row",
        type=int,
        default=DEFAULT_INFO_ROWS.validity_range,
        help="Row containing the definition of validity ranges.",
    )
    try:
        import argcomplete

        argcomplete.autocomplete(parser)
    except ImportError:
        pass
    return parser.parse_args()


def main():
    args = parse_args()
    source = args.source
    output = args.output or pathlib.Path(f"{source.stem}.rst")
    doc_title = args.title or title_from_source(source)
    skip_sheets = {sheet.casefold() for sheet in args.skip_sheet}
    field_info_rows = FieldInfoRows(
        title=args.title_row,
        field_id=args.field_id_row,
        mandatory=args.mandatory_row,
        description=args.description_row,
        validity_range=args.validity_range_row,
    )

    model = DataModel(
        source, skip_sheet=skip_sheets, field_info_rows=field_info_rows
    )

    preamble = [f"Source file: ``{source.name}``"]
    generator = SingleFileDocGenerator(
        model, doc_title, preamble=preamble, formatter=SimpleFormatter()
    )
    generator.generate_doc(output)

    print(output.resolve())


if __name__ == "__main__":
    main()
