"""Table operations module for Apple Pages.

CRITICAL: All table operations must go through ``tell page 1`` inside
the document tell block.  Creating or accessing tables directly on the
document object fails in the AppleScript API.

Covers the full sdef range-object API: cell value/formula/formatted value,
range font/size/color/format/alignment/background/text-wrap, row height,
column width, header/footer counts, merge/unmerge/clear/sort/delete/fill.
"""

from pages_cli.utils.pages_backend import _run_applescript, ensure_pages_running
from pages_cli.utils.helpers import _doc_target, _esc


def _table_tell(table_name: str, document: str | None) -> tuple[str, str]:
    """Return (open_block, close_block) for a table tell chain."""
    target = _doc_target(document)
    escaped_table = _esc(table_name)
    open_block = (
        f'tell application "Pages"\n'
        f'  tell {target}\n'
        f'    tell page 1\n'
        f'      tell table "{escaped_table}"\n'
    )
    close_block = (
        f'      end tell\n'
        f'    end tell\n'
        f'  end tell\n'
        f'end tell'
    )
    return open_block, close_block


# ---------------------------------------------------------------------------
# Create
# ---------------------------------------------------------------------------

def add_table(
    rows: int = 3,
    cols: int = 3,
    name: str | None = None,
    header_rows: int = 1,
    header_columns: int = 0,
    footer_rows: int = 0,
    document: str | None = None,
) -> dict:
    """Add a table to page 1 of the document.

    Args:
        rows: Number of rows. Defaults to 3.
        cols: Number of columns. Defaults to 3.
        name: Optional name for the table.
        header_rows: Number of header rows. Defaults to 1.
        header_columns: Number of header columns. Defaults to 0.
        footer_rows: Number of footer rows. Defaults to 0.
        document: Optional document name.

    Returns:
        dict with keys: name, rows, columns, header_rows,
        header_columns, footer_rows.
    """
    ensure_pages_running()
    target = _doc_target(document)

    props = (
        f"row count:{rows}, column count:{cols}, "
        f"header row count:{header_rows}, "
        f"header column count:{header_columns}, "
        f"footer row count:{footer_rows}"
    )
    if name:
        props += f', name:"{_esc(name)}"'

    script = (
        f'tell application "Pages"\n'
        f'  tell {target}\n'
        f'    tell page 1\n'
        f'      set newTable to make new table with properties {{{props}}}\n'
        f'      set tN to name of newTable\n'
        f'      set tR to row count of newTable\n'
        f'      set tC to column count of newTable\n'
        f'      set tHR to header row count of newTable\n'
        f'      set tHC to header column count of newTable\n'
        f'      set tFR to footer row count of newTable\n'
        f'      return tN & "|" & tR & "|" & tC & "|" & tHR & "|" & tHC & "|" & tFR\n'
        f'    end tell\n'
        f'  end tell\n'
        f'end tell'
    )
    result = _run_applescript(script)
    parts = result.strip().split("|")

    return {
        "name": parts[0] if len(parts) > 0 else "",
        "rows": int(parts[1]) if len(parts) > 1 else rows,
        "columns": int(parts[2]) if len(parts) > 2 else cols,
        "header_rows": int(parts[3]) if len(parts) > 3 else header_rows,
        "header_columns": int(parts[4]) if len(parts) > 4 else header_columns,
        "footer_rows": int(parts[5]) if len(parts) > 5 else footer_rows,
    }


# ---------------------------------------------------------------------------
# Delete table
# ---------------------------------------------------------------------------

def delete_table(
    table_name: str | None = None,
    index: int | None = None,
    document: str | None = None,
) -> None:
    """Delete a table by name or index."""
    if table_name is None and index is None:
        raise ValueError("Provide either table_name or index.")
    ensure_pages_running()
    target = _doc_target(document)
    if table_name:
        table_ref = f'table "{_esc(table_name)}"'
    else:
        table_ref = f"table {index}"
    script = (
        f'tell application "Pages"\n'
        f'  tell {target}\n'
        f'    tell page 1\n'
        f'      delete {table_ref}\n'
        f'    end tell\n'
        f'  end tell\n'
        f'end tell'
    )
    _run_applescript(script)


# ---------------------------------------------------------------------------
# Cell access
# ---------------------------------------------------------------------------

def set_cell(
    table_name: str,
    row: int,
    col: int,
    value: str,
    document: str | None = None,
) -> None:
    """Set a cell value (or formula if value starts with '=')."""
    ensure_pages_running()
    o, c = _table_tell(table_name, document)
    escaped_value = _esc(str(value))
    script = f'{o}        set value of cell {col} of row {row} to "{escaped_value}"\n{c}'
    _run_applescript(script)


def get_cell(
    table_name: str,
    row: int,
    col: int,
    document: str | None = None,
) -> str:
    """Get a cell value from a table."""
    ensure_pages_running()
    o, c = _table_tell(table_name, document)
    script = (
        f'{o}'
        f'        set cellVal to value of cell {col} of row {row}\n'
        f'        if cellVal is missing value then\n'
        f'          return ""\n'
        f'        else\n'
        f'          return cellVal as text\n'
        f'        end if\n'
        f'{c}'
    )
    return _run_applescript(script)


def get_cell_formula(
    table_name: str,
    row: int,
    col: int,
    document: str | None = None,
) -> str:
    """Get the formula text of a cell (empty string if no formula)."""
    ensure_pages_running()
    o, c = _table_tell(table_name, document)
    script = (
        f'{o}'
        f'        try\n'
        f'          set f to formula of cell {col} of row {row}\n'
        f'          if f is missing value then return ""\n'
        f'          return f\n'
        f'        on error\n'
        f'          return ""\n'
        f'        end try\n'
        f'{c}'
    )
    return _run_applescript(script)


def get_cell_formatted_value(
    table_name: str,
    row: int,
    col: int,
    document: str | None = None,
) -> str:
    """Get the formatted display value of a cell."""
    ensure_pages_running()
    o, c = _table_tell(table_name, document)
    script = (
        f'{o}'
        f'        try\n'
        f'          set fv to formatted value of cell {col} of row {row}\n'
        f'          if fv is missing value then return ""\n'
        f'          return fv\n'
        f'        on error\n'
        f'          return ""\n'
        f'        end try\n'
        f'{c}'
    )
    return _run_applescript(script)


# ---------------------------------------------------------------------------
# Listing
# ---------------------------------------------------------------------------

def list_tables(document: str | None = None) -> list[dict]:
    """List all tables on page 1 of the document."""
    ensure_pages_running()
    target = _doc_target(document)

    count_script = (
        f'tell application "Pages"\n'
        f'  tell {target}\n'
        f'    tell page 1\n'
        f'      count of tables\n'
        f'    end tell\n'
        f'  end tell\n'
        f'end tell'
    )
    count = int(_run_applescript(count_script).strip())
    if count == 0:
        return []

    tables = []
    for i in range(1, count + 1):
        info_script = (
            f'tell application "Pages"\n'
            f'  tell {target}\n'
            f'    tell page 1\n'
            f'      set t to table {i}\n'
            f'      set tName to name of t\n'
            f'      set tRows to row count of t\n'
            f'      set tCols to column count of t\n'
            f'      set tHR to header row count of t\n'
            f'      set tHC to header column count of t\n'
            f'      set tFR to footer row count of t\n'
            f'      try\n'
            f'        set tCR to name of cell range of t\n'
            f'      on error\n'
            f'        set tCR to ""\n'
            f'      end try\n'
            f'      try\n'
            f'        set tSR to name of selection range of t\n'
            f'      on error\n'
            f'        set tSR to ""\n'
            f'      end try\n'
            f'      return tName & "|" & tRows & "|" & tCols & "|" & tHR & "|" & tHC & "|" & tFR & "|" & tCR & "|" & tSR\n'
            f'    end tell\n'
            f'  end tell\n'
            f'end tell'
        )
        result = _run_applescript(info_script).strip()
        parts = result.split("|")
        tables.append({
            "name": parts[0] if len(parts) > 0 else f"Table {i}",
            "rows": int(parts[1]) if len(parts) > 1 else 0,
            "columns": int(parts[2]) if len(parts) > 2 else 0,
            "header_rows": int(parts[3]) if len(parts) > 3 else 0,
            "header_columns": int(parts[4]) if len(parts) > 4 else 0,
            "footer_rows": int(parts[5]) if len(parts) > 5 else 0,
            "cell_range": parts[6] if len(parts) > 6 else "",
            "selection_range": parts[7] if len(parts) > 7 else "",
        })
    return tables


def get_table_info(
    table_name: str,
    document: str | None = None,
) -> dict:
    """Get detailed info for a single table by name."""
    ensure_pages_running()
    o, c = _table_tell(table_name, document)
    script = (
        f'{o}'
        f'        set tRows to row count\n'
        f'        set tCols to column count\n'
        f'        set tHR to header row count\n'
        f'        set tHC to header column count\n'
        f'        set tFR to footer row count\n'
        f'        set tN to name\n'
        f'        try\n'
        f'          set tCR to name of cell range\n'
        f'        on error\n'
        f'          set tCR to ""\n'
        f'        end try\n'
        f'        try\n'
        f'          set tSR to name of selection range\n'
        f'        on error\n'
        f'          set tSR to ""\n'
        f'        end try\n'
        f'        return tN & "|" & tRows & "|" & tCols & "|" & tHR & "|" & tHC & "|" & tFR & "|" & tCR & "|" & tSR\n'
        f'{c}'
    )
    result = _run_applescript(script).strip()
    parts = result.split("|")
    return {
        "name": parts[0] if len(parts) > 0 else table_name,
        "rows": int(parts[1]) if len(parts) > 1 else 0,
        "columns": int(parts[2]) if len(parts) > 2 else 0,
        "header_rows": int(parts[3]) if len(parts) > 3 else 0,
        "header_columns": int(parts[4]) if len(parts) > 4 else 0,
        "footer_rows": int(parts[5]) if len(parts) > 5 else 0,
        "cell_range": parts[6] if len(parts) > 6 else "",
        "selection_range": parts[7] if len(parts) > 7 else "",
    }


# ---------------------------------------------------------------------------
# Get all data / fill
# ---------------------------------------------------------------------------

def get_table_data(
    table_name: str,
    document: str | None = None,
) -> list[list[str]]:
    """Get all cell values as a 2D array."""
    ensure_pages_running()
    o, c = _table_tell(table_name, document)

    dim_script = (
        f'{o}'
        f'        return (row count as text) & "|" & (column count as text)\n'
        f'{c}'
    )
    dim_result = _run_applescript(dim_script).strip()
    dim_parts = dim_result.split("|")
    row_count = int(dim_parts[0])
    col_count = int(dim_parts[1])

    data = []
    for r in range(1, row_count + 1):
        row_values = []
        for col in range(1, col_count + 1):
            cell_script = (
                f'{o}'
                f'        set cellVal to value of cell {col} of row {r}\n'
                f'        if cellVal is missing value then\n'
                f'          return ""\n'
                f'        else\n'
                f'          return cellVal as text\n'
                f'        end if\n'
                f'{c}'
            )
            try:
                val = _run_applescript(cell_script).strip()
            except RuntimeError:
                val = ""
            row_values.append(val)
        data.append(row_values)
    return data


def fill_table(
    table_name: str,
    data: list[list[str]],
    document: str | None = None,
) -> None:
    """Fill a table from a 2D array of values, starting at row 1 col 1."""
    ensure_pages_running()
    o, c = _table_tell(table_name, document)
    for r_idx, row in enumerate(data, start=1):
        for c_idx, value in enumerate(row, start=1):
            escaped_value = _esc(str(value))
            script = (
                f'{o}'
                f'        set value of cell {c_idx} of row {r_idx} to "{escaped_value}"\n'
                f'{c}'
            )
            _run_applescript(script)


# ---------------------------------------------------------------------------
# Merge / Unmerge / Clear / Sort
# ---------------------------------------------------------------------------

def merge_cells(
    table_name: str,
    range_str: str,
    document: str | None = None,
) -> None:
    """Merge cells in a table range (e.g. "A1:B2")."""
    ensure_pages_running()
    o, c = _table_tell(table_name, document)
    script = f'{o}        merge range "{_esc(range_str)}"\n{c}'
    _run_applescript(script)


def unmerge_cells(
    table_name: str,
    range_str: str,
    document: str | None = None,
) -> None:
    """Unmerge previously merged cells in a table range."""
    ensure_pages_running()
    o, c = _table_tell(table_name, document)
    script = f'{o}        unmerge range "{_esc(range_str)}"\n{c}'
    _run_applescript(script)


def clear_range(
    table_name: str,
    range_str: str,
    document: str | None = None,
) -> None:
    """Clear contents and formatting of a cell range."""
    ensure_pages_running()
    o, c = _table_tell(table_name, document)
    script = f'{o}        clear range "{_esc(range_str)}"\n{c}'
    _run_applescript(script)


def sort_table(
    table_name: str,
    column: int,
    direction: str = "ascending",
    document: str | None = None,
) -> None:
    """Sort a table by a column.

    Args:
        direction: "ascending" or "descending".
    """
    if direction not in ("ascending", "descending"):
        raise ValueError(f"Invalid sort direction '{direction}'.")
    ensure_pages_running()
    o, c = _table_tell(table_name, document)
    sort_order = f"sort order {direction}"
    script = f'{o}        sort by column {column} direction {sort_order}\n{c}'
    _run_applescript(script)


# ---------------------------------------------------------------------------
# Header / Footer counts
# ---------------------------------------------------------------------------

def set_header_rows(
    table_name: str, count: int, document: str | None = None,
) -> None:
    """Set number of header rows (0-5)."""
    ensure_pages_running()
    o, c = _table_tell(table_name, document)
    script = f'{o}        set header row count to {count}\n{c}'
    _run_applescript(script)


def set_header_columns(
    table_name: str, count: int, document: str | None = None,
) -> None:
    """Set number of header columns (0-5)."""
    ensure_pages_running()
    o, c = _table_tell(table_name, document)
    script = f'{o}        set header column count to {count}\n{c}'
    _run_applescript(script)


def set_footer_rows(
    table_name: str, count: int, document: str | None = None,
) -> None:
    """Set number of footer rows (0-5)."""
    ensure_pages_running()
    o, c = _table_tell(table_name, document)
    script = f'{o}        set footer row count to {count}\n{c}'
    _run_applescript(script)


# ---------------------------------------------------------------------------
# Row height / Column width
# ---------------------------------------------------------------------------

def set_row_height(
    table_name: str, row: int, height: float, document: str | None = None,
) -> None:
    """Set the height of a row in points."""
    ensure_pages_running()
    o, c = _table_tell(table_name, document)
    script = f'{o}        set height of row {row} to {height}\n{c}'
    _run_applescript(script)


def get_row_height(
    table_name: str, row: int, document: str | None = None,
) -> float:
    """Get the height of a row in points."""
    ensure_pages_running()
    o, c = _table_tell(table_name, document)
    script = f'{o}        get height of row {row}\n{c}'
    return float(_run_applescript(script).strip())


def get_row_address(
    table_name: str, row: int, document: str | None = None,
) -> str:
    """Get the address of a row."""
    ensure_pages_running()
    o, c = _table_tell(table_name, document)
    script = f'{o}        get address of row {row}\n{c}'
    return _run_applescript(script).strip()


def set_column_width(
    table_name: str, col: int, width: float, document: str | None = None,
) -> None:
    """Set the width of a column in points."""
    ensure_pages_running()
    o, c = _table_tell(table_name, document)
    script = f'{o}        set width of column {col} to {width}\n{c}'
    _run_applescript(script)


def get_column_width(
    table_name: str, col: int, document: str | None = None,
) -> float:
    """Get the width of a column in points."""
    ensure_pages_running()
    o, c = _table_tell(table_name, document)
    script = f'{o}        get width of column {col}\n{c}'
    return float(_run_applescript(script).strip())


def get_column_address(
    table_name: str, col: int, document: str | None = None,
) -> str:
    """Get the address of a column."""
    ensure_pages_running()
    o, c = _table_tell(table_name, document)
    script = f'{o}        get address of column {col}\n{c}'
    return _run_applescript(script).strip()


# ---------------------------------------------------------------------------
# Range properties (font, size, color, format, alignment, background, wrap)
# ---------------------------------------------------------------------------

# Valid cell format constants in Pages AppleScript
CELL_FORMATS = [
    "automatic", "checkbox", "currency", "date and time", "fraction",
    "number", "percent", "pop up menu", "scientific", "slider",
    "stepper", "text", "duration", "rating", "numeral system",
]

ALIGNMENTS = ["auto align", "center", "justify", "left", "right"]

VERTICAL_ALIGNMENTS = ["top", "center", "bottom"]


def set_range_font(
    table_name: str, range_str: str, font_name: str,
    document: str | None = None,
) -> None:
    """Set font name on a cell range."""
    ensure_pages_running()
    o, c = _table_tell(table_name, document)
    script = (
        f'{o}'
        f'        set font name of range "{_esc(range_str)}" to "{_esc(font_name)}"\n'
        f'{c}'
    )
    _run_applescript(script)


def get_range_font(
    table_name: str, range_str: str, document: str | None = None,
) -> str:
    """Get font name of a cell range."""
    ensure_pages_running()
    o, c = _table_tell(table_name, document)
    script = f'{o}        get font name of range "{_esc(range_str)}"\n{c}'
    return _run_applescript(script).strip()


def set_range_font_size(
    table_name: str, range_str: str, size: float,
    document: str | None = None,
) -> None:
    """Set font size on a cell range."""
    ensure_pages_running()
    o, c = _table_tell(table_name, document)
    script = f'{o}        set font size of range "{_esc(range_str)}" to {size}\n{c}'
    _run_applescript(script)


def get_range_font_size(
    table_name: str, range_str: str, document: str | None = None,
) -> float:
    """Get font size of a cell range."""
    ensure_pages_running()
    o, c = _table_tell(table_name, document)
    script = f'{o}        get font size of range "{_esc(range_str)}"\n{c}'
    return float(_run_applescript(script).strip())


def set_range_format(
    table_name: str, range_str: str, format_type: str,
    document: str | None = None,
) -> None:
    """Set the cell format on a range.

    format_type: one of CELL_FORMATS (e.g. "currency", "percent").
    """
    ensure_pages_running()
    o, c = _table_tell(table_name, document)
    script = f'{o}        set format of range "{_esc(range_str)}" to {format_type}\n{c}'
    _run_applescript(script)


def get_range_format(
    table_name: str, range_str: str, document: str | None = None,
) -> str:
    """Get the cell format of a range."""
    ensure_pages_running()
    o, c = _table_tell(table_name, document)
    script = f'{o}        get format of range "{_esc(range_str)}"\n{c}'
    return _run_applescript(script).strip()


def set_range_alignment(
    table_name: str, range_str: str, alignment: str,
    document: str | None = None,
) -> None:
    """Set horizontal alignment on a range.

    alignment: one of "auto align", "center", "justify", "left", "right".
    """
    ensure_pages_running()
    o, c = _table_tell(table_name, document)
    script = f'{o}        set alignment of range "{_esc(range_str)}" to {alignment}\n{c}'
    _run_applescript(script)


def get_range_alignment(
    table_name: str, range_str: str, document: str | None = None,
) -> str:
    """Get horizontal alignment of a range."""
    ensure_pages_running()
    o, c = _table_tell(table_name, document)
    script = f'{o}        get alignment of range "{_esc(range_str)}"\n{c}'
    return _run_applescript(script).strip()


def set_range_vertical_alignment(
    table_name: str, range_str: str, alignment: str,
    document: str | None = None,
) -> None:
    """Set vertical alignment on a range: top, center, bottom."""
    ensure_pages_running()
    o, c = _table_tell(table_name, document)
    script = (
        f'{o}'
        f'        set vertical alignment of range "{_esc(range_str)}" to {alignment}\n'
        f'{c}'
    )
    _run_applescript(script)


def get_range_vertical_alignment(
    table_name: str, range_str: str, document: str | None = None,
) -> str:
    """Get vertical alignment of a range."""
    ensure_pages_running()
    o, c = _table_tell(table_name, document)
    script = f'{o}        get vertical alignment of range "{_esc(range_str)}"\n{c}'
    return _run_applescript(script).strip()


def set_range_text_color(
    table_name: str, range_str: str,
    r: int, g: int, b: int,
    document: str | None = None,
) -> None:
    """Set text color on a range (RGB 0-65535)."""
    ensure_pages_running()
    o, c = _table_tell(table_name, document)
    script = (
        f'{o}'
        f'        set text color of range "{_esc(range_str)}" to {{{r}, {g}, {b}}}\n'
        f'{c}'
    )
    _run_applescript(script)


def get_range_text_color(
    table_name: str, range_str: str, document: str | None = None,
) -> str:
    """Get text color of a range as comma-separated RGB."""
    ensure_pages_running()
    o, c = _table_tell(table_name, document)
    script = f'{o}        get text color of range "{_esc(range_str)}"\n{c}'
    return _run_applescript(script).strip()


def set_range_background_color(
    table_name: str, range_str: str,
    r: int, g: int, b: int,
    document: str | None = None,
) -> None:
    """Set background color on a range (RGB 0-65535)."""
    ensure_pages_running()
    o, c = _table_tell(table_name, document)
    script = (
        f'{o}'
        f'        set background color of range "{_esc(range_str)}" to {{{r}, {g}, {b}}}\n'
        f'{c}'
    )
    _run_applescript(script)


def get_range_background_color(
    table_name: str, range_str: str, document: str | None = None,
) -> str:
    """Get background color of a range as comma-separated RGB."""
    ensure_pages_running()
    o, c = _table_tell(table_name, document)
    script = f'{o}        get background color of range "{_esc(range_str)}"\n{c}'
    return _run_applescript(script).strip()


def set_range_text_wrap(
    table_name: str, range_str: str, wrap: bool,
    document: str | None = None,
) -> None:
    """Set text wrap on a range."""
    ensure_pages_running()
    o, c = _table_tell(table_name, document)
    val = "true" if wrap else "false"
    script = f'{o}        set text wrap of range "{_esc(range_str)}" to {val}\n{c}'
    _run_applescript(script)


def get_range_text_wrap(
    table_name: str, range_str: str, document: str | None = None,
) -> bool:
    """Get text wrap setting of a range."""
    ensure_pages_running()
    o, c = _table_tell(table_name, document)
    script = f'{o}        get text wrap of range "{_esc(range_str)}"\n{c}'
    return _run_applescript(script).strip().lower() == "true"


def get_range_name(
    table_name: str, range_str: str, document: str | None = None,
) -> str:
    """Get the name (coordinates) of a range."""
    ensure_pages_running()
    o, c = _table_tell(table_name, document)
    script = f'{o}        get name of range "{_esc(range_str)}"\n{c}'
    return _run_applescript(script).strip()


# ---------------------------------------------------------------------------
# Cell-level format
# ---------------------------------------------------------------------------

def set_cell_format(
    table_name: str, row: int, col: int, format_type: str,
    document: str | None = None,
) -> None:
    """Set the format of a single cell."""
    ensure_pages_running()
    o, c = _table_tell(table_name, document)
    script = (
        f'{o}'
        f'        set format of cell {col} of row {row} to {format_type}\n'
        f'{c}'
    )
    _run_applescript(script)
