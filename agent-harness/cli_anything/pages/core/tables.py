"""Table operations module for Apple Pages."""

from cli_anything.pages.utils.pages_backend import _run_applescript, ensure_pages_running


def _doc_target(document: str | None) -> str:
    """Return the AppleScript target reference for a document."""
    if document:
        escaped = document.replace('"', '\\"')
        return f'document "{escaped}"'
    return "front document"


def add_table(
    rows: int = 3,
    cols: int = 3,
    name: str | None = None,
    document: str | None = None,
) -> dict:
    """Add a table to the document.

    Args:
        rows: Number of rows. Defaults to 3.
        cols: Number of columns. Defaults to 3.
        name: Optional name for the table.
        document: Optional document name.

    Returns:
        dict: Table info with keys: name, rows, columns.

    Raises:
        RuntimeError: If the table cannot be created.
    """
    ensure_pages_running()
    target = _doc_target(document)

    props = f"row count:{rows}, column count:{cols}"
    if name:
        escaped_name = name.replace('"', '\\"')
        props += f', name:"{escaped_name}"'

    script = (
        f'tell application "Pages"\n'
        f'  tell {target}\n'
        f'    set newTable to make new table with properties {{{props}}}\n'
        f'    set tableName to name of newTable\n'
        f'    set tableRows to row count of newTable\n'
        f'    set tableCols to column count of newTable\n'
        f'    return tableName & "|" & tableRows & "|" & tableCols\n'
        f'  end tell\n'
        f'end tell'
    )
    result = _run_applescript(script)
    parts = result.strip().split("|")

    return {
        "name": parts[0] if len(parts) > 0 else "",
        "rows": int(parts[1]) if len(parts) > 1 else rows,
        "columns": int(parts[2]) if len(parts) > 2 else cols,
    }


def set_cell(
    table_name: str,
    row: int,
    col: int,
    value: str,
    document: str | None = None,
) -> None:
    """Set a cell value in a table.

    Args:
        table_name: The name of the table.
        row: Row index (1-based).
        col: Column index (1-based).
        value: The value to set.
        document: Optional document name.

    Raises:
        RuntimeError: If the cell cannot be set.
    """
    ensure_pages_running()
    target = _doc_target(document)
    escaped_table = table_name.replace('"', '\\"')
    escaped_value = str(value).replace("\\", "\\\\").replace('"', '\\"')

    script = (
        f'tell application "Pages"\n'
        f'  tell {target}\n'
        f'    tell table "{escaped_table}"\n'
        f'      set value of cell {col} of row {row} to "{escaped_value}"\n'
        f'    end tell\n'
        f'  end tell\n'
        f'end tell'
    )
    _run_applescript(script)


def get_cell(
    table_name: str,
    row: int,
    col: int,
    document: str | None = None,
) -> str:
    """Get a cell value from a table.

    Args:
        table_name: The name of the table.
        row: Row index (1-based).
        col: Column index (1-based).
        document: Optional document name.

    Returns:
        str: The cell value.

    Raises:
        RuntimeError: If the cell value cannot be retrieved.
    """
    ensure_pages_running()
    target = _doc_target(document)
    escaped_table = table_name.replace('"', '\\"')

    script = (
        f'tell application "Pages"\n'
        f'  tell {target}\n'
        f'    tell table "{escaped_table}"\n'
        f'      get value of cell {col} of row {row}\n'
        f'    end tell\n'
        f'  end tell\n'
        f'end tell'
    )
    return _run_applescript(script)


def list_tables(document: str | None = None) -> list[dict]:
    """List all tables in the document.

    Args:
        document: Optional document name.

    Returns:
        list[dict]: List of table info dicts with keys: name, rows, columns.

    Raises:
        RuntimeError: If tables cannot be listed.
    """
    ensure_pages_running()
    target = _doc_target(document)

    count_script = (
        f'tell application "Pages"\n'
        f'  tell {target}\n'
        f'    count of tables\n'
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
            f'    set t to table {i}\n'
            f'    set tName to name of t\n'
            f'    set tRows to row count of t\n'
            f'    set tCols to column count of t\n'
            f'    return tName & "|" & tRows & "|" & tCols\n'
            f'  end tell\n'
            f'end tell'
        )
        result = _run_applescript(info_script).strip()
        parts = result.split("|")
        tables.append({
            "name": parts[0] if len(parts) > 0 else f"Table {i}",
            "rows": int(parts[1]) if len(parts) > 1 else 0,
            "columns": int(parts[2]) if len(parts) > 2 else 0,
        })

    return tables


def merge_cells(
    table_name: str,
    range_str: str,
    document: str | None = None,
) -> None:
    """Merge cells in a table.

    Args:
        table_name: The name of the table.
        range_str: The range to merge (e.g. "A1:B2").
        document: Optional document name.

    Raises:
        RuntimeError: If cells cannot be merged.
    """
    ensure_pages_running()
    target = _doc_target(document)
    escaped_table = table_name.replace('"', '\\"')
    escaped_range = range_str.replace('"', '\\"')

    script = (
        f'tell application "Pages"\n'
        f'  tell {target}\n'
        f'    tell table "{escaped_table}"\n'
        f'      merge range "{escaped_range}"\n'
        f'    end tell\n'
        f'  end tell\n'
        f'end tell'
    )
    _run_applescript(script)


def sort_table(
    table_name: str,
    column: int,
    direction: str = "ascending",
    document: str | None = None,
) -> None:
    """Sort a table by a column.

    Args:
        table_name: The name of the table.
        column: The column index (1-based) to sort by.
        direction: Sort direction, either "ascending" or "descending".
        document: Optional document name.

    Raises:
        RuntimeError: If the table cannot be sorted.
        ValueError: If an invalid direction is provided.
    """
    if direction not in ("ascending", "descending"):
        raise ValueError(
            f"Invalid sort direction '{direction}'. Must be 'ascending' or 'descending'."
        )

    ensure_pages_running()
    target = _doc_target(document)
    escaped_table = table_name.replace('"', '\\"')

    sort_order = "sort order ascending" if direction == "ascending" else "sort order descending"

    script = (
        f'tell application "Pages"\n'
        f'  tell {target}\n'
        f'    tell table "{escaped_table}"\n'
        f'      sort by column {column} direction {sort_order}\n'
        f'    end tell\n'
        f'  end tell\n'
        f'end tell'
    )
    _run_applescript(script)
