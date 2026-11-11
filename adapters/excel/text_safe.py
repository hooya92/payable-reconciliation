"""Keep user-supplied text from becoming a live Excel formula in generated sheets."""


def force_text_formulas(sheet):
    """Turn every formula cell of a generated sheet into plain text.

    Only for sheets this program builds from scratch (review lists, audit trails),
    where no formula is intended: openpyxl treats any string starting with "="
    as a formula, so source text such as "=1+1" or "=SUM(" would otherwise be
    calculated or make Excel ask to repair the file.
    """
    for row in sheet.iter_rows():
        for cell in row:
            if cell.data_type=="f":
                cell.data_type="s"


def set_text(cell, value):
    """Write a user-supplied value; a leading "=" stays text, not a formula."""
    cell.value=value
    if isinstance(value,str) and value.startswith("="):
        cell.data_type="s"
