import unittest

from adapters.excel.reader import _header_row, _partial_header_row


class StreamingOnlySheet:
    """Exercise the shared .xlsx/.xls iterator interface without allowing cell reparses."""

    def __init__(self, rows):
        self.rows=rows
        self.max_row=len(rows)
        self.max_column=max((len(row) for row in rows), default=0)
        self.rows_read=0

    def cell(self, *args, **kwargs):
        raise AssertionError("header scan must not use random worksheet.cell access")

    def iter_rows(self, min_row=1, values_only=False):
        assert min_row == 1 and values_only
        for row in self.rows:
            self.rows_read+=1
            yield tuple(row)


class HeaderStreamingTests(unittest.TestCase):
    def test_full_header_preserves_row_index_and_names(self):
        rows=[["제목"], [None], ["거래처코드", "거래처명", "적요", "금액"]]
        ws=StreamingOnlySheet(rows)
        groups=[("거래처코드",), ("거래처명",), ("적요",), ("금액",)]
        self.assertEqual(_header_row(ws, groups), (3, ["거래처코드", "거래처명", "적요", "금액"]))
        self.assertEqual(ws.rows_read, 3)

    def test_partial_header_is_bounded_and_preserves_score(self):
        rows=[["표지"]]*28 + [["거래처코드", "거래처명", "적요"], ["거래처코드", "거래처명"]] + [["거래처코드", "거래처명", "적요", "금액"]]
        ws=StreamingOnlySheet(rows)
        groups=[("거래처코드",), ("거래처명",), ("적요",), ("금액",)]
        self.assertEqual(_partial_header_row(ws, groups, 3), (29, 3))
        self.assertEqual(ws.rows_read, 30)
        self.assertEqual(_header_row(StreamingOnlySheet(rows), groups), (None, []))

    def test_header_aliases_and_unmatched_sheet(self):
        groups=[("거래처코드", "코드"), ("거래처명", "업체명")]
        self.assertEqual(_header_row(StreamingOnlySheet([["코드", "업체명"]]), groups), (1, ["코드", "업체명"]))
        self.assertEqual(_partial_header_row(StreamingOnlySheet([["날짜", "잔액"]]), groups, 1), (None, 0))


if __name__ == "__main__":
    unittest.main()
