"""Month-end draft export that never stops only because the template cannot be kept."""
from adapters.excel.month_end_statement import write_month_end_statement
from adapters.excel.writer import write_result


def export_month_end_draft(path, prior_paths, run, period, source_paths):
    """Return (kind, reason).

    kind "template": the original statement format was preserved (reason is "").
    kind "plain": the format could not be preserved (e.g. .xls, merged body cells,
    protected sheet, unreadable amount in a detail row), so the reconciliation
    result is saved in the standard result workbook instead. The reason is also
    written into its 요약 sheet. Real failures (locked files, writing over an
    input file) are not hidden: PermissionError passes through, and if the plain
    export fails too, the original reason is raised.
    """
    try:
        write_month_end_statement(
            path,prior_paths,run.results,run.new_items,run.issues,period,
            standalone_debits=run.standalone_debits,
            standalone_debit_sources=run.standalone_debit_sources,
            journal_items=run.journal_items,
            journal_sources=run.journal_sources,
            source_paths=source_paths,
        )
        return "template",""
    except ValueError as exc:
        reason=str(exc)
        notice=("전월 명세서의 양식을 그대로 유지할 수 없어 대사 결과 기준의 간이 초안으로 저장했습니다. "
                f"사유: {reason}")
        try:
            write_result(path,run.results,run.new_items,run.issues,
                         [*prior_paths,*source_paths],period.label,notice)
        except ValueError:
            raise exc
        return "plain",reason
