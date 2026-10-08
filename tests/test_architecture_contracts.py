"""Protect dependency direction so future feature work cannot tangle modules."""
import ast
import unittest
from pathlib import Path


ROOT=Path(__file__).resolve().parents[1]


def imports(path):
    tree=ast.parse((ROOT/path).read_text(encoding="utf-8"))
    targets=[]
    for node in ast.walk(tree):
        if isinstance(node,ast.Import):
            targets.extend(alias.name for alias in node.names)
        elif isinstance(node,ast.ImportFrom):
            targets.append(node.module or "")
    return targets


class DependencyContractTests(unittest.TestCase):
    def test_domain_never_imports_ui_excel_or_application(self):
        for path in (ROOT/"domain").glob("*.py"):
            with self.subTest(module=path.name):
                self.assertFalse([
                    item for item in imports(path.relative_to(ROOT))
                    if item.startswith(("app","adapters","application","openpyxl","tkinter"))
                ])

    def test_planner_has_no_excel_ui_or_reader_dependencies(self):
        self.assertFalse([
            item for item in imports("application/month_end_plan.py")
            if item.startswith(("app","adapters","openpyxl","tkinter"))
        ])

    def test_exporter_and_service_do_not_import_ui_or_one_another(self):
        for path in ("application/service.py","adapters/excel/month_end_statement.py"):
            with self.subTest(path=path):
                self.assertFalse([
                    item for item in imports(path)
                    if item == "app" or item.startswith("tkinter")
                ])
        self.assertNotIn("adapters.excel.month_end_statement",
                         imports("application/service.py"))

    def test_preview_ci_watches_all_runtime_dependencies_once(self):
        workflow=(ROOT/".github/workflows/month-end-preview.yml").read_text(encoding="utf-8")
        for fragment in ('"application/**"','"domain/**"',
                         '"adapters/excel/**"','"requirements.txt"'):
            self.assertIn(fragment,workflow)
        self.assertNotIn("  pull_request:",workflow)


if __name__=="__main__":
    unittest.main()
