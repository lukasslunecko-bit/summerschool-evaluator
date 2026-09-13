from __future__ import annotations

import importlib.util
import shutil
import sys
import tempfile
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
MODULE_PATH = ROOT / "dgt_test_run_evaluator_v0.1.0.py"
SPEC = importlib.util.spec_from_file_location("dgt_test_run_evaluator", MODULE_PATH)
assert SPEC and SPEC.loader
evaluator = importlib.util.module_from_spec(SPEC)
sys.modules[SPEC.name] = evaluator
SPEC.loader.exec_module(evaluator)


class NormalisationTests(unittest.TestCase):
    def test_pipe_header_wins_over_trailing_commas(self) -> None:
        self.assertEqual(evaluator.detect_delimiter("SEGMENT|ORI|TRA,,,,,,\n"), "|")

    def test_segment_lists_and_ranges(self) -> None:
        self.assertEqual(evaluator.parse_segment_tokens("2, 5-7", -1), {"1", "4", "5", "6"})

    def test_project_filename_variants_match(self) -> None:
        self.assertTrue(evaluator.filename_matches(
            "jirkoma - ENV-2026-01157-00-00-FR-SDLXLIFF-00 - Adjusted.csv",
            "ENV-2026-01157-00-00-FR",
        ))

    def test_seed_without_segment_matches_exact_text(self) -> None:
        seed = evaluator.SeededError(
            error_id="1", tester="user", filename="CLIMA-2026-00325.tmx",
            segment="", ori="Source text", correct_tra="Correct text",
            modified_tra="Modified text", comment="", changed_words="", similar_segments="",
        )
        method = evaluator.match_finding(
            seed,
            finding_filename="CLIMA-2026-00325 - Adjusted.csv",
            default_filename="CLIMA-2026-00325 - Adjusted.csv",
            segments=set(), ori="Source text", modified_tra="Modified text",
            modified_lookup={},
        )
        self.assertEqual(method, "modified_text+filename")

    def test_nested_json_column_suffix_is_detected(self) -> None:
        headers = ["location.segment", "diagnostic.message"]
        self.assertEqual(evaluator.resolve_column(headers, "segment"), "location.segment")
        self.assertEqual(evaluator.resolve_column(headers, "comment"), "diagnostic.message")


class WorkbookIntegrationTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.register = ROOT / "registers" / "Project-4-dataset-register.xlsx"
        cls.golden = ROOT / "GRP-LTC Summer School 2026 - Project 4" / "Project data" / "1-golden-standard-files" / "jirkoma - ENV-2026-01157-00-00-FR-SDLXLIFF-00 - Golden standard.csv"
        cls.modified = ROOT / "GRP-LTC Summer School 2026 - Project 4" / "Project data" / "2-modified-files" / "jirkoma - ENV-2026-01157-00-00-FR-SDLXLIFF-00 - Adjusted.csv"

    def setUp(self) -> None:
        for path in (self.register, self.golden, self.modified):
            if not path.exists():
                self.skipTest(f"Workspace fixture is unavailable: {path}")

    def test_append_and_duplicate_guard(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            temporary = Path(directory)
            findings = temporary / "checker-results.csv"
            findings.write_text(
                "segment|message|severity\n"
                "37|Known seeded issue|high\n"
                "999|Unseeded issue|low\n",
                encoding="utf-8",
            )
            output = temporary / "updated-register.xlsx"
            result = evaluator.evaluate(evaluator.EvaluationOptions(
                golden=self.golden,
                modified=self.modified,
                register=self.register,
                findings=findings,
                output_register=output,
                update_in_place=False,
                tester="test-user",
                tool_name="test-checker",
            ))
            self.assertEqual(result.true_positives, 1)
            self.assertEqual(result.false_positives, 1)
            self.assertTrue(output.exists())
            workbook = evaluator.XlsxRegister(output)
            self.assertEqual(len(workbook.sheet_rows("findings")), 3)
            self.assertEqual(len(workbook.sheet_rows("evaluation")), 2)
            with self.assertRaises(evaluator.DuplicateRunError):
                evaluator.evaluate(evaluator.EvaluationOptions(
                    golden=self.golden,
                    modified=self.modified,
                    register=output,
                    findings=findings,
                    dry_run=True,
                ))

    def test_register_header_aliases(self) -> None:
        other = ROOT / "registers" / "Project-1-team-2-dataset-register.xlsx"
        if not other.exists():
            self.skipTest("Project 1 Team 2 fixture is unavailable")
        seeds = evaluator.seeded_errors_from_register(evaluator.XlsxRegister(other))
        self.assertGreater(len(seeds), 0)
        self.assertEqual(seeds[0].error_id, "1")
        self.assertTrue(seeds[0].changed_words)

    def test_in_place_update_creates_recoverable_backup(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            temporary = Path(directory)
            register_copy = temporary / "register.xlsx"
            shutil.copy2(self.register, register_copy)
            findings = temporary / "findings.csv"
            findings.write_text("segment|message\n37|Known seeded issue\n", encoding="utf-8")
            result = evaluator.evaluate(evaluator.EvaluationOptions(
                golden=self.golden,
                modified=self.modified,
                register=register_copy,
                findings=findings,
                update_in_place=True,
                create_backup=True,
            ))
            self.assertIsNotNone(result.backup)
            assert result.backup is not None
            self.assertTrue(result.backup.exists())
            self.assertEqual(evaluator.XlsxRegister(result.backup).sheet_rows("findings"), [])
            self.assertEqual(len(evaluator.XlsxRegister(register_copy).sheet_rows("findings")), 2)

    def test_empty_seeded_register_is_supported(self) -> None:
        empty_register = ROOT / "registers" / "Project-2-team-2-dataset-register.xlsx"
        if not empty_register.exists():
            self.skipTest("Project 2 Team 2 fixture is unavailable")
        self.assertEqual(evaluator.seeded_errors_from_register(evaluator.XlsxRegister(empty_register)), [])


if __name__ == "__main__":
    unittest.main()
