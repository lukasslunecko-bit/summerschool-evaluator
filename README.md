# DGT Test Run Evaluator 0.1.0

`dgt_test_run_evaluator_v0.1.0.py` evaluates the output of a checking tool against a Summer School dataset register.

It reads:

- one golden-standard CSV or TSV;
- one modified CSV or TSV;
- one dataset-register XLSX workbook;
- one findings file in CSV, TSV, JSON, or XML.

It adds or appends two worksheets in the register:

- `findings`: one detailed row per tool finding, including the segment text, match result, matched seeded-error IDs, and source finding data;
- `evaluation`: one row per test run with the file name, processing time, true positives, false positives, false negatives, precision, recall, and file hash.

The tool uses only the Python standard library. It has the same basic operating pattern as the DGT Data Converter: a standalone file, a Tkinter desktop interface, a headless command-line mode, saved settings, clear logs, and atomic output writes.

## Desktop use

Run or double-click:

```text
dgt_test_run_evaluator_v0.1.0.py
```

Select the four input files, enter the checker name, and click **Evaluate and update register**.

The selected register is updated in place by default. A timestamped backup is created beside it before replacement. The golden CSV, modified CSV, and findings file are never changed.

The advanced mapping fields normally stay on `auto`. Enter an exact findings-file header only when automatic detection does not recognise that column. Use `none` to disable a mapping.

## Headless use

```powershell
python dgt_test_run_evaluator_v0.1.0.py --headless `
  --golden "golden.csv" `
  --modified "modified.csv" `
  --register "Project-4-dataset-register.xlsx" `
  --findings "checker-findings.csv" `
  --tool-name "Checker name" `
  --tester "username"
```

To save an updated copy instead of replacing the selected register:

```powershell
python dgt_test_run_evaluator_v0.1.0.py --headless `
  --golden "golden.csv" `
  --modified "modified.csv" `
  --register "dataset-register.xlsx" `
  --findings "checker-findings.json" `
  --output-register "dataset-register-evaluated.xlsx"
```

Use `--dry-run` to calculate and print the metrics without writing a workbook. Run `python dgt_test_run_evaluator_v0.1.0.py --help` for all options.

## Findings-column detection

Common headers are detected automatically, including:

- segment: `segment`, `segment_number`, `segment_id`, `row`, `line`, `unit_id`;
- file: `filename`, `file`, `source_file`, `document`;
- message: `finding`, `issue`, `message`, `comment`, `description`, `details`;
- severity: `severity`, `level`, `priority`, `risk`, `category`;
- text: `ORI`, `source_text`, `TRA`, `target_text`, `translation`, `modified_TRA`.

Nested JSON keys such as `location.segment` are also recognised. XML finding attributes and child elements become columns.

If a checker reports physical CSV line numbers instead of segment numbers, set a segment offset. For example, use `-1` when line 1 is the CSV header and line 2 represents segment 1.

## Matching rules

A tool finding is counted as a true positive when it matches a seeded error using one of these exact rules:

1. compatible filename plus intersecting segment number;
2. intersecting segment number plus exact modified text;
3. for seeded rows without segment numbers, exact modified text plus compatible filename or exact ORI text;
4. for findings without a segment number, compatible filename plus exact modified text.

Text matching normalises Unicode presentation forms, curly/straight quotes, case, non-breaking spaces, and repeated whitespace. It does not use fuzzy semantic matching. Segment lists and short ranges such as `12, 18-20` are supported.

True positives and false positives count finding rows. `Unique_seeded_errors_matched` counts distinct ground-truth rows, which prevents duplicate tool findings from inflating recall. Seeded errors in scope are selected by compatible filename or exact source/modified segment text.

## Repeat-run protection

The `evaluation` sheet stores the SHA-256 hash of every findings file. Reprocessing identical content is blocked by default. The desktop UI asks for explicit confirmation; headless mode requires `--allow-duplicate`.

## Current boundary

The evaluator supports row-oriented CSV, TSV, JSON, and XML findings. If a future checker returns a materially different structure, use the explicit column mappings or extend the input adapter while keeping the workbook and matching layers unchanged.
