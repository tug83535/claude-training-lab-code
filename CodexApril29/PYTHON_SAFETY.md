# Python Safety Standard (Finance Toolkit v1)

These are the operating safeguards for any new Python automation added in this project direction.

## Non-Negotiable Rules

1. No internet/API calls.
2. No credential or secret handling.
3. No live DB connection in v1.
4. Never modify source input files in-place.
5. Never delete user files.
6. Write outputs to timestamped run folders.
7. Always create run logs.
8. Include sample mode for safe demo/testing.
9. Show friendly user-facing error messages.
10. Put technical detail in logs.
11. Prefer stdlib + pandas + openpyxl first.
12. Keep tool behavior deterministic and reproducible.

## Required Runtime Behaviors

Every Python tool should support:

- `--input` (or equivalent file/folder argument)
- `--output-dir` (or equivalent)
- `--sample` mode
- clear completion summary
- clear output path display

## Output Structure Recommendation

For each run:

- `outputs/<tool_name>/<YYYYMMDD_HHMMSS>/`
  - `result.*`
  - `summary.md`
  - `run.log`
  - optional `manifest.csv`

## Error Handling Standard

- If user error (missing column/path): explain what to fix in plain English.
- If technical exception: write traceback to log and keep terminal message concise.
- Never fail silently.

## Safety Validation Checklist

Before accepting any tool:

- [ ] Did not overwrite source file
- [ ] Did not delete source file
- [ ] Created timestamped output folder
- [ ] Created run log
- [ ] Works in sample mode
- [ ] No internet/API usage
- [ ] No credentials required
