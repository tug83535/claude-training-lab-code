# Finance Copilot (Production-Ready MVP)

A practical finance automation starter built from the RecTrial roadmap.

## What this package includes

1. **Data Contract Checker**
2. **Exception Triage Engine**
3. **Control Evidence Pack Generator**
4. **CFO One-Page Pulse Report**
5. **Adoption Telemetry Logger** (automatic CSV logging for every CLI run)

## Setup

```bash
cd CodexApril29/finance_copilot
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Commands

### Data Contract Checker
```bash
python finance_copilot.py data-contract \
  --input ./examples/input_dataset.csv \
  --contract ./templates/data_contract_template.json \
  --output-dir ./output
```

### Exception Triage Engine
```bash
python finance_copilot.py triage \
  --input ./examples/exceptions.csv \
  --weights ./templates/triage_weights_template.json \
  --output-dir ./output \
  --top-n 25
```

### Control Evidence Pack
```bash
python finance_copilot.py evidence-pack \
  --input-dir ./output \
  --output-dir ./output \
  --pack-name month_end_controls
```

### CFO One-Page Pulse Report
```bash
python finance_copilot.py cfo-pulse \
  --input ./examples/kpis.csv \
  --thresholds ./templates/cfo_thresholds_template.json \
  --output-dir ./output
```

## Telemetry (auto)
All commands append usage data to:
- default: `./output/tool_usage.csv`
- override with: `--telemetry <path>`

Tracked fields:
- timestamp_utc
- command
- status
- duration_ms
- output_ref
- error_message

## Expected minimum input fields

### Exception triage CSV
- `id`
- `impact` (0..1)
- `confidence` (0..1)
- `days_open`

### CFO pulse CSV
- `kpi`
- `value`

## Testing
```bash
python -m unittest discover -s tests -p 'test_*.py'
python -m py_compile finance_copilot.py tools/*.py
```

## Validation behavior (hardening update)
- `triage` now enforces numeric values and strict bounds:
  - `impact` in [0,1]
  - `confidence` in [0,1]
  - `days_open` >= 0
- `cfo-pulse` now requires numeric `value` entries.
- `evidence-pack` now ignores existing zip artifacts in the output folder when input/output are the same path.
