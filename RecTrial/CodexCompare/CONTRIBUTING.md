# Contributing Guide

## 1) Development flow

1. Read `README.md`, `PLAN.md`, and `CONSTRAINTS.md` before making changes.
2. Keep sample files in `samples/` unchanged.
3. Build features in small, testable commits.
4. Run smoke checks before opening or updating a PR.

## 2) Required checks

Run one of the following from repository root:

```bash
bash scripts/run_stage_smoke.sh
```

or:

```bash
make smoke
```

## 3) Commit expectations

- Use clear commit titles.
- Keep user-facing text in plain English.
- Keep demo outputs branded using `BRAND.md`.

## 4) Common commands

```bash
make py-compile
make unit
make smoke
```

## 5) Pull request checklist

- [ ] Smoke checks pass
- [ ] New files have clear purpose and naming
- [ ] Documentation updated for user-facing changes
- [ ] No modifications to `samples/` files
