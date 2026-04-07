# Harness Setup

This folder now includes a reusable eval harness based on the harness technique:

1. Define cases as data (`harness/evals/*.eval.json`)
2. Run all cases automatically
3. Grade each case with deterministic checks
4. Save machine-readable reports (`harness/reports/*.json`)

## Quick Start

```bash
npm run harness
```

CI-style run with 3 attempts per case:

```bash
npm run harness:ci
```

## File Layout

```text
package.json
HARNESS.md
harness/
  run-harness.mjs
  evals/
    smoke.eval.json
  reports/
```

## Eval Case Schema

Each case in `*.eval.json` supports:

- `id`: unique case id
- `description`: optional human description
- `command`: command string or string array (array recommended)
- `expect.exitCode`: expected exit code (default `0`)
- `expect.stdoutIncludes`: all tokens must exist in stdout
- `expect.stdoutNotIncludes`: tokens must not exist in stdout
- `expect.stderrIncludes`: all tokens must exist in stderr
- `expect.stderrNotIncludes`: tokens must not exist in stderr

## Why This Helps

- Clear pass/fail criteria before implementation
- Repeatable and auditable validation
- pass@k and pass^k reliability metrics
- Easy regression checks as project files are added
