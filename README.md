# PRD Forge

PRD Forge is an engineering-grade Codex skill for generating structured Product Requirements Documents (PRDs) from requirements, design drafts, interaction specs, and meeting notes.

It includes:

- PRD writing workflow and formatting references
- screenshot manifest tracking
- `.docx` validator with structured JSON reports
- golden-test regression guardrails

## Usage

Use the skill when you need to write or validate a PRD. For generated documents, run:

```bash
python scripts/validate_prd.py <PRD.docx> [--strict] [--manifest screenshot_manifest.yaml]
```

When changing the skill, validator, schema, or output rules, run:

```bash
bash tests/golden/run_golden_tests.sh
```

## License

Apache-2.0
