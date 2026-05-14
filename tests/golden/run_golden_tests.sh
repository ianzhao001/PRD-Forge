#!/usr/bin/env bash
set -u

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/../.." && pwd)"
GOLDEN_DIR="$ROOT_DIR/tests/golden"
FIXTURE_DIR="$GOLDEN_DIR/fixtures"
BASELINE_DIR="$GOLDEN_DIR/baseline"
VALIDATOR="$ROOT_DIR/scripts/validate_prd.py"

TMP_DIR="$(mktemp -d "${TMPDIR:-/tmp}/prd-forge-golden.XXXXXX")"
trap 'rm -rf "$TMP_DIR"' EXIT

normalize_json() {
  local input="$1"
  local output="$2"
  python3 - "$input" "$output" <<'PY'
import json
import sys
from pathlib import Path

src = Path(sys.argv[1])
dst = Path(sys.argv[2])
text = src.read_text(encoding="utf-8")
decoder = json.JSONDecoder()
payload, _ = decoder.raw_decode(text.lstrip())
payload.setdefault("meta", {})
payload["meta"].pop("path", None)
payload["meta"].pop("manifest", None)
dst.write_text(json.dumps(payload, ensure_ascii=False, indent=2, sort_keys=True) + "\n", encoding="utf-8")
PY
}

run_case() {
  local label="$1"
  local baseline_name="$2"
  local docx="$3"
  local strict="$4"
  local manifest="${5:-}"

  local raw="$TMP_DIR/$baseline_name.raw"
  local actual="$TMP_DIR/$baseline_name.actual"
  local expected="$TMP_DIR/$baseline_name.expected"
  local baseline="$BASELINE_DIR/$baseline_name"

  local cmd=(python3 "$VALIDATOR" "$FIXTURE_DIR/$docx")
  if [[ "$strict" == "strict" ]]; then
    cmd+=(--strict)
  fi
  if [[ -n "$manifest" ]]; then
    cmd+=(--manifest "$FIXTURE_DIR/$manifest")
  fi

  "${cmd[@]}" >"$raw" 2>&1
  normalize_json "$raw" "$actual"
  normalize_json "$baseline" "$expected"

  if diff -u "$expected" "$actual" >"$TMP_DIR/$baseline_name.diff"; then
    echo "[PASS] $label"
    return 0
  fi

  echo "[FAIL] $label"
  cat "$TMP_DIR/$baseline_name.diff"
  return 1
}

failures=0

run_case "good_sample@strict:no_manifest" "good_sample.strict.no_manifest.json" "good_sample.docx" "strict" || failures=$((failures + 1))
run_case "good_sample_zh@strict:no_manifest" "good_sample_zh.strict.no_manifest.json" "good_sample_zh.docx" "strict" || failures=$((failures + 1))
run_case "good_sample_with_manifest@strict:manifest" "good_sample_with_manifest.strict.manifest.json" "good_sample_with_manifest.docx" "strict" "good_sample.manifest.yaml" || failures=$((failures + 1))
run_case "bad_sample@default:no_manifest" "bad_sample.default.no_manifest.json" "bad_sample.docx" "default" || failures=$((failures + 1))
run_case "bad_sample@strict:no_manifest" "bad_sample.strict.no_manifest.json" "bad_sample.docx" "strict" || failures=$((failures + 1))
run_case "bad_sample@default:manifest" "bad_sample.default.manifest.json" "bad_sample.docx" "default" "bad_manifest.yaml" || failures=$((failures + 1))
run_case "bad_sample@strict:manifest" "bad_sample.strict.manifest.json" "bad_sample.docx" "strict" "bad_manifest.yaml" || failures=$((failures + 1))

if [[ "$failures" -ne 0 ]]; then
  echo "$failures golden test(s) failed"
  exit 1
fi

echo "All golden tests passed"
