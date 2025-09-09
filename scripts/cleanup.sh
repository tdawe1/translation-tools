#!/usr/bin/env bash
set -euo pipefail
shopt -s nullglob dotglob

MODE="${1:-aggressive}"
TS="$(date +%Y%m%d_%H%M%S)"
ROOT_DIR="$(cd "$(dirname "$0")/.." && pwd)"
cd "$ROOT_DIR"

echo "[clean] mode=$MODE at $TS"

# Create a reversible stash for small artifacts
STASH_DIR="archived_runs/cleanup-${TS}"
mkdir -p "$STASH_DIR"/{logs,audits,csv,caches}

move_if_any() {
  local dest="$1"; shift
  local moved=0
  for f in "$@"; do
    if [[ -e "$f" ]]; then
      mv -f "$f" "$dest"/
      echo "[clean] moved $(basename "$f") -> $dest/"
      moved=1
    fi
  done
  return $moved
}

echo "[clean] Archiving logs, audits, csv, caches (reversible)"
move_if_any "$STASH_DIR/logs" *.log scripts/*.log || true
move_if_any "$STASH_DIR/audits" audit*.json || true
move_if_any "$STASH_DIR/csv" bilingual*.csv residual_rows_translated_only.csv || true

# Keep the active cache, stash the rest
for f in translation_cache*; do
  [[ -e "$f" ]] || continue
  if [[ "$f" == "translation_cache.json" ]]; then
    echo "[clean] keeping $f"
  else
    mv -f "$f" "$STASH_DIR/caches/"
    echo "[clean] moved $f -> $STASH_DIR/caches/"
  fi
done

echo "[clean] Removing __pycache__ and .pytest_cache"
find . -type d \( -name "__pycache__" -o -name ".pytest_cache" \) -prune -exec rm -rf {} +

echo "[clean] Removing run-artifacts directory"
rm -rf run-artifacts || true

if [[ "$MODE" == "aggressive" ]]; then
  echo "[clean] Aggressive: pruning archived_runs to last 5 by mtime"
  mapfile -t RUN_DIRS < <(ls -td archived_runs/run-*-artifacts 2>/dev/null || true)
  if (( ${#RUN_DIRS[@]} > 5 )); then
    for ((i=5; i<${#RUN_DIRS[@]}; i++)); do
      echo "[clean] deleting ${RUN_DIRS[$i]}"
      rm -rf "${RUN_DIRS[$i]}"
    done
  else
    echo "[clean] nothing to prune in archived_runs"
  fi

  echo "[clean] Aggressive: deleting non-final outputs (*.pptx)"
  KEEP_REGEX='^(final.*|output_en_final.*)\.pptx$'
  if compgen -G "outputs/*.pptx" > /dev/null; then
    for f in outputs/*.pptx; do
      base="$(basename "$f")"
      if [[ "$base" =~ $KEEP_REGEX ]]; then
        echo "[clean] keeping outputs/$base"
      else
        echo "[clean] deleting outputs/$base"
        rm -f "$f"
      fi
    done
  fi
fi

echo "[clean] Done. Size summary:"
du -sh . 2>/dev/null || true
