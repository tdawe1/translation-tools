#!/usr/bin/env bash
set -euo pipefail

# Purges sensitive artifacts from Git history.
# Prefer git-filter-repo; fallback to git filter-branch if unavailable.
# Destructive to history — run from a clean working tree.

read -r -p "This will rewrite history. Have you pushed backups? (yes/no) " ok
if [[ ${ok,,} != "yes" ]]; then
  echo "Aborting."; exit 1
fi

patterns=(
  'archived_runs/**'
  'inputs/**'
  'outputs/**'
  'audits/**'
  'backup/**'
  'comparisons/**'
  'data/**'
  'glossary.json'
  'translation_cache*.json'
  '*.csv'
  '*.log'
  'prompt*.md'
)

python3 - <<'PY'
import json, os, sys
ps = [
  'archived_runs/**','inputs/**','outputs/**','audits/**','backup/**','comparisons/**','data/**',
  'glossary.json','translation_cache*.json','*.csv','*.log','prompt*.md'
]
with open('.gitignore','a') as f:
    f.write('\n# ensured by purge_history\n'+'\n'.join(ps)+'\n')
print('Ensured patterns appended to .gitignore')
PY

# Commit .gitignore update so working tree is clean
git add .gitignore >/dev/null 2>&1 || true
git commit -m "chore(purge): ensure ignore patterns (auto)" >/dev/null 2>&1 || true

if command -v git-filter-repo >/dev/null 2>&1 || git filter-repo -h >/dev/null 2>&1; then
  echo "Using git-filter-repo"
  git filter-repo --force $(printf -- "--path-glob %q " "${patterns[@]}") --invert-paths
else
  echo "git-filter-repo not found; falling back to git filter-branch (slower)"
  export FILTER_BRANCH_SQUELCH_WARNING=1
  git branch backup/pre-purge || true
  git filter-branch --force --prune-empty --index-filter \
    "git rm -r --cached --ignore-unmatch ${patterns[*]}" \
    --tag-name-filter cat -- --all
fi

echo "History rewritten. Next steps:"
echo "  1) git push --force-with-lease origin HEAD"
echo "  2) Ask collaborators to re-clone"
echo "  3) Rotate any credentials that were ever committed"
