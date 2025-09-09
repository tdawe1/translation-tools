#!/usr/bin/env bash
set -euo pipefail

# Purges sensitive artifacts from Git history using git-filter-repo.
# This is destructive to history; run from a clean working tree.
# Requires: pip install git-filter-repo OR brew install git-filter-repo

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

# Run filter
git filter-repo \
  --force \
  $(printf -- "--path-glob %q " "${patterns[@]}") \
  --invert-paths

echo "History rewritten. Next steps:"
echo "  1) git push --force-with-lease origin HEAD"
echo "  2) Ask collaborators to re-clone"
echo "  3) Rotate any credentials that were ever committed"
