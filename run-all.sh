#!/bin/bash
# 跑全部测试。用法：把 index.html 和 AppsScript.gs 放在同一目录，然后 bash run-all.sh
set -u
total=0; failed=0
for t in gastest savtest bondtest e2etest test2 test4 test5 test6 uitest; do
  out=$(node "$t.mjs" 2>&1)
  p=$(echo "$out" | grep -cE '^✅')
  f=$(echo "$out" | grep -cE '^❌')
  total=$((total+p)); failed=$((failed+f))
  printf '%-12s ✅%-4s ❌%s\n' "$t" "$p" "$f"
  echo "$out" | grep -E '^❌'
done
echo "----------------------------------"
echo "通过 $total 项，失败 $failed 项"
[ "$failed" -eq 0 ] || exit 1
