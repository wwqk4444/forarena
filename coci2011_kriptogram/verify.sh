#!/usr/bin/env bash
# 编译并对拍验证 KRIPTOGRAM 的解法。
#   用法: ./verify.sh [随机组数，默认 1000]
set -e
cd "$(dirname "$0")"
ROUNDS=${1:-1000}

g++ -O2 -std=c++17 -o /tmp/krip_sol   sol_lean.cpp
g++ -O2 -std=c++17 -o /tmp/krip_brute brute.cpp

echo "== 样例 =="
for f in sample1 sample2 sample3; do
    printf "%-9s -> %s\n" "$f" "$(/tmp/krip_sol < $f.in)"
done

echo "== 随机对拍 $ROUNDS 组 =="
for s in $(seq 1 "$ROUNDS"); do
    python3 gen.py "$s" > /tmp/krip.in
    a=$(/tmp/krip_sol < /tmp/krip.in)
    b=$(/tmp/krip_brute < /tmp/krip.in)
    if [ "$a" != "$b" ]; then
        echo "对拍失败 seed=$s (解法 $a / 暴力 $b)"
        cat /tmp/krip.in
        exit 1
    fi
done
echo "全部一致 ✔"
