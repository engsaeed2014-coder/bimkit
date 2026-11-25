#!/usr/bin/env bash
# REVEAL whole-project 3D/4D/5D page in one command.
# usage: examples/reveal-all/build.sh <reveal-4d clone> <baseline master.xlsx> <actualized master.xlsx> <data date YYYY-MM-DD> [out dir]
set -euo pipefail
R4D=${1:?reveal-4d clone}; BASE=${2:?baseline master xlsx}; UPD=${3:?actualized master xlsx}; DD=${4:?data date}; OUT=${5:-out/reveal-all}
cd "$(dirname "$0")/../.."
P=examples/reveal-all/project.json
python3 adapters/reveal4d.py --reveal4d "$R4D" --master "$BASE" --area WB EB NB SB Mall --site --pits \
        --stands examples/reveal-all/pump_stands.json --out examples/reveal-all --name "REVEAL Compound"
rm -rf "$OUT"
python3 schedule.py "$P" "$OUT"
python3 adapters/reveal_progress.py --template "$OUT/REVEAL_ALL_Update.xlsx" --master "$UPD" --data-date "$DD" \
        --out examples/reveal-all/REVEAL_ALL_Update.xlsx
python3 update.py "$P" "$OUT" examples/reveal-all/REVEAL_ALL_Update.xlsx --data-date "$DD" --assume-baseline-progress
python3 cost.py "$P" "$OUT"
python3 viewer.py "$P" "$OUT"
python3 verify.py "$P" "$OUT"
echo "page: $OUT/index.html"
