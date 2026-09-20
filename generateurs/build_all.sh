#!/usr/bin/env bash
# Reconstruit l'ensemble des classeurs et les synchronise.
set -euo pipefail
cd "$(dirname "$0")"
RECALC="${RECALC:-/root/.claude/skills/synced/xlsx/scripts/recalc.py}"
mkdir -p ../classeurs
echo "1/4 paramétrage"
python3 build_parametrage.py
[ -f "$RECALC" ] && python3 "$RECALC" "../classeurs/AA - Parametrage 2026-2027.xlsx" 180
echo "2/4 classeurs de ligne"
python3 build_suivi.py
echo "3/4 synchronisation (recalcul inclus)"
python3 sync_referentiel.py ../classeurs
echo "4/4 contrôle"
python3 check_classeurs.py
