#!/usr/bin/env bash
# Régénère les fichiers binaires du dossier Installation/ (réservé à la maintenance du projet :
# les utilisateurs n'ont rien à compiler, ces fichiers sont fournis dans le dépôt).
# Prérequis : SDK .NET 8, MinGW-w64 (voir build-shim.sh).
set -euo pipefail
root="$(cd "$(dirname "$0")/.." && pwd)"
out="${1:-$root/Installation}"
dotnet test "$root/tests/WordTableToExcel.Tests" -c Release
dotnet build "$root/src/WordTableToExcel" -c Release
cp "$root/src/WordTableToExcel/bin/Release/net40/WordTableToExcel.dll" "$out/"
"$root/build/build-shim.sh" "$out"
echo "Dossier prêt : $out"
