#!/usr/bin/env bash
# Compile le chargeur natif (32 et 64 bits) avec MinGW-w64.
# Usage : build/build-shim.sh <dossier de sortie>
# Prérequis (Debian/Ubuntu) : apt install gcc-mingw-w64-i686 gcc-mingw-w64-x86-64
set -euo pipefail
root="$(cd "$(dirname "$0")/.." && pwd)"
out="${1:?dossier de sortie}"
mkdir -p "$out"
for target in i686:32 x86_64:64; do
  arch="${target%%:*}"
  bits="${target##*:}"
  "$arch-w64-mingw32-gcc" -std=c99 -O2 -Wall -Wextra -Werror -shared -s -static-libgcc \
    -o "$out/WordTableToExcel.Shim$bits.dll" \
    "$root/src/WordTableToExcel.Shim/shim.c" "$root/src/WordTableToExcel.Shim/shim.def" \
    -Wl,--no-insert-timestamp -Wl,--enable-stdcall-fixup
  echo "OK : $out/WordTableToExcel.Shim$bits.dll"
done
