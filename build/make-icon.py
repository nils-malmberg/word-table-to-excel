#!/usr/bin/env python3
"""Génère l'icône de l'application (src/TableauxWordExcel/TableauxWordExcel.ico).

Carré arrondi coupé en diagonale (bleu Word / vert Excel) portant un tableau blanc.
Chaque taille est dessinée séparément (traits proportionnés) avec suréchantillonnage.
Usage : python3 build/make-icon.py [fichier.ico]   (nécessite Pillow)
"""
import sys
from PIL import Image, ImageDraw

BLUE = (24, 90, 189, 255)
GREEN = (16, 124, 65, 255)
WHITE = (255, 255, 255, 255)
SIZES = [16, 20, 24, 32, 40, 48, 64, 128, 256]


def draw(size):
    scale = 8
    s = size * scale
    img = Image.new("RGBA", (s, s), (0, 0, 0, 0))

    # Fond : diagonale bleu (haut gauche) / vert (bas droite), coins arrondis.
    background = Image.new("RGBA", (s, s), GREEN)
    ImageDraw.Draw(background).polygon([(0, 0), (s, 0), (0, s)], fill=BLUE)
    mask = Image.new("L", (s, s), 0)
    margin = round(s * 0.03)
    ImageDraw.Draw(mask).rounded_rectangle([margin, margin, s - margin - 1, s - margin - 1], radius=round(s * 0.2), fill=255)
    img.paste(background, (0, 0), mask)

    # Tableau blanc : cadre, ligne d'en-tête pleine, grille 3 × 3.
    d = ImageDraw.Draw(img)
    left, top, right, bottom = s * 0.2, s * 0.24, s * 0.8, s * 0.76
    line = max(scale, round(s * (0.06 if size <= 24 else 0.045)))
    header = top + (bottom - top) * 0.3
    d.rectangle([left, top, right, header], fill=WHITE)
    d.rectangle([left, top, right, bottom], outline=WHITE, width=line)
    if size >= 20:
        rows = [header + (bottom - header) / 2]
        cols = [left + (right - left) / 3, left + 2 * (right - left) / 3]
    else:
        rows = []
        cols = [left + (right - left) / 2]
    for y in rows:
        d.rectangle([left, y - line / 2, right, y + line / 2], fill=WHITE)
    for x in cols:
        d.rectangle([x - line / 2, header, x + line / 2, bottom], fill=WHITE)

    return img.resize((size, size), Image.LANCZOS)


def main():
    path = sys.argv[1] if len(sys.argv) > 1 else "src/TableauxWordExcel/TableauxWordExcel.ico"
    images = [draw(size) for size in SIZES]
    largest = images[-1]
    largest.save(path, format="ICO", sizes=[(i.width, i.height) for i in images], append_images=images[:-1])
    print("Icône écrite :", path)


if __name__ == "__main__":
    main()
