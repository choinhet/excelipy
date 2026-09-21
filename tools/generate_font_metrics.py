"""
Build the metrics excelipy measures with when a font is not installed.

Sizing needs the advance of every character and the kerning between
neighbouring pairs. Those are numbers, so they can be carried rather than
measured: a box with no fonts of its own then lays a sheet out exactly as a
workstation does, without a font file in sight.

They are read off the metric-compatible substitutes, which are drawn to the
same widths as the Microsoft fonts by design and are free to redistribute -
Carlito for Calibri, Caladea for Cambria, and Liberation for Arial, Times New
Roman and Courier New. On Debian or Ubuntu:

    apt-get install fonts-crosextra-carlito fonts-crosextra-caladea fonts-liberation

Then, from the repository root:

    python tools/generate_font_metrics.py

Widths are stored in the same fraction of a pixel, at the same size, that
excelipy measures an installed font at, so a carried font and an installed one
come to exactly the same number. The file is gzipped because it is read once
and never edited by hand.
"""

import gzip
import json
import sys
from pathlib import Path

from fontTools.ttLib import TTFont
from PIL import ImageFont

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from excelipy.writers.table import MEASURE_PX, METRICS_SUBPIXEL  # noqa: E402

# The face each font is measured from. Every one of these is metrically
# identical to the font it stands for, so the numbers are that font's.
SOURCES = {
    "calibri": ("Carlito-Regular.ttf", "Carlito-Bold.ttf"),
    "cambria": ("Caladea-Regular.ttf", "Caladea-Bold.ttf"),
    "arial": ("LiberationSans-Regular.ttf", "LiberationSans-Bold.ttf"),
    "times new roman": ("LiberationSerif-Regular.ttf", "LiberationSerif-Bold.ttf"),
    "courier new": ("LiberationMono-Regular.ttf", "LiberationMono-Bold.ttf"),
}

# Characters a spreadsheet is written in: ASCII, the Latin-1 letters, and the
# punctuation a report picks up from a word processor.
CHARS = (
    [chr(code) for code in range(0x20, 0x7F)]
    + [chr(code) for code in range(0xA0, 0x100)]
    + list("‐–—‘’“”•…‰")
    + list("€™←→≠≤≥×÷")
)

# Kerning below this is a hundredth of a pixel at 11pt, which is noise
KERN_FLOOR = 2


def face_metrics(path: str) -> dict:
    """
    Advances and kerning pairs of one face, in units of the em.

    Coverage is read from the font's character map, since a face draws the
    character it has no glyph for at a width of its own - in a monospaced font
    that is every other character's width, so measuring cannot tell them apart.
    Widths are then measured through Pillow, which is what does the measuring
    when the font itself is installed, so both paths agree to the pixel.
    """
    resolved = ImageFont.truetype(path, 1).path
    covered = TTFont(resolved).getBestCmap()
    font = ImageFont.truetype(path, MEASURE_PX)
    scale = METRICS_SUBPIXEL
    present = {char: font.getlength(char) for char in CHARS if ord(char) in covered}
    dropped = sorted(set(CHARS) - set(present))
    if dropped:
        print(f"    no glyph for {''.join(dropped)!r}")

    kerning = {}
    for first, first_width in present.items():
        for second, second_width in present.items():
            pair = first + second
            kern = round((font.getlength(pair) - first_width - second_width) * scale)
            if abs(kern) >= KERN_FLOOR:
                kerning[pair] = kern

    return {
        "source": Path(path).name,
        "advances": {char: round(width * scale) for char, width in present.items()},
        "kerning": kerning,
    }


def main() -> None:
    out = Path(__file__).resolve().parent.parent / "excelipy/resources/font_metrics.json.gz"
    faces: dict[str, dict] = {}
    for family, (regular, bold) in SOURCES.items():
        faces[family] = {}
        for weight, file_name in (("regular", regular), ("bold", bold)):
            print(f"{family} {weight}: {file_name}")
            faces[family][weight] = face_metrics(file_name)

    blob = {
        "measured_at": MEASURE_PX,
        "subpixel": METRICS_SUBPIXEL,
        "note": (
            "Advance widths and kerning of the fonts named, read off the "
            "metric-compatible faces listed in each entry's source: Carlito, "
            "Caladea and Liberation, all SIL Open Font License 1.1. Built by "
            "tools/generate_font_metrics.py."
        ),
        "faces": faces,
    }
    packed = json.dumps(blob, separators=(",", ":"), ensure_ascii=False, sort_keys=True)
    out.write_bytes(gzip.compress(packed.encode("utf-8"), mtime=0))
    print(f"\n{out}: {len(packed) / 1024:.0f}KiB of metrics, {out.stat().st_size / 1024:.0f}KiB packed")


if __name__ == "__main__":
    main()
