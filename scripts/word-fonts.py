#!/usr/bin/env python3
"""Build the static TrueType faces Word embeds.

The browser uses static WOFF2 fonts in src/fonts. Word's font embedding
takes TrueType faces, so this script decompresses those exact upstream faces
without changing their existing default axes, gives every face a valid family
of its own (Word has no CSS-like weight selection) and writes
src/fonts/word/<key>/<face>.ttf plus manifest.json.

It is a build-time tool: Python 3 with fontTools and brotli. See
src/fonts/README.md for the command.

    python3 scripts/word-fonts.py            # rebuild every face
    python3 scripts/word-fonts.py --check    # verify manifest and files
"""

from __future__ import annotations

import argparse
import hashlib
import json
import io
import sys
from pathlib import Path

try:
    from fontTools.ttLib import TTFont
except ImportError:  # pragma: no cover - a clear message beats a traceback
    sys.exit("This script needs fontTools: pip install fonttools brotli")

ROOT = Path(__file__).resolve().parent.parent
FONT_ROOT = ROOT / "src" / "fonts"
OUT_ROOT = FONT_ROOT / "word"
MANIFEST = OUT_ROOT / "manifest.json"

# Theme fonts are keyed by their ThemeFont enum value. `family` is the name
# Word shows; it avoids every Reserved Font Name in the OFL files (Plex, Lora,
# Playfair Display and Source).
THEME_FONTS = [
    {
        "key": "NEWSREADER",
        "family": "Newsreader",
        "stem": "newsreader/Newsreader",
        "weights": [400, 500, 600, 700],
    },
    {
        "key": "NEWSREADER_ITALIC",
        "family": "Newsreader",
        "stem": "newsreader/Newsreader",
        "weights": [400, 500, 600, 700],
        "style": "italic",
    },
    {
        "key": "GARAMOND",
        "family": "EB Garamond",
        "stem": "eb-garamond/EBGaramond",
        "weights": [400, 500, 600, 700],
    },
    {
        "key": "SOURCE_SERIF",
        "family": "Dossier Serif",
        "stem": "source-serif-4/SourceSerif4",
        "weights": [400, 500, 600, 700],
    },
    {
        "key": "LORA",
        "family": "Dossier Literary",
        "stem": "lora/Lora",
        "weights": [400, 500, 600, 700],
    },
    {
        "key": "LIBRE_CASLON",
        "family": "Libre Caslon Text",
        "stem": "libre-caslon-text/LibreCaslonText",
        "weights": [400, 500, 600, 700],
    },
    {
        "key": "PLAYFAIR",
        "family": "Dossier Display",
        "stem": "playfair-display/PlayfairDisplay",
        "weights": [400, 500, 600, 700],
    },
    {
        "key": "FRAUNCES",
        "family": "Fraunces",
        "stem": "fraunces/Fraunces",
        "weights": [400, 500, 600, 700],
    },
    {
        "key": "GEIST",
        "family": "Geist",
        "stem": "geist/Geist",
        "weights": [400, 500, 600, 700],
    },
    {
        "key": "PLEX_SANS",
        "family": "Dossier Grotesk",
        "stem": "ibm-plex-sans/IBMPlexSans",
        "weights": [400, 500, 600, 700],
    },
    {
        "key": "INTER",
        "family": "Inter",
        "stem": "inter/Inter",
        "weights": [400, 500, 600, 700],
    },
    {
        "key": "MANROPE",
        "family": "Manrope",
        "stem": "manrope/Manrope",
        "weights": [400, 500, 600, 700],
    },
    {
        "key": "WORK_SANS",
        "family": "Work Sans",
        "stem": "work-sans/WorkSans",
        "weights": [400, 500, 600, 700],
    },
    {
        "key": "SPACE_GROTESK",
        "family": "Space Grotesk",
        "stem": "space-grotesk/SpaceGrotesk",
        "weights": [400, 500, 600, 700],
    },
    {
        "key": "PLEX_MONO",
        "family": "Dossier Mono",
        "stem": "ibm-plex-mono/IBMPlexMono",
        "weights": [400, 500, 600],
    },
    {
        "key": "JETBRAINS_MONO",
        "family": "JetBrains Mono",
        "stem": "jetbrains-mono/JetBrainsMono",
        "weights": [400, 500, 600, 700],
    },
]

# Fallbacks for scripts the theme fonts don't cover, as in src/fonts/index.ts.
SCRIPT_FONTS = [
    {
        "key": "ARABIC",
        "family": "Dossier Arabic",
        "stem": "ibm-plex-sans-arabic/IBMPlexSansArabic",
        "weights": [400, 500, 600],
    },
    {
        "key": "DEVANAGARI",
        "family": "Dossier Devanagari",
        "stem": "ibm-plex-sans-devanagari/IBMPlexSansDevanagari",
        "weights": [400, 500, 600],
    },
    {
        "key": "JAPANESE",
        "family": "Dossier Japanese",
        "stem": "ibm-plex-sans-jp/IBMPlexSansJP",
        "weights": [400, 500, 600],
    },
    {
        "key": "KOREAN",
        "family": "Dossier Korean",
        "stem": "ibm-plex-sans-kr/IBMPlexSansKR",
        "weights": [400, 500, 600],
    },
    {
        "key": "CHINESE",
        "family": "Dossier Chinese",
        "stem": "noto-sans-sc/NotoSansSC",
        "weights": [400, 600],
    },
]

WEIGHT_SUFFIX = {400: "", 500: " Medium", 600: " SemiBold", 700: " Bold"}

# The current sources are already static instances (all non-weight axes at
# their upstream defaults). Resolve every weight to the file the browser uses;
# never restore variable sources or generate distinct text/display instances.
STYLES = {400: "Regular", 500: "Medium", 600: "SemiBold", 700: "Bold"}
for spec in THEME_FONTS + SCRIPT_FONTS:
    stem = spec["stem"]
    spec["files"] = {
        weight: f"{stem}-{'Italic' if weight == 400 else STYLES[weight] + 'Italic'}.woff2"
        if spec.get("style") == "italic" else f"{stem}-{STYLES[weight]}.woff2"
        for weight in spec["weights"]
    }


def face_name(base: str, weight: int, optical: str, style: str) -> str:
    """A family name Word can select: one family per weight and optical size."""
    name = base
    if style == "italic":
        name += " Italic"
    if optical == "display":
        name += " Display"
    return f"{name}{WEIGHT_SUFFIX[weight]}"


def set_names(font: TTFont, family: str, style: str, weight: int) -> None:
    """Rewrite the name table for a one-face family, keeping the licence fields."""
    subfamily = "Italic" if style == "italic" else "Regular"
    postscript = "".join(ch for ch in family if ch.isalnum() or ch in "-_")
    records = {
        1: family,
        2: subfamily,
        3: f"{family} {subfamily}; Dossier Word export",
        4: f"{family} {subfamily}",
        6: postscript,
        16: family,
        17: subfamily,
    }
    name = font["name"]
    name.names = [record for record in name.names if record.nameID not in records]
    for name_id, value in records.items():
        name.setName(value, name_id, 3, 1, 0x409)
        name.setName(value, name_id, 1, 0, 0)
    os2 = font["OS/2"]
    os2.usWeightClass = weight
    # Each face is the regular member of its own family. The italic bit stays,
    # so Word and editors show it as italic rather than slanting it again.
    os2.fsSelection = (os2.fsSelection & ~0b1100001) | (0b1 if style == "italic" else 0b1000000)
    font["head"].macStyle = 0b10 if style == "italic" else 0
    # Preserve the source's real italic angle; changing names must not change
    # any of the browser face's design metrics.


def instance(source: Path) -> TTFont:
    font = TTFont(source, lazy=False, recalcTimestamp=False)
    if "fvar" in font or "glyf" not in font:
        raise ValueError(f"{source} must be a static TrueType face")
    return font


def build_face(font: TTFont, out: Path | io.BytesIO) -> None:
    font.flavor = None  # TrueType in the package, not WOFF2
    font.save(out)


_DIGESTS: dict[Path, str] = {}


def digest(path: Path) -> str:
    if path not in _DIGESTS:
        _DIGESTS[path] = hashlib.sha256(path.read_bytes()).hexdigest()
    return _DIGESTS[path]


def embedding_permission(path: Path) -> tuple[int, str]:
    """OS/2.fsType and what it allows: 0 or 8 is editable embedding."""
    fs_type = int(TTFont(path, lazy=True)["OS/2"].fsType)
    if fs_type & 0x0002:
        return fs_type, "restricted"
    if fs_type & 0x0004:
        return fs_type, "preview-print"
    return fs_type, "editable"


def build(check: bool) -> int:
    faces: list[dict] = []
    expected: set[Path] = set()

    def add(key: str, family: str, source: Path, weight: int, optical: str, style: str) -> None:
        name = face_name(family, weight, optical, style)
        out = OUT_ROOT / key.lower() / f"{optical}{'-' if optical else ''}{weight}{'-italic' if style == 'italic' else ''}.ttf"
        expected.add(out)
        fs_type, embedding = embedding_permission(source)
        if embedding != "editable" or fs_type & 0x0200:
            raise ValueError(f"{source.name} does not allow editable outline embedding (fsType={fs_type})")
        font = instance(source)
        if font["OS/2"].usWeightClass != weight:
            raise ValueError(f"{source.name} is not weight {weight}")
        set_names(font, name, style, weight)
        buffer = io.BytesIO()
        build_face(font, buffer)
        data = buffer.getvalue()
        faces.append(
            {
                "key": key,
                "family": name,
                "weight": weight,
                "style": style,
                "optical": optical,
                "file": str(out.relative_to(FONT_ROOT)),
                "source": str(source.relative_to(FONT_ROOT)),
                "sourceSha256": digest(source),
                "sha256": hashlib.sha256(data).hexdigest(),
                "fsType": fs_type,
                "embedding": embedding,
            }
        )
        if check:
            if not out.exists() or out.read_bytes() != data:
                raise ValueError(f"Missing or stale generated face: {out}; rebuild scripts/word-fonts.py")
        else:
            out.parent.mkdir(parents=True, exist_ok=True)
            out.write_bytes(data)

    for spec in THEME_FONTS:
        for weight in spec["weights"]:
            source = FONT_ROOT / spec["files"][weight]
            add(spec["key"], spec["family"], source, weight, "text", spec.get("style", "normal"))

    for spec in SCRIPT_FONTS:
        for weight in spec["weights"]:
            source = FONT_ROOT / spec["files"][weight]
            add(spec["key"], spec["family"], source, weight, "text", "normal")

    expected.add(MANIFEST)
    if check:
        obsolete = set(OUT_ROOT.rglob("*.ttf")) - expected
        if obsolete:
            print("Obsolete Word faces; regenerate: " + ", ".join(str(path) for path in sorted(obsolete)))
            return 1
        missing = sorted(path for path in expected if not path.exists())
        if missing:
            print("Missing generated Word fonts:\n" + "\n".join(str(path) for path in missing))
            return 1
        current = json.loads(MANIFEST.read_text())
        if current.get("faces") != faces:
            print("manifest.json is out of date; run python3 scripts/word-fonts.py")
            return 1
        print(f"Verified {len(faces)} Word faces, source hashes and generated bytes.")
        return 0

    # Only the faces this script owns: leave anything else in word/ alone.
    for old in OUT_ROOT.rglob("*.ttf"):
        if old not in expected:
            old.unlink()
    manifest = {
        "note": "Generated by scripts/word-fonts.py from the bundled WOFF2 fonts. Do not edit.",
        "faces": faces,
    }
    MANIFEST.write_text(json.dumps(manifest, indent=2) + "\n")
    total = sum((OUT_ROOT / face["file"].removeprefix("word/")).stat().st_size for face in faces)
    print(f"Wrote {len(faces)} faces, {total / 1024 / 1024:.1f} MB, to {OUT_ROOT}")
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--check", action="store_true", help="check the generated files instead of writing them")
    return build(parser.parse_args().check)


if __name__ == "__main__":
    raise SystemExit(main())
