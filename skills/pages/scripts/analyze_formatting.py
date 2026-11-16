#!/usr/bin/env python3
"""
Analyze font/size/color consistency in a .docx export of an Apple Pages document.

Usage:
    python3 analyze_formatting.py <path-to-docx> [--full]

Requires python-docx (pip install python-docx) in whatever venv you run this with.

What it does:
- Walks every paragraph and run, resolving effective font/size (run override,
  falling back to the paragraph's style default).
- Reports font, size, and color usage counts overall.
- Breaks size usage down PER paragraph style, so you can see whether a given
  style (e.g. "Body Text") is ever rendered at more than one size -- that is
  the real signal of an unintentional formatting slip. A style legitimately
  having its own consistent size (e.g. "Essay Title" always at 14pt) is not
  a problem; a style whose size *varies* is.
- Flags paragraphs that mix more than one font or size within themselves.

This does NOT try to decide what's "wrong" on its own -- font/size variety
across different styles is often intentional (headings vs. body, a style used
specifically to remove a first-line indent after a heading or block quote,
etc). Read the per-style breakdown and the flagged paragraphs in context
before concluding something is a mistake.
"""
import sys
from collections import Counter, defaultdict

try:
    import docx
except ImportError:
    print("python-docx is not installed. Run: pip install python-docx", file=sys.stderr)
    print("(use a dedicated venv -- the system Python is likely externally managed)", file=sys.stderr)
    sys.exit(1)


def effective(run, para):
    font = run.font.name or (para.style.font.name if para.style else None)
    size = run.font.size or (para.style.font.size if para.style else None)
    size_pt = size.pt if size else None
    color = None
    if run.font.color and run.font.color.type is not None:
        try:
            color = str(run.font.color.rgb)
        except Exception:
            color = None
    return font, size_pt, color


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    path = sys.argv[1]
    show_full = "--full" in sys.argv[2:]

    d = docx.Document(path)

    font_counter = Counter()
    size_counter = Counter()
    color_counter = Counter()
    style_counter = Counter()
    by_style_size = defaultdict(Counter)
    mixed_paragraphs = []

    for i, para in enumerate(d.paragraphs):
        text = para.text.strip()
        if not text:
            continue
        style = para.style.name if para.style else "None"
        style_counter[style] += 1

        runs_info = []
        for run in para.runs:
            if not run.text.strip():
                continue
            font, size_pt, color = effective(run, para)
            font_counter[font] += 1
            size_counter[size_pt] += 1
            color_counter[color] += 1
            by_style_size[style][size_pt] += 1
            runs_info.append((font, size_pt, color, run.text[:30]))

        if runs_info:
            fonts = {r[0] for r in runs_info}
            sizes = {r[1] for r in runs_info}
            if len(fonts) > 1 or len(sizes) > 1:
                mixed_paragraphs.append((i, style, text[:60], runs_info))

    print(f"Document: {path}")
    print(f"Non-empty paragraphs: {sum(style_counter.values())}\n")

    print("=== PARAGRAPH STYLE USAGE ===")
    for s, c in style_counter.most_common():
        print(f"  {c:5d}  {s}")

    print("\n=== FONT USAGE (run-level) ===")
    for f, c in font_counter.most_common():
        print(f"  {c:5d}  {f}")
    if len(font_counter) > 1:
        print("  ^ more than one font in use -- check whether that's intentional")

    print("\n=== COLOR USAGE (run-level) ===")
    for col, c in color_counter.most_common():
        print(f"  {c:5d}  {col}")

    print("\n=== SIZE DISTRIBUTION PER STYLE (the real consistency check) ===")
    real_issues = []
    for style, counter in by_style_size.items():
        if len(counter) > 1:
            real_issues.append((style, counter))
    if not real_issues:
        print("  Every paragraph style renders at exactly one size. No size drift within any style.")
    else:
        for style, counter in real_issues:
            total = sum(counter.values())
            dominant, dom_count = counter.most_common(1)[0]
            print(f"\n  Style {style!r} (total runs={total}) has MULTIPLE sizes in use:")
            for size, c in counter.most_common():
                flag = "  <-- minority, investigate" if c != dom_count else ""
                print(f"     {c:5d}  {size}pt{flag}")

    if show_full:
        print(f"\n=== FULL SIZE-PER-STYLE TABLE ===")
        for style, counter in by_style_size.items():
            total = sum(counter.values())
            print(f"\n  Style: {style!r} (total runs={total})")
            for size, c in counter.most_common():
                print(f"     {c:5d}  {size}pt")

    print(f"\n=== PARAGRAPHS MIXING FONT/SIZE WITHIN THEMSELVES: {len(mixed_paragraphs)} ===")
    for i, style, text, runs_info in mixed_paragraphs[:40]:
        print(f"\n  [para {i}] style={style!r} text={text!r}")
        for font, size_pt, color, sample in runs_info:
            print(f"     font={font!r} size={size_pt} color={color} sample={sample!r}")
    if len(mixed_paragraphs) > 40:
        print(f"\n  ... and {len(mixed_paragraphs) - 40} more")


if __name__ == "__main__":
    main()
