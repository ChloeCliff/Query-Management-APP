#!/usr/bin/env python3
"""Generate PNG icons and a Windows .ico from qbox-logo.svg."""
import argparse
import os
import subprocess
import sys

def _normalize_icon_png(png_file, size, padding_ratio=0.04, alpha_threshold=28):
    """Trim transparent margins, then center with consistent padding.

    This makes the icon fill the canvas like a typical Windows app icon.
    """
    try:
        from PIL import Image
        img = Image.open(png_file).convert("RGBA")
        alpha = img.split()[-1]
        # Ignore very faint pixels so translucent decorative background doesn't
        # make the icon appear smaller than typical Windows app icons.
        mask = alpha.point(lambda p: 255 if p >= alpha_threshold else 0)
        bbox = mask.getbbox()
        if bbox:
            img = img.crop(bbox)

        inner = max(1, int(size * (1 - 2 * padding_ratio)))
        w, h = img.size
        scale = min(inner / max(1, w), inner / max(1, h))
        nw, nh = max(1, int(w * scale)), max(1, int(h * scale))
        img = img.resize((nw, nh), Image.Resampling.LANCZOS)

        out = Image.new("RGBA", (size, size), (0, 0, 0, 0))
        out.paste(img, ((size - nw) // 2, (size - nh) // 2), img)
        out.save(png_file)
        return True
    except Exception:
        return False

def generate_icons(svg_file="qbox-logo.svg", out_prefix="qbox", out_dir="."):
    """Generate PNG icons in multiple sizes."""
    sizes = [256, 128, 64, 32, 16]
    use_windows = sys.platform.startswith("win")

    if not os.path.exists(svg_file):
        print(f"✗ Source SVG not found: {svg_file}")
        return False

    os.makedirs(out_dir, exist_ok=True)

    def p(name):
        return os.path.join(out_dir, name)
    
    for size in sizes:
        png_file = p(f"{out_prefix}-icon-{size}.png")
        try:
            # First choice: cairosvg (pure Python path, no shell tool ambiguity)
            import cairosvg
            cairosvg.svg2png(url=svg_file, write_to=png_file, output_width=size, output_height=size)
            _normalize_icon_png(png_file, size)
            print(f"✓ Generated {os.path.relpath(png_file)}")
        except ImportError:
            try:
                # Fallback: ImageMagick via `magick` (on Windows, avoid plain `convert`)
                cmd = ["magick", "convert", "-density", "300", "-resize", f"{size}x{size}", svg_file, png_file] if use_windows else ["magick", "-density", "300", "-resize", f"{size}x{size}", svg_file, png_file]
                subprocess.run(cmd, check=True, capture_output=True, text=True, timeout=30)
                _normalize_icon_png(png_file, size)
                print(f"✓ Generated {os.path.relpath(png_file)}")
            except (FileNotFoundError, subprocess.CalledProcessError, subprocess.TimeoutExpired):
                print(f"✗ Could not generate {png_file} - install cairosvg or ImageMagick (magick)")
                return False
    
    try:
        from PIL import Image

        png_256 = p(f"{out_prefix}-icon-256.png")
        if not os.path.exists(png_256):
            print(f"✗ Could not create {out_prefix}.ico - {out_prefix}-icon-256.png missing")
            return False

        img = Image.open(png_256)
        ico_path = p(f"{out_prefix}.ico")
        img.save(ico_path, format="ICO", sizes=[(256, 256), (128, 128), (64, 64), (32, 32), (16, 16)])
        print(f"✓ Generated {os.path.relpath(ico_path)}")
    except ImportError:
        print("✗ Could not generate qbox.ico - install Pillow")
        return False

    print("\n✓ All icons generated successfully!")
    print("\nFiles created:")
    for size in sizes:
        print(f"  - {os.path.relpath(p(f'{out_prefix}-icon-{size}.png'))}")
    print(f"  - {os.path.relpath(p(f'{out_prefix}.ico'))}")
    return True


def _args():
    parser = argparse.ArgumentParser(description="Generate QBOX icon assets from SVG.")
    parser.add_argument("--source", default="qbox-logo.svg", help="Path to source SVG.")
    parser.add_argument("--prefix", default="qbox", help="Output file prefix (default: qbox).")
    parser.add_argument("--out-dir", default=".", help="Output directory for generated assets.")
    parser.add_argument("--all", action="store_true", help="Generate all concept variants into icon-previews/.")
    return parser.parse_args()


def generate_all_variants():
    variants = [
        ("office-q", "qbox-logo-office-q.svg"),
        ("layered-sheets", "qbox-logo-layered-sheets.svg"),
        ("q-grid", "qbox-logo-q-grid.svg"),
        ("current", "qbox-logo.svg"),
    ]
    base_out = "icon-previews"
    ok = True
    for name, src in variants:
        print(f"\n=== {name} ({src}) ===")
        out_dir = os.path.join(base_out, name)
        success = generate_icons(svg_file=src, out_prefix="qbox", out_dir=out_dir)
        ok = ok and success
    if ok:
        print("\n✓ All variants generated in icon-previews/")
    return ok

if __name__ == "__main__":
    args = _args()
    if args.all:
        success = generate_all_variants()
    else:
        success = generate_icons(svg_file=args.source, out_prefix=args.prefix, out_dir=args.out_dir)
    sys.exit(0 if success else 1)
